using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;
using Azure.Identity;



[Authorize]
public class BajaUsuarioController : Controller
{
    private readonly IConfiguration _config;
    private readonly string _ldapBase;
    private readonly string _domainName;
    private readonly string _areasOu;
    private readonly string _usersOu;
    private readonly string _bajasOu;
    private readonly string _fsServer;
    private readonly string _shareBase;
    private readonly string _quotaBase;
    private static GraphServiceClient? _graphClient = null;



    public BajaUsuarioController(IConfiguration config)
    {
        _config = config;

        // Sección Active Directory
        var ad = _config.GetSection("ActiveDirectory");
        _ldapBase = ad["BaseLdapPrefix"] + ad["DomainComponents"];
        _domainName = ad["DomainName"];
        _areasOu = ad["AreasOu"];
        _usersOu = ad["UserAndGroupsOu"];
        _bajasOu = ad["BajasOu"];

        // Sección File Server
        var fs = _config.GetSection("FsConfig");
        _fsServer = fs["ServerName"];
        _shareBase = fs["ShareBase"];
        _quotaBase = fs["QuotaPathBase"];

    }



    [HttpGet]
    public IActionResult BajaUsuario()
    {
        List<string> usuarios;
        try
        {
            using var entry = new DirectoryEntry(_ldapBase);
            using var search = new DirectorySearcher(entry)
            {
                Filter = "(objectClass=user)",
                PageSize = 1000,
                SearchScope = SearchScope.Subtree
            };
            search.PropertiesToLoad.Add("displayName");
            search.PropertiesToLoad.Add("sAMAccountName");

            usuarios = search.FindAll()
                             .Cast<System.DirectoryServices.SearchResult>()
                             .Where(r => r.Properties.Contains("displayName")
                                      && r.Properties.Contains("sAMAccountName"))
                             .Select(r => $"{r.Properties["displayName"][0]} ({r.Properties["sAMAccountName"][0]})")
                             .OrderBy(s => s)
                             .ToList();

            // Leer las acciones adicionales desde appsettings
            var additionalActions = _config
                .GetSection("UserDeactivation:AdditionalActions")
                .Get<Dictionary<string, string>>()
                ?? new Dictionary<string, string>();

            ViewBag.AdditionalActions = additionalActions;
        }
        catch
        {
            usuarios = new List<string>();
        }

        ViewBag.Usuarios = usuarios;
        return View();
    }

    [HttpPost]
    public IActionResult BajaUsuario([FromBody] Dictionary<string, object> requestData)
    {
        // 0) Inicializar GraphServiceClient
        var tenantId = _config["AzureAd:TenantId"] ?? throw new InvalidOperationException("Falta AzureAd:TenantId");
        var clientId = _config["AzureAd:ClientId"] ?? throw new InvalidOperationException("Falta AzureAd:ClientId");
        var clientSecret = _config["AzureAd:ClientSecret"] ?? throw new InvalidOperationException("Falta AzureAd:ClientSecret");
        var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        _graphClient = new GraphServiceClient(credential);


        var messages = new List<string>();
        bool userDisabled = false;
        List<string> selectedActions = new List<string>();

        // 1) Validación básica
        if (requestData == null || !requestData.ContainsKey("username"))
            return Json(new { success = false, messages = "No se proporcionó usuario.", message = "Usuario no especificado." });

        string input = requestData["username"]?.ToString();
        string username = ExtractUsername(input);
        if (string.IsNullOrEmpty(username))
            return Json(new { success = false, messages = "Formato de usuario inválido.", message = "Formato de usuario inválido." });

        // 1.1) Leer acciones seleccionadas
        if (requestData.ContainsKey("selectedActions"))
        {
            try
            {
                var arr = (JsonElement)requestData["selectedActions"];
                if (arr.ValueKind == JsonValueKind.Array)
                {
                    selectedActions = arr.EnumerateArray()
                                         .Select(e => e.GetString())
                                         .Where(s => !string.IsNullOrEmpty(s))
                                         .ToList();
                    messages.Add($"Acciones seleccionadas: {string.Join(", ", selectedActions)}");
                }
                else messages.Add("El campo 'selectedActions' no es un array.");
            }
            catch (Exception ex)
            {
                messages.Add($"Error procesando acciones: {ex.Message}");
                return Json(new { success = false, messages = string.Join("\n", messages), message = "Error procesando acciones." });
            }
        }
        else
        {
            messages.Add("No se seleccionaron acciones adicionales.");
        }

        try
        {
            // 2) Buscar y modificar usuario en AD
            using var ctx = new PrincipalContext(ContextType.Domain, _domainName);
            using var usr = UserPrincipal.FindByIdentity(ctx, username);
            if (usr == null)
                return Json(new { success = false, messages = $"Usuario '{username}' no encontrado.", message = "Usuario no encontrado en AD." });

            var de = (DirectoryEntry)usr.GetUnderlyingObject();
            string userDn = de.Properties["distinguishedName"].Value.ToString();
            string ouOriginal = userDn.Substring(userDn.IndexOf("OU="));

            // 2.1) Eliminar de grupos
            if (de.Properties.Contains("memberOf"))
            {
                var grupos = de.Properties["memberOf"]
                               .Cast<object>()
                               .Select(dn => ExtractCNFromDN(dn.ToString()))
                               .ToList();
                foreach (var grp in grupos)
                {
                    using var ge = FindGroupByName(grp);
                    if (ge != null && ge.Properties["member"].Contains(userDn))
                    {
                        ge.Properties["member"].Remove(userDn);
                        ge.CommitChanges();
                        messages.Add($"Eliminado de grupo '{grp}'.");
                    }
                }
            }
            else messages.Add("No pertenecía a ningún grupo.");

            // 3) Eliminar cuota en servidor de cuotas
            string quotaPath = Path.Combine(_quotaBase, username);
            Type qmType = Type.GetTypeFromProgID("Fsrm.FsrmQuotaManager", _fsServer);
            if (qmType != null)
            {
                dynamic qm = Activator.CreateInstance(qmType);
                try
                {
                    dynamic existing = null;
                    try { existing = qm.GetQuota(quotaPath); } catch { }
                    if (existing != null)
                    {
                        messages.Add($"[DEBUG] Eliminando cuota en {quotaPath}");
                        existing.Delete();
                        messages.Add("Cuota FSRM eliminada.");
                    }
                    else messages.Add("No había cuota que eliminar.");
                }
                finally { Marshal.ReleaseComObject(qm); }
            }
            else messages.Add("FSRM no disponible, omito cuota.");

            // 4) Eliminar carpeta personal
            string userFolder = Path.Combine(_shareBase, username);
            if (Directory.Exists(userFolder))
            {
                Directory.Delete(userFolder, true);
                messages.Add($"Carpeta eliminada: {userFolder}");
            }
            else messages.Add("Carpeta personal no encontrada.");

            // 5) Deshabilitar cuenta
            int uac = (int)de.Properties["userAccountControl"].Value;
            de.Properties["userAccountControl"].Value = uac | 0x2;
            de.CommitChanges();
            messages.Add("Usuario deshabilitado.");

            // 6) Mover a OU=Bajas
            string newOuLdap = $"{_config["ActiveDirectory:BaseLdapPrefix"]}OU={_bajasOu},OU={_areasOu},{_config["ActiveDirectory:DomainComponents"]}";
            using var ouEntry = new DirectoryEntry(newOuLdap);
            de.MoveTo(ouEntry);
            de.CommitChanges();
            messages.Add("Usuario movido a OU 'Bajas'.");

            userDisabled = true;
        }
        catch (Exception ex)
        {
            messages.Add($"Error en AD: {ex.Message}");
        }

        // 7) Envío de correos si procede
        if (userDisabled && selectedActions.Any())
        {
            foreach (var action in selectedActions)
            {
                try
                {
                    SendMailMessage(_graphClient, username, action).GetAwaiter().GetResult();
                    messages.Add($"Email para '{(action)}' enviado.");
                }
                catch (Exception ex)
                {
                    messages.Add($"Error enviando email para '{(action)}': {ex.Message}");
                }
            }
        }

        string final = userDisabled
            ? "Baja completada."
            : "No se completó la baja.";
        return Json(new { success = userDisabled, messages = string.Join("\n", messages), message = final });
    }


    private string ExtractUsername(string input)
    {
        if (string.IsNullOrEmpty(input))
            return null;

        int startIndex = input.LastIndexOf('(');
        int endIndex = input.LastIndexOf(')');
        if (startIndex >= 0 && endIndex > startIndex)
        {
            return input.Substring(startIndex + 1, endIndex - startIndex - 1).Trim();
        }
        return null;
    }

    private string ExtractCNFromDN(string dn)
    {
        if (!string.IsNullOrEmpty(dn))
        {
            int start = dn.IndexOf("CN=");
            if (start >= 0)
            {
                int end = dn.IndexOf(",", start);
                return end > start
                    ? dn.Substring(start + 3, end - start - 3)
                    : dn.Substring(start + 3);
            }
        }
        return "";
    }

    private DirectoryEntry FindGroupByName(string groupName)
    {
        try
        {
            string domainPath = "LDAP://DC=aytosa,DC=inet";
            using (DirectoryEntry rootEntry = new DirectoryEntry(domainPath))
            {
                using (DirectorySearcher searcher = new DirectorySearcher(rootEntry))
                {
                    searcher.Filter = $"(&(objectClass=group)(cn={groupName}))";
                    searcher.SearchScope = SearchScope.Subtree;
                    searcher.PropertiesToLoad.Add("distinguishedName");

                    System.DirectoryServices.SearchResult result = searcher.FindOne();
                    return result?.GetDirectoryEntry();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al buscar el grupo '{groupName}': {ex.Message}\nStackTrace: {ex.StackTrace}");
        }

        return null;
    }

    
    
    private async Task SendMailMessage(GraphServiceClient graphClient, string username, string action)
    {
       var fromEmail = _config.GetSection("SmtpSettings")["FromEmail"]
                        ?? throw new InvalidOperationException("Falta SmtpSettings:FromEmail en config");

        // 2) Construimos el mensaje igual que antes
        var body = new SendMailPostRequestBody
        {
            Message = new Message
            {
                Subject = "Baja usuario en herramienta ",
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = "Baja al usuario " + username + " en la herramienta " + action
                },
                ToRecipients = new List<Recipient>
            {
                new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = fromEmail
                    }
                }
            }
            }
        };

        // 3) Enviamos desde la cuenta que acabamos de leer
        await graphClient
            .Users[fromEmail]
            .SendMail
            .PostAsync(body);
    }

}