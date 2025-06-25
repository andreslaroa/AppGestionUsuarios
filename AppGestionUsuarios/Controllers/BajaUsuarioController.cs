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
using Microsoft.Win32.SafeHandles;
using System.ComponentModel;
using System.Security.Principal;
using Microsoft.AspNetCore.DataProtection;



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

    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    static extern bool LogonUser(
        string lpszUsername,
        string lpszDomain,
        string lpszPassword,
        int dwLogonType,
        int dwLogonProvider,
        out IntPtr phToken);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    static extern bool CloseHandle(IntPtr handle);

    const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
    const int LOGON32_PROVIDER_DEFAULT = 0;

    private readonly IDataProtector _protector;

    public BajaUsuarioController(IConfiguration config, IDataProtectionProvider dp)
    {
        _config = config;
        _protector = dp.CreateProtector("CredencialesProtector");
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
    [Produces("application/json")]
    public IActionResult BajaUsuario([FromBody] Dictionary<string, object> requestData)
    {
        // 1) Recuperar credenciales de sesión y dominio
        string adminUsername = HttpContext.Session.GetString("adminUser");
        var encryptedPass = HttpContext.Session.GetString("adminPassword");
        var adminPassword = _protector.Unprotect(encryptedPass);

        var domainName = _config["ActiveDirectory:DomainName"];
        if (string.IsNullOrWhiteSpace(domainName))
        {
            return Json(new { success = false, message = "Configuración incorrecta: falta ActiveDirectory:DomainName" });
        }

        // 2) Impersonación
        if (!LogonUser(
                adminUsername,
                domainName,
                adminPassword,
                LOGON32_LOGON_NEW_CREDENTIALS,
                LOGON32_PROVIDER_DEFAULT,
                out var userToken))
        {
            var err = new Win32Exception(Marshal.GetLastWin32Error()).Message;
            return Json(new { success = false, message = $"Imposible impersonar: {err}" });
        }

        using var safeToken = new SafeAccessTokenHandle(userToken);

        IActionResult finalResult = Json(new { success = false, message = "No se completó la baja de usuario." });

        // 3) Ejecutar bajo las credenciales impersonadas
        WindowsIdentity.RunImpersonated(safeToken, () =>
        {
            var messages = new List<string>();
            bool userDisabled = false;
            var selectedActions = new List<string>();

            // 3.1) Validación básica
            if (requestData == null || !requestData.TryGetValue("username", out var rawUser))
            {
                finalResult = Json(new { success = false, messages = "No se proporcionó usuario.", message = "Usuario no especificado." });
                return;
            }

            string username = ExtractUsername(rawUser?.ToString());
            if (string.IsNullOrEmpty(username))
            {
                finalResult = Json(new { success = false, messages = "Formato de usuario inválido.", message = "Formato de usuario inválido." });
                return;
            }

            // 3.2) Leer acciones seleccionadas
            if (requestData.TryGetValue("selectedActions", out var rawActions))
            {
                try
                {
                    var arr = (JsonElement)rawActions;
                    if (arr.ValueKind == JsonValueKind.Array)
                        selectedActions = arr.EnumerateArray()
                                             .Select(e => e.GetString())
                                             .Where(s => !string.IsNullOrEmpty(s))
                                             .ToList();
                    messages.Add($"Acciones seleccionadas: {string.Join(", ", selectedActions)}");
                }
                catch (Exception ex)
                {
                    finalResult = Json(new { success = false, messages = $"Error procesando acciones: {ex.Message}", message = "Error procesando acciones." });
                    return;
                }
            }
            else
            {
                messages.Add("No se seleccionaron acciones adicionales.");
            }

            // 3.3) Lógica de baja en AD y recursos
            try
            {
                using var ctx = new PrincipalContext(ContextType.Domain, domainName);
                using var usr = UserPrincipal.FindByIdentity(ctx, username);
                if (usr == null)
                {
                    finalResult = Json(new { success = false, messages = $"Usuario '{username}' no encontrado.", message = "Usuario no encontrado en AD." });
                    return;
                }

                var de = (DirectoryEntry)usr.GetUnderlyingObject();
                string userDn = de.Properties["distinguishedName"].Value.ToString();

                // Quitar de grupos
                if (de.Properties.Contains("memberOf"))
                {
                    foreach (var dn in de.Properties["memberOf"].Cast<object>())
                    {
                        var grpName = ExtractCNFromDN(dn.ToString());
                        using var ge = FindGroupByName(grpName);
                        if (ge != null && ge.Properties["member"].Contains(userDn))
                        {
                            ge.Properties["member"].Remove(userDn);
                            ge.CommitChanges();
                            messages.Add($"Eliminado de grupo '{grpName}'.");
                        }
                    }
                }
                else messages.Add("No pertenecía a ningún grupo.");

                // Eliminar cuota FSRM
                try
                {
                    string quotaPath = Path.Combine(_quotaBase, username);
                    var qmType = Type.GetTypeFromProgID("Fsrm.FsrmQuotaManager", _fsServer);
                    if (qmType != null)
                    {
                        dynamic qm = Activator.CreateInstance(qmType);
                        try
                        {
                            dynamic existing = null;
                            try { existing = qm.GetQuota(quotaPath); } catch { }
                            if (existing != null)
                            {
                                existing.Delete();
                                messages.Add("Cuota FSRM eliminada.");
                            }
                            else messages.Add("No había cuota que eliminar.");
                        }
                        finally { Marshal.ReleaseComObject(qm); }
                    }
                    else messages.Add("FSRM no disponible, omito cuota.");
                }
                catch (Exception exFsrm)
                {
                    messages.Add($"Error al eliminar cuota FSRM: {exFsrm.Message}");
                }
                // Eliminar carpeta personal
                string userFolder = Path.Combine(_shareBase, username);
                if (Directory.Exists(userFolder))
                {
                    Directory.Delete(userFolder, true);
                    messages.Add($"Carpeta eliminada: {userFolder}");
                }
                else messages.Add("Carpeta personal no encontrada.");

                // Deshabilitar cuenta
                int uac = (int)de.Properties["userAccountControl"].Value;
                de.Properties["userAccountControl"].Value = uac | 0x2;
                de.CommitChanges();
                messages.Add("Usuario deshabilitado.");

                // Mover a OU=Bajas
                string newOu = $"{_config["ActiveDirectory:BaseLdapPrefix"]}OU={_bajasOu},OU={_areasOu},{_config["ActiveDirectory:DomainComponents"]}";
                using var ouEntry = new DirectoryEntry(newOu);
                de.MoveTo(ouEntry);
                de.CommitChanges();
                messages.Add("Usuario movido a OU 'Bajas'.");

                userDisabled = true;
            }
            catch (Exception exAd)
            {
                messages.Add($"Error en AD: {exAd.Message}");
            }

            // 3.4) Envío de correos si procede
            if (userDisabled && selectedActions.Any())
            {
                foreach (var action in selectedActions)
                {
                    try
                    {
                        SendMailMessage(_graphClient, username, action).GetAwaiter().GetResult();
                        messages.Add($"Email para '{action}' enviado.");
                    }
                    catch (Exception exMail)
                    {
                        messages.Add($"Error enviando email para '{action}': {exMail.Message}");
                    }
                }
            }

            // 3.5) Preparar resultado final
            string finalMsg = userDisabled ? "Baja completada." : "No se completó la baja.";
            finalResult = Json(new { success = userDisabled, messages = string.Join("\n", messages), message = finalMsg });
        });

        // 4) Cerrar handle
        CloseHandle(userToken);

        // 5) Devolver resultado JSON
        return finalResult;
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