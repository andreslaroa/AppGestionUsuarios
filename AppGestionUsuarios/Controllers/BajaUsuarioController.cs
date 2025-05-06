using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text.Json;

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
                             .Cast<SearchResult>()
                             .Where(r => r.Properties.Contains("displayName")
                                      && r.Properties.Contains("sAMAccountName"))
                             .Select(r => $"{r.Properties["displayName"][0]} ({r.Properties["sAMAccountName"][0]})")
                             .OrderBy(s => s)
                             .ToList();
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

            // 3) Eliminar cuota en Leonardo
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
        //if (userDisabled && selectedActions.Any())
        //{
        //    try
        //    {
        //        SendEmailToAdmin(username, /*…*/);
        //        messages.Add("Email administrador enviado.");
        //    }
        //    catch (Exception ex)
        //    {
        //        messages.Add($"Error envío email admin: {ex.Message}");
        //    }
        //    foreach (var action in selectedActions)
        //    {
        //        try
        //        {
        //            SendEmailForAction(action, /*…*/);
        //            messages.Add($"Email para acción {action} enviado.");
        //        }
        //        catch (Exception ex)
        //        {
        //            messages.Add($"Error email acción {action}: {ex.Message}");
        //        }
        //    }
        //}

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

                    SearchResult result = searcher.FindOne();
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

    //private void SendEmailToAdmin(string username, string nombreCompleto, string dni, string ouOriginal, DateTime fechaBaja, List<string> processMessages)
    //{
    //    // Validar la configuración SMTP
    //    string smtpServer = _configuration["SmtpSettings:Server"];
    //    string smtpPortStr = _configuration["SmtpSettings:Port"];
    //    string smtpUsername = _configuration["SmtpSettings:Username"];
    //    string smtpPassword = _configuration["SmtpSettings:Password"];
    //    string fromEmail = _configuration["SmtpSettings:FromEmail"];
    //    string adminEmail = _configuration["SmtpSettings:AdminEmail"];

    //    // Validar que todas las configuraciones estén presentes
    //    if (string.IsNullOrEmpty(smtpServer) || string.IsNullOrEmpty(smtpPortStr) ||
    //        string.IsNullOrEmpty(smtpUsername) || string.IsNullOrEmpty(smtpPassword) ||
    //        string.IsNullOrEmpty(fromEmail) || string.IsNullOrEmpty(adminEmail))
    //    {
    //        throw new Exception("Configuración SMTP incompleta en appsettings.json. Verifique las claves SmtpSettings.");
    //    }

    //    if (!int.TryParse(smtpPortStr, out int smtpPort))
    //    {
    //        throw new Exception("El puerto SMTP en appsettings.json no es un número válido.");
    //    }

    //    using (var client = new SmtpClient(smtpServer, smtpPort))
    //    {
    //        client.EnableSsl = true; // Habilitar TLS
    //        client.Credentials = new System.Net.NetworkCredential(smtpUsername, smtpPassword);

    //        using (var mailMessage = new MailMessage())
    //        {
    //            mailMessage.From = new MailAddress(fromEmail, "Sistema de Gestión de Usuarios");
    //            mailMessage.To.Add(adminEmail);
    //            mailMessage.Subject = $"Notificación: Baja de usuario '{username}'";

    //            // Construir el cuerpo del correo
    //            string body = "Se ha dado de baja a un usuario en el sistema. A continuación, los detalles:\n\n";
    //            body += $"Username: {username}\n";
    //            body += $"Nombre completo: {nombreCompleto}\n";
    //            body += $"DNI: {dni}\n";
    //            body += $"OU original: {ouOriginal}\n";
    //            body += $"Fecha y hora de la baja: {fechaBaja:dd/MM/yyyy HH:mm:ss}\n";
    //            body += "\nDetalles del proceso:\n- " + (processMessages.Any() ? string.Join("\n- ", processMessages) : "No se registraron eventos adicionales.");

    //            mailMessage.Body = body;
    //            mailMessage.IsBodyHtml = false; // Texto plano

    //            client.Send(mailMessage);
    //        }
    //    }
    //}

    //private void SendEmailForAction(string action, string username, string nombreCompleto, string dni, string ouOriginal, DateTime fechaBaja)
    //{
    //    // Validar la configuración SMTP
    //    string smtpServer = _configuration["SmtpSettings:Server"];
    //    string smtpPortStr = _configuration["SmtpSettings:Port"];
    //    string smtpUsername = _configuration["SmtpSettings:Username"];
    //    string smtpPassword = _configuration["SmtpSettings:Password"];
    //    string fromEmail = _configuration["SmtpSettings:FromEmail"];
    //    string adminEmail = _configuration["SmtpSettings:AdminEmail"];

    //    // Validar que todas las configuraciones estén presentes
    //    if (string.IsNullOrEmpty(smtpServer) || string.IsNullOrEmpty(smtpPortStr) ||
    //        string.IsNullOrEmpty(smtpUsername) || string.IsNullOrEmpty(smtpPassword) ||
    //        string.IsNullOrEmpty(fromEmail) || string.IsNullOrEmpty(adminEmail))
    //    {
    //        throw new Exception("Configuración SMTP incompleta en appsettings.json. Verifique las claves SmtpSettings.");
    //    }

    //    if (!int.TryParse(smtpPortStr, out int smtpPort))
    //    {
    //        throw new Exception("El puerto SMTP en appsettings.json no es un número válido.");
    //    }

    //    using (var client = new SmtpClient(smtpServer, smtpPort))
    //    {
    //        client.EnableSsl = true; // Habilitar TLS
    //        client.Credentials = new System.Net.NetworkCredential(smtpUsername, smtpPassword);

    //        using (var mailMessage = new MailMessage())
    //        {
    //            mailMessage.From = new MailAddress(fromEmail, "Sistema de Gestión de Usuarios");
    //            mailMessage.To.Add(adminEmail);
    //            mailMessage.Subject = $"Incidencia: {GetActionDescription(action)} para el usuario '{username}'";

    //            // Construir el cuerpo del correo
    //            string body = $"Se ha solicitado una acción adicional para un usuario dado de baja. A continuación, los detalles:\n\n";
    //            body += $"Acción solicitada: {GetActionDescription(action)}\n";
    //            body += $"Username: {username}\n";
    //            body += $"Nombre completo: {nombreCompleto}\n";
    //            body += $"DNI: {dni}\n";
    //            body += $"OU original: {ouOriginal}\n";
    //            body += $"Fecha y hora de la solicitud: {fechaBaja:dd/MM/yyyy HH:mm:ss}\n";
    //            body += "\nPor favor, procese esta incidencia según corresponda.";

    //            mailMessage.Body = body;
    //            mailMessage.IsBodyHtml = false; // Texto plano

    //            client.Send(mailMessage);
    //        }
    //    }
    //}

    private string GetActionDescription(string action)
    {
        return action switch
        {
            "BajaARU" => "Baja de usuario en ARU",
            "BajaEquipamientoPuesto" => "Baja de equipamiento de puesto de trabajo del usuario",
            "BajaEquipamientoMovil" => "Baja de equipamiento móvil",
            "BajaMytao" => "Baja de usuario en Mytao",
            "BajaPlataformasEstado" => "Baja de usuario en Plataformas del Estado",
            "RevocacionCertificados" => "Revocación de certificados electrónicos municipales",
            _ => "Acción desconocida"
        };
    }
}