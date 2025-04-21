using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Net.Mail;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Configuration;

[Authorize]
public class BajaUsuarioController : Controller
{
    private readonly IConfiguration _configuration;

    public BajaUsuarioController(IConfiguration configuration)
    {
        _configuration = configuration;
    }

    [HttpGet]
    public IActionResult BajaUsuario()
    {
        try
        {
            int pageSize = 1000;
            using (DirectoryEntry entry = new DirectoryEntry("LDAP://DC=aytosa,DC=inet"))
            using (DirectorySearcher searcher = new DirectorySearcher(entry))
            {
                searcher.Filter = "(objectClass=user)";
                searcher.PageSize = pageSize;
                searcher.PropertiesToLoad.Add("displayName");
                searcher.PropertiesToLoad.Add("sAMAccountName");
                searcher.SearchScope = SearchScope.Subtree;

                List<string> usuarios = new List<string>();
                foreach (SearchResult result in searcher.FindAll())
                {
                    if (result.Properties.Contains("displayName") && result.Properties.Contains("sAMAccountName"))
                    {
                        string displayName = result.Properties["displayName"][0].ToString();
                        string samAccountName = result.Properties["sAMAccountName"][0].ToString();
                        usuarios.Add($"{displayName} ({samAccountName})");
                    }
                }

                ViewBag.Usuarios = usuarios.OrderBy(u => u).ToList();
            }
        }
        catch (Exception ex)
        {
            ViewBag.Usuarios = new List<string>();
            Console.WriteLine($"Error al cargar los usuarios: {ex.Message}\nStackTrace: {ex.StackTrace}");
        }

        return View();
    }

    [HttpPost]
    public IActionResult BajaUsuario([FromBody] Dictionary<string, object> requestData)
    {
        // Lista para almacenar los mensajes del proceso
        List<string> messages = new List<string>();
        bool userDisabled = false;

        try
        {
            // Validar la solicitud
            if (requestData == null || !requestData.ContainsKey("username"))
            {
                messages.Add("Error: Solicitud inválida. No se proporcionó el nombre de usuario.");
                return Json(new { success = false, messages = string.Join("\n", messages), message = "Usuario no especificado." });
            }

            string input = requestData["username"]?.ToString();
            string username = ExtractUsername(input);

            if (string.IsNullOrEmpty(username))
            {
                messages.Add("Error: El formato del usuario seleccionado no es válido.");
                return Json(new { success = false, messages = string.Join("\n", messages), message = "El formato del usuario seleccionado no es válido." });
            }

            // Obtener las acciones seleccionadas (si no hay, será una lista vacía)
            List<string> selectedActions = new List<string>();
            if (requestData.ContainsKey("selectedActions"))
            {
                try
                {
                    var selectedActionsElement = (System.Text.Json.JsonElement)requestData["selectedActions"];
                    if (selectedActionsElement.ValueKind == System.Text.Json.JsonValueKind.Array)
                    {
                        selectedActions = selectedActionsElement.EnumerateArray()
                            .Select(action => action.GetString())
                            .Where(action => !string.IsNullOrEmpty(action))
                            .ToList();
                        messages.Add($"Acciones seleccionadas: {string.Join(", ", selectedActions)}.");
                    }
                    else
                    {
                        messages.Add("El campo 'selectedActions' no es un array válido.");
                    }
                }
                catch (Exception ex)
                {
                    messages.Add($"Error al procesar las acciones seleccionadas: {ex.Message}\nStackTrace: {ex.StackTrace}");
                    return Json(new { success = false, messages = string.Join("\n", messages), message = "Error al procesar las acciones seleccionadas." });
                }
            }
            else
            {
                messages.Add("No se seleccionaron acciones adicionales.");
            }

            // Variables para el correo
            string nombreCompleto = "";
            string dni = "";
            string ouOriginal = "";
            DateTime fechaBaja = DateTime.Now;

            // Procesar la baja del usuario en Active Directory
            try
            {
                using (var context = new PrincipalContext(ContextType.Domain, "aytosa.inet"))
                {
                    using (var user = UserPrincipal.FindByIdentity(context, username))
                    {
                        if (user == null)
                        {
                            messages.Add($"Error: El usuario '{username}' no fue encontrado en Active Directory.");
                            return Json(new { success = false, messages = string.Join("\n", messages), message = "Usuario no encontrado en Active Directory." });
                        }

                        using (var userEntry = (DirectoryEntry)user.GetUnderlyingObject())
                        {
                            string userDN = userEntry.Properties["distinguishedName"].Value.ToString();
                            ouOriginal = userDN.Substring(userDN.IndexOf("OU=")); // Guardar la OU original

                            // Obtener datos del usuario para el correo
                            nombreCompleto = userEntry.Properties.Contains("displayName") ? userEntry.Properties["displayName"].Value?.ToString() : "N/A";
                            dni = userEntry.Properties.Contains("description") ? userEntry.Properties["description"].Value?.ToString() : "N/A";

                            // 1. Eliminar al usuario de todos los grupos
                            try
                            {
                                if (userEntry.Properties.Contains("memberOf"))
                                {
                                    List<string> grupos = new List<string>();
                                    foreach (var groupDN in userEntry.Properties["memberOf"])
                                    {
                                        string groupCN = ExtractCNFromDN(groupDN.ToString());
                                        grupos.Add(groupCN);
                                    }
                                    foreach (string groupCN in grupos)
                                    {
                                        DirectoryEntry groupEntry = FindGroupByName(groupCN);
                                        if (groupEntry != null)
                                        {
                                            try
                                            {
                                                if (groupEntry.Properties["member"].Contains(userDN))
                                                {
                                                    groupEntry.Properties["member"].Remove(userDN);
                                                    groupEntry.CommitChanges();
                                                    messages.Add($"Usuario eliminado del grupo '{groupCN}' correctamente.");
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                messages.Add($"Error al eliminar usuario del grupo '{groupCN}': {ex.Message}\nStackTrace: {ex.StackTrace}. Continuando con el proceso...");
                                            }
                                            finally
                                            {
                                                groupEntry?.Dispose();
                                            }
                                        }
                                        else
                                        {
                                            messages.Add($"Grupo '{groupCN}' no encontrado para eliminar al usuario.");
                                        }
                                    }
                                }
                                else
                                {
                                    messages.Add("El usuario no pertenece a ningún grupo.");
                                }
                            }
                            catch (Exception ex)
                            {
                                messages.Add($"Error al eliminar usuario de grupos: {ex.Message}\nStackTrace: {ex.StackTrace}. Continuando con el proceso...");
                            }

                            // 2. Eliminar la carpeta personal del usuario (si existe)
                            string userFolderPath = $"\\\\fs1.aytosa.inet\\home\\{username}";
                            try
                            {
                                if (Directory.Exists(userFolderPath))
                                {
                                    Directory.Delete(userFolderPath, true);
                                    messages.Add($"Carpeta de usuario '{userFolderPath}' eliminada correctamente.");
                                }
                                else
                                {
                                    messages.Add($"Carpeta de usuario '{userFolderPath}' no encontrada, no se eliminó.");
                                }
                            }
                            catch (Exception ex)
                            {
                                messages.Add($"Error al eliminar carpeta personal del usuario '{username}': {ex.Message}\nStackTrace: {ex.StackTrace}. Continuando con el proceso...");
                            }

                            // 3. Eliminar la cuota FSRM asociada
                            string quotaPath = $"G:\\home\\{username}";
                            string serverName = "ribera";
                            try
                            {
                                Type fsrmQuotaManagerType = Type.GetTypeFromProgID("Fsrm.FsrmQuotaManager", serverName);
                                if (fsrmQuotaManagerType == null)
                                {
                                    messages.Add($"No se pudo crear una instancia de FsrmQuotaManager en {serverName}. Continuando con el proceso...");
                                }
                                else
                                {
                                    dynamic quotaManager = Activator.CreateInstance(fsrmQuotaManagerType);
                                    try
                                    {
                                        dynamic existingQuota = null;
                                        try
                                        {
                                            existingQuota = quotaManager.GetQuota(quotaPath);
                                        }
                                        catch { /* Ignorar si no existe */ }

                                        if (existingQuota != null)
                                        {
                                            quotaManager.DeleteQuota(quotaPath);
                                            messages.Add($"Cuota FSRM eliminada para '{quotaPath}'.");
                                        }
                                        else
                                        {
                                            messages.Add($"No se encontró cuota FSRM para '{quotaPath}'.");
                                        }
                                    }
                                    finally
                                    {
                                        if (quotaManager != null)
                                        {
                                            Marshal.ReleaseComObject(quotaManager);
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                messages.Add($"Error al eliminar la cuota FSRM para '{username}': {ex.Message}\nStackTrace: {ex.StackTrace}. Continuando con el proceso...");
                            }

                            // 4. Deshabilitar al usuario
                            try
                            {
                                int userAccountControl = (int)userEntry.Properties["userAccountControl"].Value;
                                userAccountControl |= 0x2; // Establecer el bit ACCOUNT_DISABLED
                                userEntry.Properties["userAccountControl"].Value = userAccountControl;
                                userEntry.CommitChanges();
                                messages.Add($"Usuario '{username}' deshabilitado correctamente.");
                            }
                            catch (Exception ex)
                            {
                                messages.Add($"Error al deshabilitar el usuario '{username}': {ex.Message}\nStackTrace: {ex.StackTrace}.");
                                return Json(new { success = false, messages = string.Join("\n", messages), message = $"Error al deshabilitar el usuario: {ex.Message}" });
                            }

                            // 5. Mover al usuario a la OU "Bajas" dentro de "AREAS"
                            try
                            {
                                string newOUPath = "LDAP://OU=Bajas,OU=AREAS,DC=aytosa,DC=inet";
                                using (DirectoryEntry newOUEntry = new DirectoryEntry(newOUPath))
                                {
                                    userEntry.MoveTo(newOUEntry);
                                    userEntry.CommitChanges();
                                    userDisabled = true; // Marcamos que el usuario fue deshabilitado y movido
                                    messages.Add($"Usuario '{username}' movido a la OU 'Bajas' correctamente.");
                                }
                            }
                            catch (Exception ex)
                            {
                                messages.Add($"Error al mover el usuario '{username}' a la OU 'Bajas': {ex.Message}\nStackTrace: {ex.StackTrace}.");
                                return Json(new { success = false, messages = string.Join("\n", messages), message = $"Error al mover el usuario a la OU Bajas: {ex.Message}" });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                messages.Add($"Error al procesar la baja del usuario en Active Directory: {ex.Message}\nStackTrace: {ex.StackTrace}");
                return Json(new { success = false, messages = string.Join("\n", messages), message = "Error al procesar la baja del usuario en Active Directory." });
            }

            // Enviar correos solo si hay acciones seleccionadas (checkboxes marcados)
            if (userDisabled && selectedActions.Any())
            {
                try
                {
                    SendEmailToAdmin(username, nombreCompleto, dni, ouOriginal, fechaBaja, messages);
                    messages.Add("Correo de notificación enviado al administrador.");
                }
                catch (Exception ex)
                {
                    messages.Add($"Error al enviar el correo de notificación al administrador: {ex.Message}\nStackTrace: {ex.StackTrace}");
                }

                foreach (var action in selectedActions)
                {
                    try
                    {
                        SendEmailForAction(action, username, nombreCompleto, dni, ouOriginal, fechaBaja);
                        messages.Add($"Incidencia generada: Correo enviado para '{GetActionDescription(action)}'.");
                    }
                    catch (Exception ex)
                    {
                        messages.Add($"Error al enviar el correo para la acción '{GetActionDescription(action)}': {ex.Message}\nStackTrace: {ex.StackTrace}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            messages.Add($"Error inesperado en el proceso de baja del usuario: {ex.Message}\nStackTrace: {ex.StackTrace}");
        }

        // Construir el mensaje de respuesta
        string finalMessage = userDisabled
            ? "Usuario deshabilitado y movido a la OU 'Bajas' correctamente."
            : "No se pudo deshabilitar ni mover el usuario.";
        finalMessage += "\nDetalles del proceso:\n- " + (messages.Any() ? string.Join("\n- ", messages) : "No se registraron eventos adicionales.");

        return Json(new { success = userDisabled, messages = string.Join("\n", messages), message = finalMessage });
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

    private void SendEmailToAdmin(string username, string nombreCompleto, string dni, string ouOriginal, DateTime fechaBaja, List<string> processMessages)
    {
        // Validar la configuración SMTP
        string smtpServer = _configuration["SmtpSettings:Server"];
        string smtpPortStr = _configuration["SmtpSettings:Port"];
        string smtpUsername = _configuration["SmtpSettings:Username"];
        string smtpPassword = _configuration["SmtpSettings:Password"];
        string fromEmail = _configuration["SmtpSettings:FromEmail"];
        string adminEmail = _configuration["SmtpSettings:AdminEmail"];

        // Validar que todas las configuraciones estén presentes
        if (string.IsNullOrEmpty(smtpServer) || string.IsNullOrEmpty(smtpPortStr) ||
            string.IsNullOrEmpty(smtpUsername) || string.IsNullOrEmpty(smtpPassword) ||
            string.IsNullOrEmpty(fromEmail) || string.IsNullOrEmpty(adminEmail))
        {
            throw new Exception("Configuración SMTP incompleta en appsettings.json. Verifique las claves SmtpSettings.");
        }

        if (!int.TryParse(smtpPortStr, out int smtpPort))
        {
            throw new Exception("El puerto SMTP en appsettings.json no es un número válido.");
        }

        using (var client = new SmtpClient(smtpServer, smtpPort))
        {
            client.EnableSsl = true; // Habilitar TLS
            client.Credentials = new System.Net.NetworkCredential(smtpUsername, smtpPassword);

            using (var mailMessage = new MailMessage())
            {
                mailMessage.From = new MailAddress(fromEmail, "Sistema de Gestión de Usuarios");
                mailMessage.To.Add(adminEmail);
                mailMessage.Subject = $"Notificación: Baja de usuario '{username}'";

                // Construir el cuerpo del correo
                string body = "Se ha dado de baja a un usuario en el sistema. A continuación, los detalles:\n\n";
                body += $"Username: {username}\n";
                body += $"Nombre completo: {nombreCompleto}\n";
                body += $"DNI: {dni}\n";
                body += $"OU original: {ouOriginal}\n";
                body += $"Fecha y hora de la baja: {fechaBaja:dd/MM/yyyy HH:mm:ss}\n";
                body += "\nDetalles del proceso:\n- " + (processMessages.Any() ? string.Join("\n- ", processMessages) : "No se registraron eventos adicionales.");

                mailMessage.Body = body;
                mailMessage.IsBodyHtml = false; // Texto plano

                client.Send(mailMessage);
            }
        }
    }

    private void SendEmailForAction(string action, string username, string nombreCompleto, string dni, string ouOriginal, DateTime fechaBaja)
    {
        // Validar la configuración SMTP
        string smtpServer = _configuration["SmtpSettings:Server"];
        string smtpPortStr = _configuration["SmtpSettings:Port"];
        string smtpUsername = _configuration["SmtpSettings:Username"];
        string smtpPassword = _configuration["SmtpSettings:Password"];
        string fromEmail = _configuration["SmtpSettings:FromEmail"];
        string adminEmail = _configuration["SmtpSettings:AdminEmail"];

        // Validar que todas las configuraciones estén presentes
        if (string.IsNullOrEmpty(smtpServer) || string.IsNullOrEmpty(smtpPortStr) ||
            string.IsNullOrEmpty(smtpUsername) || string.IsNullOrEmpty(smtpPassword) ||
            string.IsNullOrEmpty(fromEmail) || string.IsNullOrEmpty(adminEmail))
        {
            throw new Exception("Configuración SMTP incompleta en appsettings.json. Verifique las claves SmtpSettings.");
        }

        if (!int.TryParse(smtpPortStr, out int smtpPort))
        {
            throw new Exception("El puerto SMTP en appsettings.json no es un número válido.");
        }

        using (var client = new SmtpClient(smtpServer, smtpPort))
        {
            client.EnableSsl = true; // Habilitar TLS
            client.Credentials = new System.Net.NetworkCredential(smtpUsername, smtpPassword);

            using (var mailMessage = new MailMessage())
            {
                mailMessage.From = new MailAddress(fromEmail, "Sistema de Gestión de Usuarios");
                mailMessage.To.Add(adminEmail);
                mailMessage.Subject = $"Incidencia: {GetActionDescription(action)} para el usuario '{username}'";

                // Construir el cuerpo del correo
                string body = $"Se ha solicitado una acción adicional para un usuario dado de baja. A continuación, los detalles:\n\n";
                body += $"Acción solicitada: {GetActionDescription(action)}\n";
                body += $"Username: {username}\n";
                body += $"Nombre completo: {nombreCompleto}\n";
                body += $"DNI: {dni}\n";
                body += $"OU original: {ouOriginal}\n";
                body += $"Fecha y hora de la solicitud: {fechaBaja:dd/MM/yyyy HH:mm:ss}\n";
                body += "\nPor favor, procese esta incidencia según corresponda.";

                mailMessage.Body = body;
                mailMessage.IsBodyHtml = false; // Texto plano

                client.Send(mailMessage);
            }
        }
    }

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