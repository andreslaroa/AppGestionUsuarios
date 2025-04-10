using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.IO;
using System.DirectoryServices.AccountManagement;
using System;
using System.DirectoryServices;
using System.Globalization;
using System.Linq;
using Microsoft.AspNetCore.Authorization;
using System.Collections.ObjectModel;
using System.Management.Automation;

[Authorize]
public class GestionUsuariosController : Controller
{
    private readonly ILogger<GestionUsuariosController> _logger;

    public class UserModelAltaUsuario
    {
        public string Nombre { get; set; }
        public string Apellido1 { get; set; }
        public string Apellido2 { get; set; }
        public string NTelefono { get; set; }
        public string Username { get; set; }
        public string NFuncionario { get; set; }
        public string OUPrincipal { get; set; }
        public string OUSecundaria { get; set; }
        public string Departamento { get; set; }
        public string FechaCaducidadOp { get; set; }
        public DateTime FechaCaducidad { get; set; }
        public string Cuota { get; set; }
        public List<string> Grupos { get; set; }
    }

    public class userInputModel
    {
        public string Nombre { set; get; }
        public string Apellido1 { set; get; }
        public string Apellido2 { set; get; }
    }

    [HttpGet]
    public IActionResult HabilitarDeshabilitarUsuario()
    {
        try
        {
            int pageSize = 1000;
            using (var entry = new DirectoryEntry("LDAP://DC=aytosa,DC=inet"))
            {
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = "(objectClass=user)";
                    searcher.PropertiesToLoad.Add("displayName");
                    searcher.PropertiesToLoad.Add("sAMAccountName");
                    searcher.SearchScope = SearchScope.Subtree;
                    searcher.PageSize = pageSize;

                    var usuarios = new List<string>();
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
        }
        catch (Exception ex)
        {
            ViewBag.Usuarios = new List<string>();
            Console.WriteLine($"Error al cargar los usuarios: {ex.Message}");
        }

        return View();
    }

    [HttpPost]
    public IActionResult CheckUsernameExists([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData != null && requestData.ContainsKey("username"))
        {
            string username = requestData["username"];
            bool exists = CheckUserInActiveDirectory(username);
            return Json(exists);
        }

        return Json(false);
    }

    private bool CheckUserInActiveDirectory(string username)
    {
        try
        {
            using (var context = new PrincipalContext(ContextType.Domain))
            {
                using (var user = UserPrincipal.FindByIdentity(context, username))
                {
                    return user != null;
                }
            }
        }
        catch
        {
            return true;
        }
    }

    [HttpPost]
    public IActionResult CheckNumberIdExists([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData != null && requestData.ContainsKey("nFuncionario"))
        {
            string numberId = requestData["nFuncionario"];
            if (string.IsNullOrEmpty(numberId))
            {
                return Json(new { success = false, exists = false, message = "El identificador está vacío." });
            }

            try
            {
                string domain = "aytosa.inet";
                string attributeName = "description";
                string ldapPath = $"LDAP://{domain}";

                using (DirectoryEntry entry = new DirectoryEntry(ldapPath))
                {
                    using (DirectorySearcher searcher = new DirectorySearcher(entry))
                    {
                        searcher.Filter = $"({attributeName}={numberId})";
                        searcher.SearchScope = SearchScope.Subtree;
                        SearchResult result = searcher.FindOne();

                        if (result != null)
                        {
                            return Json(new { success = true, exists = true, message = "El identificador ya existe." });
                        }
                        else
                        {
                            return Json(new { success = true, exists = false, message = "El identificador no existe." });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, exists = false, message = $"Error al buscar el identificador: {ex.Message}" });
            }
        }

        return Json(new { success = false, exists = false, message = "No se recibió el identificador." });
    }

    [HttpPost]
    public IActionResult CheckTelephoneExists([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData != null && requestData.ContainsKey("nTelefono"))
        {
            string telefono = requestData["nTelefono"];
            if (string.IsNullOrEmpty(telefono))
            {
                return Json(new { success = false, exists = false, message = "El campo teléfono está vacío." });
            }

            try
            {
                string domain = "aytosa.inet";
                string attributeName = "telephoneNumber";
                string ldapPath = $"LDAP://{domain}";

                using (DirectoryEntry entry = new DirectoryEntry(ldapPath))
                {
                    using (DirectorySearcher searcher = new DirectorySearcher(entry))
                    {
                        searcher.Filter = $"({attributeName}={telefono})";
                        searcher.SearchScope = SearchScope.Subtree;
                        SearchResult result = searcher.FindOne();

                        if (result != null)
                        {
                            return Json(new { success = true, exists = true, message = "El teléfono ya existe." });
                        }
                        else
                        {
                            return Json(new { success = true, exists = false, message = "El teléfono no existe." });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, exists = false, message = $"Error al buscar el identificador: {ex.Message}" });
            }
        }

        return Json(new { success = false, exists = false, message = "No se recibió el identificador." });
    }

    [HttpPost]
    public IActionResult GenerateUsername([FromBody] userInputModel userInput)
    {
        if (string.IsNullOrEmpty(userInput.Nombre) || string.IsNullOrEmpty(userInput.Apellido1) || string.IsNullOrEmpty(userInput.Apellido2))
        {
            return Json(new { success = true, username = "" });
        }

        try
        {
            string[] nombrePartes = userInput.Nombre.Trim().ToLower().Split(' ');
            string[] apellido1Partes = userInput.Apellido1.Trim().ToLower().Split(' ');
            string[] apellido2Partes = string.IsNullOrEmpty(userInput.Apellido2)
                ? new string[0]
                : userInput.Apellido2.Trim().ToLower().Split(' ');

            List<string> candidatos = new List<string>();
            string candidato1 = $"{GetInicial(nombrePartes)}{GetCompleto(apellido1Partes)}{GetInicial(apellido2Partes)}";
            candidatos.Add(candidato1.Substring(0, Math.Min(12, candidato1.Length)));
            string candidato2 = $"{GetNombreCompuesto(nombrePartes)}{GetInicial(apellido1Partes)}{GetInicial(apellido2Partes)}";
            candidatos.Add(candidato2.Substring(0, Math.Min(12, candidato2.Length)));
            string candidato3 = $"{GetInicial(nombrePartes)}{GetInicial(apellido1Partes)}{GetCompleto(apellido2Partes)}";
            candidatos.Add(candidato3.Substring(0, Math.Min(12, candidato3.Length)));

            foreach (string candidato in candidatos)
            {
                if (!CheckUserInActiveDirectory(candidato))
                {
                    return Json(new { success = true, username = candidato });
                }
            }

            return Json(new { success = false, message = "No se pudo generar un nombre de usuario único." });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al generar el nombre de usuario: {ex.Message}" });
        }
    }

    private string GetInicial(string[] partes)
    {
        return partes.Length > 0 ? partes[0][0].ToString() : "";
    }

    private string GetNombreCompuesto(string[] partes)
    {
        if (partes.Length == 0) return "";
        return partes[0] + string.Join("", partes.Skip(1).Select(p => p[0]));
    }

    private string GetCompleto(string[] partes)
    {
        return partes.Length > 0 ? string.Join("", partes) : "";
    }

    [HttpPost]
    public IActionResult ManageUserStatus([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("username") || !requestData.ContainsKey("action"))
        {
            return Json(new { success = false, message = "Datos inválidos. Se requiere el usuario y la acción." });
        }

        string input = requestData["username"];
        string action = requestData["action"].ToLower();
        string username = ExtractUsername(input); // Este método se mantiene aquí porque lo usa esta vista

        if (string.IsNullOrEmpty(username))
        {
            return Json(new { success = false, message = "El formato del usuario seleccionado no es válido." });
        }

        try
        {
            string ldapPath = $"LDAP://DC=aytosa,DC=inet";
            using (DirectoryEntry root = new DirectoryEntry(ldapPath))
            {
                using (DirectorySearcher searcher = new DirectorySearcher(root))
                {
                    searcher.Filter = $"(&(objectClass=user)(sAMAccountName={username}))";
                    searcher.SearchScope = SearchScope.Subtree;
                    searcher.PropertiesToLoad.Add("userAccountControl");

                    SearchResult result = searcher.FindOne();
                    if (result == null)
                    {
                        return Json(new { success = false, message = $"Usuario {username} no encontrado." });
                    }

                    using (DirectoryEntry userEntry = result.GetDirectoryEntry())
                    {
                        int userAccountControl = (int)userEntry.Properties["userAccountControl"].Value;
                        if (action == "enable")
                        {
                            userAccountControl &= ~0x2;
                        }
                        else if (action == "disable")
                        {
                            userAccountControl |= 0x2;
                        }
                        else
                        {
                            return Json(new { success = false, message = "Acción no válida. Use 'enable' o 'disable'." });
                        }

                        userEntry.Properties["userAccountControl"].Value = userAccountControl;
                        userEntry.CommitChanges();

                        return Json(new { success = true, message = $"Usuario {username} {(action == "enable" ? "habilitado" : "deshabilitado")} correctamente." });
                    }
                }
            }
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al realizar la acción: {ex.Message}" });
        }
    }

    private string ExtractUsername(string input) // Se mantiene aquí porque lo usa ManageUserStatus
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return null;
        }

        int startIndex = input.LastIndexOf('(');
        int endIndex = input.LastIndexOf(')');
        if (startIndex >= 0 && endIndex > startIndex)
        {
            return input.Substring(startIndex + 1, endIndex - startIndex - 1).Trim();
        }

        return null;
    }

    [HttpPost]
    public IActionResult GetUserGroups([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("username"))
            return Json(new { success = false, message = "Usuario no especificado." });

        string username = ExtractUsername(requestData["username"]);
        if (string.IsNullOrEmpty(username))
            return Json(new { success = false, message = "El formato del usuario seleccionado no es válido." });

        try
        {
            using (var context = new PrincipalContext(ContextType.Domain))
            using (var user = UserPrincipal.FindByIdentity(context, username))
            {
                if (user == null)
                    return Json(new { success = false, message = "Usuario no encontrado." });

                var groups = user.GetAuthorizationGroups()
                                 .Where(g => g is GroupPrincipal)
                                 .Select(g => g.Name)
                                 .ToList();

                return Json(new { success = true, groups });
            }
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al obtener los grupos del usuario: {ex.Message}" });
        }
    }

    [HttpPost]
    public IActionResult ModifyUserGroup([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("username") || !requestData.ContainsKey("group") || !requestData.ContainsKey("action"))
        {
            return Json(new { success = false, message = "Datos insuficientes para modificar el grupo." });
        }

        string input = requestData["username"];
        string username = ExtractUsername(input);
        if (string.IsNullOrEmpty(username))
            return Json(new { success = false, message = "El formato del usuario seleccionado no es válido." });

        string group = requestData["group"];
        string action = requestData["action"];

        try
        {
            // Este método se eliminó, necesitarás mover FindGroupByName aquí si lo usas en otras vistas
            using (var context = new PrincipalContext(ContextType.Domain, "aytosa.inet"))
            using (var user = UserPrincipal.FindByIdentity(context, username))
            {
                if (user == null)
                    return Json(new { success = false, message = "Usuario no encontrado en Active Directory." });

                using (var userEntry = (DirectoryEntry)user.GetUnderlyingObject())
                {
                    string userDN = userEntry.Properties["distinguishedName"].Value.ToString();
                    // Aquí necesitarías FindGroupByName o una alternativa
                    return Json(new { success = false, message = "FindGroupByName no está disponible aquí." });
                }
            }
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al modificar grupo: {ex.Message}" });
        }
    }

    [HttpPost]
    public IActionResult ModifyUserOU([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("username") || !requestData.ContainsKey("ouPrincipal") || !requestData.ContainsKey("ouSecundaria") || !requestData.ContainsKey("departamento") || !requestData.ContainsKey("lugarEnvio"))
        {
            return Json(new { success = false, message = "Datos insuficientes para modificar la información del usuario." });
        }

        string input = requestData["username"];
        string username = ExtractUsername(input);
        string newOUPrincipal = requestData["ouPrincipal"];
        string newOUSecundaria = requestData["ouSecundaria"];
        string newDepartamento = requestData["departamento"];
        string newLugarEnvio = requestData["lugarEnvio"];

        if (string.IsNullOrEmpty(username))
            return Json(new { success = false, message = "El formato del usuario seleccionado no es válido." });

        try
        {
            string ldapBasePath = "LDAP://DC=aytosa,DC=inet";
            string newLDAPPath = $"LDAP://OU={newOUSecundaria},OU=Usuarios y Grupos,OU={newOUPrincipal},DC=aytosa,DC=inet";

            using (DirectoryEntry root = new DirectoryEntry(ldapBasePath))
            {
                using (DirectorySearcher searcher = new DirectorySearcher(root))
                {
                    searcher.Filter = $"(&(objectClass=user)(sAMAccountName={username}))";
                    searcher.SearchScope = SearchScope.Subtree;
                    searcher.PropertiesToLoad.Add("distinguishedName");

                    SearchResult result = searcher.FindOne();
                    if (result == null)
                        return Json(new { success = false, message = $"Usuario {username} no encontrado en Active Directory." });

                    using (DirectoryEntry userEntry = result.GetDirectoryEntry())
                    {
                        using (DirectoryEntry newOUEntry = new DirectoryEntry(newLDAPPath))
                        {
                            userEntry.MoveTo(newOUEntry);
                        }
                        userEntry.Properties["physicalDeliveryOfficeName"].Value = newDepartamento;
                        userEntry.Properties["streetAddress"].Value = newLugarEnvio;
                        userEntry.CommitChanges();
                    }
                }
            }

            return Json(new { success = true, message = "OU, departamento y lugar de envío modificados correctamente." });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al modificar el usuario: {ex.Message}" });
        }
    }

    [HttpPost]
    public IActionResult GetUserDetails([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("username"))
            return Json(new { success = false, message = "Usuario no especificado." });

        string input = requestData["username"];
        string username = ExtractUsername(input);
        if (string.IsNullOrEmpty(username))
            return Json(new { success = false, message = "El formato del usuario seleccionado no es válido." });

        try
        {
            string ldapPath = "LDAP://DC=aytosa,DC=inet";
            string currentOU = "";
            List<string> groups = new List<string>();

            using (DirectoryEntry root = new DirectoryEntry(ldapPath))
            {
                using (DirectorySearcher searcher = new DirectorySearcher(root))
                {
                    searcher.Filter = $"(&(objectClass=user)(sAMAccountName={username}))";
                    searcher.SearchScope = SearchScope.Subtree;
                    searcher.PropertiesToLoad.Add("distinguishedName");
                    searcher.PropertiesToLoad.Add("memberOf");

                    SearchResult result = searcher.FindOne();
                    if (result == null)
                        return Json(new { success = false, message = $"Usuario {username} no encontrado en Active Directory." });

                    using (DirectoryEntry userEntry = result.GetDirectoryEntry())
                    {
                        string distinguishedName = userEntry.Properties["distinguishedName"].Value.ToString();
                        currentOU = ExtractOUFromDN(distinguishedName);
                        groups = GetUserGroups(userEntry);
                    }
                }
            }
            return Json(new { success = true, currentOU, groups });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al obtener los detalles del usuario: {ex.Message}" });
        }
    }

    private string ExtractOUFromDN(string distinguishedName)
    {
        if (string.IsNullOrEmpty(distinguishedName))
            return "";

        string[] parts = distinguishedName.Split(',');
        foreach (string part in parts)
        {
            if (part.StartsWith("OU="))
                return part.Replace("OU=", "").Trim();
        }
        return "No se encontró OU";
    }

    private List<string> GetUserGroups(DirectoryEntry userEntry)
    {
        List<string> groups = new List<string>();
        if (userEntry.Properties["memberOf"] != null)
        {
            foreach (var groupDN in userEntry.Properties["memberOf"])
            {
                string cn = ExtractCNFromDN(groupDN.ToString());
                if (!string.IsNullOrEmpty(cn))
                    groups.Add(cn);
            }
        }
        return groups;
    }

    private string ExtractCNFromDN(string distinguishedName)
    {
        if (!string.IsNullOrEmpty(distinguishedName))
        {
            int start = distinguishedName.IndexOf("CN=");
            if (start >= 0)
            {
                int end = distinguishedName.IndexOf(",", start);
                if (end > start)
                    return distinguishedName.Substring(start + 3, end - start - 3);
                else
                    return distinguishedName.Substring(start + 3);
            }
        }
        return "";
    }

    private void RunPowerShellScript(string scriptText)
    {
        using (PowerShell ps = PowerShell.Create())
        {
            ps.AddScript(scriptText);
            ps.Invoke();

            if (ps.Streams.Error.Count > 0)
            {
                foreach (var error in ps.Streams.Error)
                {
                    Console.WriteLine($"Error en PowerShell: {error}");
                }
            }
        }
    }
}