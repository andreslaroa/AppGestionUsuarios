using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.IO;
using System.DirectoryServices.AccountManagement;
using System;
using System.DirectoryServices;
using System.Globalization;
using System.Linq;
using System.Text;
using System.DirectoryServices.ActiveDirectory;
using static Microsoft.ApplicationInsights.MetricDimensionNames.TelemetryContext;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.Extensions.Logging.Abstractions;
using System.Collections.ObjectModel;
using System.Management.Automation;


[Authorize]
public class GestionUsuariosController : Controller
{   

    public class UserDetailsResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public string OUPrincipal { get; set; }
        public string OUSecundaria { get; set; }
        public string Departamento { get; set; }
        public string LugarEnvio { get; set; }
        public List<string> Groups { get; set; }
    }

    public class userInputModel
    {
        public string Nombre { set; get; }
        public string Apellido1 { set; get; }
        public string Apellido2 { set; get; }
    }

    public class UserDetailsRequest
    {
        public string Username { get; set; }
    }

    [HttpGet]
    public IActionResult HabilitarDeshabilitarUsuario()
    {
        try
        {
            // Establecer el límite de resultados por página
            int pageSize = 1000; // Puedes cambiarlo si el servidor tiene un límite diferente

            using (var entry = new DirectoryEntry("LDAP://DC=aytosa,DC=inet"))
            {
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = "(objectClass=user)";
                    searcher.PropertiesToLoad.Add("displayName");
                    searcher.PropertiesToLoad.Add("sAMAccountName");
                    searcher.SearchScope = SearchScope.Subtree;

                    // Habilitar la paginación
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

                    ViewBag.Usuarios = usuarios.OrderBy(u => u).ToList(); // Ordenar alfabéticamente
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


    [HttpGet]
    public IActionResult ModificarUsuario()
    {
        try
        {

            int pageSize = 1000;

            // Obtener todos los usuarios del Directorio Activo
            using (var entry = new DirectoryEntry("LDAP://DC=aytosa,DC=inet"))
            {
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = "(objectClass=user)";
                    searcher.PropertiesToLoad.Add("displayName");
                    searcher.PropertiesToLoad.Add("sAMAccountName");
                    searcher.SearchScope = SearchScope.Subtree;

                    // Habilitar la paginación
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

                    ViewBag.Users = usuarios.OrderBy(u => u).ToList(); // Ordenar usuarios alfabéticamente
                }
            }

            // Obtener lista de grupos del Directorio Activo
            using (var entry = new DirectoryEntry("LDAP://DC=aytosa,DC=inet"))
            {
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = "(objectClass=group)";
                    searcher.PropertiesToLoad.Add("cn");
                    searcher.SearchScope = SearchScope.Subtree;

                    var grupos = new List<string>();

                    foreach (SearchResult result in searcher.FindAll())
                    {
                        if (result.Properties.Contains("cn"))
                        {
                            grupos.Add(result.Properties["cn"][0].ToString());
                        }
                    }

                    ViewBag.GruposAD = grupos.OrderBy(g => g).ToList(); // Ordenar grupos alfabéticamente
                }
            }

            // Obtener lista de OUs desde el servicio asociado al Excel
            var ouPrincipales = GetOUPrincipalesFromAD();
            ViewBag.OUPrincipales = ouPrincipales;
        }
        catch (Exception ex)
        {
            ViewBag.Users = new List<string>();
            ViewBag.GruposAD = new List<string>();
            ViewBag.OUPrincipales = new List<string>();
            Console.WriteLine($"Error al cargar los datos para la vista de modificación: {ex.Message}");
        }

        return View();
    }





    private List<string> GetOUPrincipalesFromAD()
    {
        var ouPrincipales = new List<string>();

        try
        {
            using (var rootEntry = new DirectoryEntry("LDAP://DC=aytosa,DC=inet"))
            {
                using (var searcher = new DirectorySearcher(rootEntry))
                {
                    // Buscar la OU "AREAS" como base
                    searcher.Filter = "(&(objectClass=organizationalUnit)(ou=AREAS))";
                    searcher.SearchScope = SearchScope.Subtree;
                    searcher.PropertiesToLoad.Add("distinguishedName");

                    SearchResult areasResult = searcher.FindOne();
                    if (areasResult == null)
                    {
                        throw new Exception("No se encontró la OU 'AREAS' en el Active Directory.");
                    }

                    string areasPath = areasResult.Path;

                    // Buscar las sub-OUs bajo "AREAS"
                    using (var areasEntry = new DirectoryEntry(areasPath))
                    {
                        foreach (DirectoryEntry child in areasEntry.Children)
                        {
                            if (child.SchemaClassName == "organizationalUnit")
                            {
                                string ouName = child.Properties["ou"].Value?.ToString();
                                if (!string.IsNullOrEmpty(ouName))
                                {
                                    ouPrincipales.Add(ouName);
                                }
                            }
                        }
                    }
                }
            }
            ouPrincipales.Sort();
        }
        catch (Exception ex)
        {
            throw new Exception("Error al obtener las OUs principales del Active Directory: " + ex.Message, ex);
        }

        return ouPrincipales;
    }




    [HttpPost]
    public IActionResult GetOUSecundarias([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("ouPrincipal"))
        {
            return Json(new List<string>());
        }

        string ouPrincipal = requestData["ouPrincipal"];
        if (string.IsNullOrEmpty(ouPrincipal))
        {
            return Json(new List<string>());
        }

        try
        {
            var ouSecundarias = new List<string>();
            string ldapPath = $"LDAP://OU=Usuarios y Grupos,OU={ouPrincipal},OU=AREAS,DC=aytosa,DC=inet";

            using (var rootEntry = new DirectoryEntry(ldapPath))
            {
                foreach (DirectoryEntry child in rootEntry.Children)
                {
                    if (child.SchemaClassName == "organizationalUnit")
                    {
                        string ouName = child.Properties["ou"].Value?.ToString();
                        if (!string.IsNullOrEmpty(ouName))
                        {
                            ouSecundarias.Add(ouName);
                        }
                    }
                }
            }

            ouSecundarias.Sort();
            return Json(ouSecundarias); // Devolver solo la lista de OUs secundarias
        }
        catch (Exception ex)
        {
            // En lugar de lanzar una excepción, devolvemos una lista vacía y registramos el error
            Console.WriteLine($"Error al obtener las OU secundarias para {ouPrincipal}: {ex.Message}");
            return Json(new List<string>());
        }
    }


    [HttpPost]
    public IActionResult GetDepartamentos([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("ouPrincipal"))
        {
            return Json(new { success = false, message = "OU principal no especificada." });
        }

        string ouPrincipal = requestData["ouPrincipal"];
        string ouSecundaria = requestData.ContainsKey("ouSecundaria") ? requestData["ouSecundaria"] : null;

        if (string.IsNullOrEmpty(ouPrincipal))
        {
            return Json(new { success = false, message = "OU principal no puede estar vacía." });
        }

        try
        {
            string ldapPath;
            if (!string.IsNullOrEmpty(ouSecundaria))
            {
                // Si hay OU secundaria, buscamos el departamento en la OU secundaria dentro de "Usuarios y Grupos"
                ldapPath = $"LDAP://OU={ouSecundaria},OU=Usuarios y Grupos,OU={ouPrincipal},OU=AREAS,DC=aytosa,DC=inet";
            }
            else
            {
                // Si no hay OU secundaria, buscamos el departamento en la OU principal
                ldapPath = $"LDAP://OU={ouPrincipal},OU=AREAS,DC=aytosa,DC=inet";
            }

            using (var ouEntry = new DirectoryEntry(ldapPath))
            {
                if (ouEntry == null)
                {
                    return Json(new { success = false, message = "No se pudo conectar a la OU especificada.", ldapPath });
                }

                // Obtener el atributo 'description' de la OU
                string departamento = ouEntry.Properties["description"]?.Value?.ToString();
                if (string.IsNullOrEmpty(departamento))
                {
                    // Si no hay descripción, usar el nombre de la OU como valor predeterminado
                    departamento = ouEntry.Properties["ou"]?.Value?.ToString() ?? "Sin departamento";
                }

                return Json(new { success = true, departamento, ldapPath });
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Error al obtener el departamento: {ex.Message}, OU Principal: {ouPrincipal}, OU Secundaria: {ouSecundaria}", ex);
        }
    }
    [HttpPost]
    public IActionResult GetLugarEnvio([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("ouPrincipal"))
        {
            return Json(new { success = false, message = "OU principal no especificada." });
        }

        string ouPrincipal = requestData["ouPrincipal"];
        string ouSecundaria = requestData.ContainsKey("ouSecundaria") ? requestData["ouSecundaria"] : null;

        if (string.IsNullOrEmpty(ouPrincipal))
        {
            return Json(new { success = false, message = "OU principal no puede estar vacía." });
        }

        try
        {
            string ldapPath;
            if (!string.IsNullOrEmpty(ouSecundaria))
            {
                // Si hay OU secundaria, buscamos el lugar de envío en la OU secundaria dentro de "Usuarios y Grupos"
                ldapPath = $"LDAP://OU={ouSecundaria},OU=Usuarios y Grupos,OU={ouPrincipal},OU=AREAS,DC=aytosa,DC=inet";
            }
            else
            {
                // Si no hay OU secundaria, buscamos el lugar de envío en la OU principal
                ldapPath = $"LDAP://OU={ouPrincipal},OU=AREAS,DC=aytosa,DC=inet";
            }

            using (var ouEntry = new DirectoryEntry(ldapPath))
            {
                if (ouEntry == null)
                {
                    return Json(new { success = false, message = "No se pudo conectar a la OU especificada.", ldapPath });
                }

                // Obtener el atributo 'city' de la OU
                string lugarEnvio = ouEntry.Properties["l"]?.Value?.ToString(); // El atributo 'city' en AD es 'l' (lowercase L)
                if (string.IsNullOrEmpty(lugarEnvio))
                {
                    // Si no hay valor para 'city', usar un valor predeterminado
                    lugarEnvio = "Sin lugar de envío";
                }

                return Json(new { success = true, lugarEnvio, ldapPath });
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Error al obtener el lugar de envío: {ex.Message}, OU Principal: {ouPrincipal}, OU Secundaria: {ouSecundaria}", ex);
        }
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

        return Json(false); // Si no hay datos, asumimos que no existe
    }

    private bool CheckUserInActiveDirectory(string username)
    {
        try
        {
            using (var context = new PrincipalContext(ContextType.Domain))
            {
                using (var user = UserPrincipal.FindByIdentity(context, username))
                {
                    return user != null; // Retorna true si el usuario existe
                }
            }
        }
        catch
        {
            // Manejo de errores
            return true; // Asumir que el usuario existe si hay un error
        }
    }

    

    // Función para obtener la inicial de la primera palabra
    private string GetInicial(string[] partes)
    {
        return partes.Length > 0 ? partes[0][0].ToString() : "";
    }

    // Función para obtener el atributo completo (primera palabra completa y las iniciales de las demás)
    private string GetNombreCompuesto(string[] partes)
    {
        if (partes.Length == 0) return "";
        return partes[0] + string.Join("", partes.Skip(1).Select(p => p[0]));
    }

    // Función para obtener el atributo completo
    private string GetCompleto(string[] partes)
    {
        return partes.Length > 0 ? string.Join("", partes) : "";
    }

    [HttpPost]
    public IActionResult CheckNumberIdExists([FromBody] Dictionary<string, string> requestData)
    {
        // Validar si se recibió el campo nFuncionario
        if (requestData != null && requestData.ContainsKey("nFuncionario"))
        {
            string numberId = requestData["nFuncionario"];

            // Validar si el identificador es nulo o vacío
            if (string.IsNullOrEmpty(numberId))
            {
                return Json(new { success = false, exists = false, message = "El identificador está vacío." });
            }

            try
            {
                // Configurar dominio y atributo a buscar
                string domain = "aytosa.inet"; // Ajusta al dominio de tu entorno
                string attributeName = "description"; // Atributo del Directorio Activo para el número de funcionario

                // Ruta LDAP al dominio
                string ldapPath = $"LDAP://{domain}";

                using (DirectoryEntry entry = new DirectoryEntry(ldapPath))
                {
                    using (DirectorySearcher searcher = new DirectorySearcher(entry))
                    {
                        // Filtro LDAP para buscar el identificador
                        searcher.Filter = $"({attributeName}={numberId})";
                        searcher.SearchScope = SearchScope.Subtree;

                        // Buscar el identificador en el Directorio Activo
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
                // Manejo de errores
                return Json(new { success = false, exists = false, message = $"Error al buscar el identificador: {ex.Message}" });
            }
        }

        return Json(new { success = false, exists = false, message = "No se recibió el identificador." });
    }


    [HttpPost]
    public IActionResult CheckTelephoneExists([FromBody] Dictionary<string, string> requestData)
    {
        // Validar si se recibió el campo nFuncionario
        if (requestData != null && requestData.ContainsKey("nTelefono"))
        {
            string telefono = requestData["nTelefono"];

            // Validar si el identificador es nulo o vacío
            if (string.IsNullOrEmpty(telefono))
            {
                return Json(new { success = false, exists = false, message = "El campo teléfono está vacío." });
            }

            try
            {
                // Configurar dominio y atributo a buscar
                string domain = "aytosa.inet"; // Ajusta al dominio de tu entorno
                string attributeName = "telephoneNumber"; // Atributo del Directorio Activo para el número de funcionario

                // Ruta LDAP al dominio
                string ldapPath = $"LDAP://{domain}";

                using (DirectoryEntry entry = new DirectoryEntry(ldapPath))
                {
                    using (DirectorySearcher searcher = new DirectorySearcher(entry))
                    {
                        // Filtro LDAP para buscar el identificador
                        searcher.Filter = $"({attributeName}={telefono})";
                        searcher.SearchScope = SearchScope.Subtree;

                        // Buscar el identificador en el Directorio Activo
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
                // Manejo de errores
                return Json(new { success = false, exists = false, message = $"Error al buscar el identificador: {ex.Message}" });
            }
        }

        return Json(new { success = false, exists = false, message = "No se recibió el identificador." });
    }

        
    //Función para buscar el grupo en el dominio del directorio activo
    private DirectoryEntry FindGroupByName(string groupName)
    {
        if (string.IsNullOrEmpty(groupName))
        {
            return null;
        }

        try
        {
            // Ruta base del dominio
            string domainPath = "LDAP://DC=aytosa,DC=inet";

            // Crear una entrada de directorio
            using (DirectoryEntry rootEntry = new DirectoryEntry(domainPath))
            {
                using (DirectorySearcher searcher = new DirectorySearcher(rootEntry))
                {
                    // Filtro para encontrar el grupo por nombre (CN)
                    searcher.Filter = $"(&(objectClass=group)(cn={groupName}))";
                    searcher.SearchScope = SearchScope.Subtree; // Asegura búsqueda en todo el dominio
                    searcher.PropertiesToLoad.Add("distinguishedName"); // Solo carga lo necesario

                    SearchResult result = searcher.FindOne();
                    if (result != null)
                    {
                        return result.GetDirectoryEntry();
                    }
                    else
                    {
                        Console.WriteLine($"Grupo '{groupName}' no encontrado en el dominio.");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al buscar el grupo '{groupName}': {ex.Message}");
        }

        return null; // Devuelve null si no se encuentra o si ocurre un error
    }




    private (bool success, string message) ConfigurarDirectorioYCuotaRemoto(string username, string quota)
    {
        try
        {
            // Script de PowerShell para ejecutar de forma remota en LEONARDO
            string script = $@"
        param(
            [string]$nameUID,
            [string]$quota
        )
        New-FsrmQuota -Path ('G:\HOME\' + $nameUID) -Template ('Users-' + $quota)
        ";

            // Configuración del comando remoto
            string comandoRemoto = $@"
        Invoke-Command -ComputerName ribera -ScriptBlock {{
            {script}
        }} -ArgumentList '{username}', '{quota}'
        ";

            using (PowerShell powerShell = PowerShell.Create())
            {
                powerShell.AddScript(comandoRemoto);

                // Ejecutar el script
                var result = powerShell.Invoke();

                // Verificar errores en la ejecución
                if (powerShell.Streams.Error.Count > 0)
                {
                    var errores = powerShell.Streams.Error.Select(e => e.ToString()).ToList();
                    return (false, string.Join("; ", errores));
                }

                return (true, "Directorio y cuota configurados exitosamente en LEONARDO.");
            }
        }
        catch (Exception ex)
        {
            return (false, $"Error en PowerShell: {ex.Message}");
        }
    }




    //Método para convertir el valor de la cuota a numérico
    private int ObtenerCuotaEnMB(string cuotaEnMB)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(cuotaEnMB))
            {
                throw new ArgumentException("La cuota no puede estar vacía.");
            }

            // Extraer el número antes del espacio
            string[] partes = cuotaEnMB.Split(' ');
            if (partes.Length == 0 || !int.TryParse(partes[0], out int cuota))
            {
                throw new FormatException("El formato de la cuota es inválido.");
            }

            return cuota; // Devuelve el número en MB
        }
        catch (Exception ex)
        {
            throw new ArgumentException($"Error al procesar la cuota: {ex.Message}");
        }
    }


    // Método para eliminar acentos de una cadena
    private static string RemoveAccents(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return text;

        text = text.Normalize(NormalizationForm.FormD);
        char[] chars = text
            .Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
            .ToArray();

        return new string(chars).Normalize(NormalizationForm.FormC);
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

        try
        {
            // Extraer el nombre de usuario entre paréntesis
            var username = ExtractUsername(input);

            if (string.IsNullOrEmpty(username))
            {
                return Json(new { success = false, message = "El formato del usuario seleccionado no es válido." });
            }

            // Buscar al usuario en el Directorio Activo
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
                            // Quitar el flag "AccountDisabled" (0x2)
                            userAccountControl &= ~0x2;
                        }
                        else if (action == "disable")
                        {
                            // Agregar el flag "AccountDisabled" (0x2)
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

    // Función para extraer el nombre de usuario entre paréntesis
    private string ExtractUsername(string input)
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

        string input = requestData["username"];
        string username = ExtractUsername(input);

        if (string.IsNullOrEmpty(username))
        {
            return Json(new { success = false, message = "Formato del usuario inválido." });
        }

        try
        {
            string ldapPath = $"LDAP://DC=aytosa,DC=inet";
            using (DirectoryEntry root = new DirectoryEntry(ldapPath))
            {
                using (DirectorySearcher searcher = new DirectorySearcher(root))
                {
                    // Búsqueda del usuario por sAMAccountName
                    searcher.Filter = $"(&(objectClass=user)(sAMAccountName={username}))";
                    searcher.SearchScope = SearchScope.Subtree;
                    searcher.PropertiesToLoad.Add("memberOf");

                    SearchResult result = searcher.FindOne();

                    if (result == null)
                    {
                        return Json(new { success = false, message = $"Usuario {username} no encontrado." });
                    }

                    // Obtener la lista de grupos del usuario
                    var groupList = new List<string>();
                    if (result.Properties.Contains("memberOf"))
                    {
                        foreach (var group in result.Properties["memberOf"])
                        {
                            // Extraer solo el CN (nombre del grupo)
                            string groupName = ExtractGroupName(group.ToString());
                            if (!string.IsNullOrEmpty(groupName))
                            {
                                groupList.Add(groupName);
                            }
                        }
                    }

                    return Json(new { success = true, groups = groupList });
                }
            }
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al obtener los grupos del usuario: {ex.Message}" });
        }
    }

    // Función para extraer el CN (nombre del grupo) de un distinguishedName
    private string ExtractGroupName(string distinguishedName)
    {
        if (string.IsNullOrWhiteSpace(distinguishedName))
            return null;

        var parts = distinguishedName.Split(',');
        foreach (var part in parts)
        {
            if (part.StartsWith("CN=", StringComparison.OrdinalIgnoreCase))
            {
                return part.Substring(3); // Eliminar "CN=" y devolver el nombre
            }
        }

        return null;
    }


    [HttpPost]
    public IActionResult ModifyUserGroup([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("username") || !requestData.ContainsKey("group") || !requestData.ContainsKey("action"))
            return Json(new { success = false, message = "Datos insuficientes para modificar el grupo." });

        string input = requestData["username"];
        string username = ExtractUsername(input); // Extrae el nombre de usuario
        string groupName = requestData["group"];  // El nombre limpio del grupo
        string action = requestData["action"].ToLower();

        DirectoryEntry groupEntry = null; // Declaración fuera del try

        try
        {
            // Buscar el grupo en el dominio
            groupEntry = FindGroupByName(groupName);
            if (groupEntry == null)
                return Json(new { success = false, message = $"Grupo '{groupName}' no encontrado en el dominio." });

            // Buscar el usuario en el dominio
            string ldapPath = "LDAP://DC=aytosa,DC=inet";
            using (var root = new DirectoryEntry(ldapPath))
            {
                using (var searcher = new DirectorySearcher(root))
                {
                    searcher.Filter = $"(&(objectClass=user)(sAMAccountName={username}))";
                    searcher.SearchScope = SearchScope.Subtree;

                    SearchResult result = searcher.FindOne();

                    if (result == null)
                        return Json(new { success = false, message = $"Usuario '{username}' no encontrado en el dominio." });

                    using (DirectoryEntry userEntry = result.GetDirectoryEntry())
                    {
                        if (action == "add")
                        {
                            // Añadir el usuario al grupo
                            groupEntry.Invoke("Add", new object[] { userEntry.Path });
                            groupEntry.CommitChanges();
                            return Json(new { success = true, message = $"El usuario '{username}' fue añadido al grupo '{groupName}' correctamente." });
                        }
                        else if (action == "remove")
                        {
                            // Eliminar el usuario del grupo
                            groupEntry.Invoke("Remove", new object[] { userEntry.Path });
                            groupEntry.CommitChanges();
                            return Json(new { success = true, message = $"El usuario '{username}' fue eliminado del grupo '{groupName}' correctamente." });
                        }
                        else
                        {
                            return Json(new { success = false, message = "Acción no válida. Use 'add' o 'remove'." });
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al modificar el grupo: {ex.Message}" });
        }
        finally
        {
            if (groupEntry != null)
            {
                groupEntry.Dispose(); // Liberar el recurso correctamente
            }
        }
    }


    [HttpPost]
    public IActionResult ModifyUserOU([FromBody] Dictionary<string, string> requestData)
    {
        // Validación de parámetros
        if (requestData == null ||
            !requestData.ContainsKey("username") ||
            !requestData.ContainsKey("ouPrincipal") ||
            !requestData.ContainsKey("departamento") ||
            !requestData.ContainsKey("lugarEnvio"))
        {
            return Json(new { success = false, message = "Faltan datos para modificar la OU." });
        }

        // Extracción de valores de entrada
        string username = ExtractUsername(requestData["username"]);
        string ouPrincipal = requestData["ouPrincipal"];
        // ouSecundaria puede venir como null o cadena vacía
        requestData.TryGetValue("ouSecundaria", out var ouSecundaria);
        string departamento = requestData["departamento"];
        string lugarEnvio = requestData["lugarEnvio"];

        // Construcción de la ruta LDAP según OU secundaria opcional
        string domainDn = "DC=aytosa,DC=inet";
        string newLDAPPath;
        if (string.IsNullOrWhiteSpace(ouSecundaria))
        {
            // Sin OU secundaria -> uso solo OU principal
            newLDAPPath = $"LDAP://OU={ouPrincipal},OU=AREAS,{domainDn}";
        }
        else
        {
            // Con OU secundaria -> ruta completa
            newLDAPPath = $"LDAP://OU={ouSecundaria},OU=Usuarios y Grupos,OU={ouPrincipal},OU=AREAS ,{domainDn}";
        }

        try
        {
            using (var context = new PrincipalContext(ContextType.Domain))
            using (var user = UserPrincipal.FindByIdentity(context, username))
            {
                if (user == null)
                    return Json(new { success = false, message = "Usuario no encontrado." });

                using (var directoryUser = (DirectoryEntry)user.GetUnderlyingObject())
                {
                    // Mover el usuario a la OU calculada
                    directoryUser.MoveTo(new DirectoryEntry(newLDAPPath));

                    // Actualizar departamento u otro atributo si es necesario
                    directoryUser.Properties["physicalDeliveryOfficeName"].Value = departamento;
                    directoryUser.CommitChanges();
                }
            }

            return Json(new { success = true, message = "OU del usuario modificada correctamente." });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al modificar la OU: {ex.Message}" });
        }
    }


}