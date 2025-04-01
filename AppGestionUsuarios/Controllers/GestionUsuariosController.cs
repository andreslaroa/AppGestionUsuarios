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
    private readonly OUService _ouService;
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

        // Nueva propiedad para los grupos seleccionados
        public List<string> Grupos { get; set; }
    }


    public class userInputModel
    {
        public string Nombre { set; get; }
        public string Apellido1 { set; get; }
        public string Apellido2 { set; get; }
    }



        public GestionUsuariosController()
    {
        string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Resources", "ArchivoDePruebasOU.xlsx");
        _ouService = new OUService(filePath);
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
            var ouPrincipales = _ouService.GetOUPrincipales();
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

   










   


    //Nos muestra una lista de las OU secundarias recibiendo la principal como parámetro
    [HttpPost]
    public IActionResult GetOUSecundarias([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData != null && requestData.ContainsKey("ouPrincipal"))
        {
            string ouPrincipal = requestData["ouPrincipal"];
            var ouSecundarias = _ouService.GetOUSecundarias(ouPrincipal);
            return Json(ouSecundarias);
        }

        return Json(new List<string>());
    }



    //Nos muestra los departamentos disponibles según la OU princpal que pasa como parámetro
    [HttpPost]
    public IActionResult GetDepartamentos([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData != null && requestData.ContainsKey("ouPrincipal"))
        {
            string ouPrincipal = requestData["ouPrincipal"];
            var departamentos = _ouService.GetDepartamentos(ouPrincipal);
            return Json(departamentos);
        }

        return Json(new List<string>());
    }

    //Devuelve el lugar de envío correspondiente al departamento
    [HttpPost]
    public IActionResult GetLugarEnvio([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData != null && requestData.ContainsKey("departamento"))
        {
            string departamento = requestData["departamento"];
            var lugaresEnvio = _ouService.GetLugarEnvio(departamento);
            return Json(lugaresEnvio);
        }

        return Json(new List<string>());
    }


    //Función encargada de comvertir el username recibido de una vista en string y pasarlo a la función que lo busca en AD
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


    //Busca si existe el nombre de usuario en el directorio activo
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


    //Comprueba si el id del usaurio existe en el directorio activo
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



    //Comprueba si el número de teléfono del usuario existe en el directorio activo
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


    //Algoritmo para generar un nombre de usuario
    [HttpPost]
    public IActionResult GenerateUsername([FromBody] userInputModel userInput)
    {
        if (string.IsNullOrEmpty(userInput.Nombre) || string.IsNullOrEmpty(userInput.Apellido1) || string.IsNullOrEmpty(userInput.Apellido2))
        {
            return Json(new { success = true, username = "" });
        }

        try
        {
            // Normalizar y dividir los atributos
            string[] nombrePartes = userInput.Nombre.Trim().ToLower().Split(' ');
            string[] apellido1Partes = userInput.Apellido1.Trim().ToLower().Split(' ');
            string[] apellido2Partes = string.IsNullOrEmpty(userInput.Apellido2)
                ? new string[0]
                : userInput.Apellido2.Trim().ToLower().Split(' ');

            // Construcción de candidatos
            List<string> candidatos = new List<string>();

            // 1. Primera inicial del nombre, primer apellido completo, primera inicial del segundo apellido
            string candidato1 = $"{GetInicial(nombrePartes)}{GetCompleto(apellido1Partes)}{GetInicial(apellido2Partes)}";
            candidatos.Add(candidato1.Substring(0, Math.Min(12, candidato1.Length)));

            // 2. Nombre completo (primera palabra completa y las iniciales de las demás), primera inicial del primer apellido, primera inicial del segundo apellido 
            string candidato2 = $"{GetNombreCompuesto(nombrePartes)}{GetInicial(apellido1Partes)}{GetInicial(apellido2Partes)}";
            candidatos.Add(candidato2.Substring(0, Math.Min(12, candidato2.Length)));

            // 3. Primera inicial del nombre, primera inicial del primer apellido, segundo apellido completo
            string candidato3 = $"{GetInicial(nombrePartes)}{GetInicial(apellido1Partes)}{GetCompleto(apellido2Partes)}";
            candidatos.Add(candidato3.Substring(0, Math.Min(12, candidato3.Length)));

            // Verificar la existencia de nombres de usuario
            foreach (string candidato in candidatos)
            {
                if (!CheckUserInActiveDirectory(candidato))
                {
                    return Json(new { success = true, username = candidato });
                }
            }

            // Si no se encuentra un nombre único
            return Json(new { success = false, message = "No se pudo generar un nombre de usuario único." });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al generar el nombre de usuario: {ex.Message}" });
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


    //Función para obtener los grupos a los que pertenece el usuario
    [HttpPost]
    public IActionResult GetUserGroups([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("username"))
            return Json(new { success = false, message = "Usuario no especificado." });

        string username = ExtractUsername(requestData["username"]);

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



    //Función para modificar los grupos a los que pertenece el usuario
    [HttpPost]
    public IActionResult ModifyUserGroup([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null ||
            !requestData.ContainsKey("username") ||
            !requestData.ContainsKey("group") ||
            !requestData.ContainsKey("action"))
        {
            return Json(new { success = false, message = "Datos insuficientes para modificar el grupo." });
        }

        // Extraer solo el sAMAccountName (por ejemplo, "jperez" de "Juan Pérez (jperez)")
        string input = requestData["username"];
        string username = ExtractUsername(input);
        if (string.IsNullOrEmpty(username))
            return Json(new { success = false, message = "El formato del usuario seleccionado no es válido." });

        string group = requestData["group"];
        string action = requestData["action"];

        try
        {
            // Buscar el grupo usando el método que ya tienes implementado
            DirectoryEntry groupEntry = FindGroupByName(group);
            if (groupEntry == null)
                return Json(new { success = false, message = $"Grupo {group} no encontrado en el dominio." });

            // Usamos el PrincipalContext con el dominio explícito para evitar problemas
            using (var context = new PrincipalContext(ContextType.Domain, "aytosa.inet"))
            using (var user = UserPrincipal.FindByIdentity(context, username))
            {
                if (user == null)
                    return Json(new { success = false, message = "Usuario no encontrado en Active Directory." });

                using (var userEntry = (DirectoryEntry)user.GetUnderlyingObject())
                {
                    // Obtenemos el distinguishedName del usuario (ejemplo: "CN=Juan Pérez,OU=Usuarios,DC=aytosa,DC=inet")
                    string userDN = userEntry.Properties["distinguishedName"].Value.ToString();

                    if (action == "add")
                    {
                        // Si el usuario no es miembro ya, se agrega
                        if (!groupEntry.Properties["member"].Contains(userDN))
                        {
                            groupEntry.Properties["member"].Add(userDN);
                            groupEntry.CommitChanges();
                        }
                    }
                    else if (action == "remove")
                    {
                        // Si el usuario es miembro, se elimina
                        if (groupEntry.Properties["member"].Contains(userDN))
                        {
                            groupEntry.Properties["member"].Remove(userDN);
                            groupEntry.CommitChanges();
                        }
                    }
                }
            }

            return Json(new { success = true, message = $"Grupo modificado correctamente: {action}." });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al modificar grupo: {ex.Message}" });
        }
    }





    //Función para modificar la OU a la que corresponde el usuario
    [HttpPost]
    public IActionResult ModifyUserOU([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null ||
            !requestData.ContainsKey("username") ||
            !requestData.ContainsKey("ouPrincipal") ||
            !requestData.ContainsKey("ouSecundaria") ||
            !requestData.ContainsKey("departamento") ||
            !requestData.ContainsKey("lugarEnvio"))
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
                        // Obtener la OU actual del usuario
                        string currentDistinguishedName = userEntry.Properties["distinguishedName"].Value.ToString();

                        // Mover usuario a la nueva OU
                        using (DirectoryEntry newOUEntry = new DirectoryEntry(newLDAPPath))
                        {
                            userEntry.MoveTo(newOUEntry);
                        }

                        // Actualizar atributos del usuario
                        userEntry.Properties["physicalDeliveryOfficeName"].Value = newDepartamento;
                        userEntry.Properties["streetAddress"].Value = newLugarEnvio; // Campo alternativo para "Lugar de Envío"

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



    //Función par obtener los datos del usuario (creo que solo los grupos)
    [HttpPost]
    public IActionResult GetUserDetails([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("username"))
            return Json(new { success = false, message = "Usuario no especificado." });

        // Extraer solo el sAMAccountName (por ejemplo, "jperez" de "Juan Pérez (jperez)")
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
                        // Extraemos la OU a partir del distinguishedName
                        string distinguishedName = userEntry.Properties["distinguishedName"].Value.ToString();
                        currentOU = ExtractOUFromDN(distinguishedName);
                        // Obtenemos solo los grupos directos a los que pertenece el usuario (propiedad memberOf)
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

    //Extrae la OU del distinguised name
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

    
    //Extrae los grupos a los que pertenece el usuario. Método para llamarse desde el propio controlador
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

    //Extrae el common name desde el distinguised name
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


    //Método para ejecutar archivos powershell
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
