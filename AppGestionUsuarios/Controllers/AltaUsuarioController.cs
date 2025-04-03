using Microsoft.AspNetCore.Mvc;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.DirectoryServices;
using Microsoft.AspNetCore.Authorization;
using System.Globalization;
using System.Text;
using System.Management.Automation;


/*En esta clase encontramos todos los métodos que son concretos del alta de usuario*/
/*En el caso de métodos que puedan usar otros menús, se almacenan en el apartado de gestión de usuarios*/

[Authorize]
public class AltaUsuarioController : Controller
{
    private readonly OUService _ouService;

    public AltaUsuarioController(OUService ouService)
    {
        _ouService = ouService;
    }

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
        public string LugarEnvio { get; set; }
        public string Dni { get; set; } 
        public string FechaCaducidadOp { get; set; }
        public DateTime FechaCaducidad { get; set; }
        public string Cuota { get; set; }
        public List<string> Grupos { get; set; }
    }

    [HttpGet]
    public IActionResult AltaUsuario()
    {
        try
        {
            // Cargar OUs principales desde el Active Directory
            ViewBag.OUPrincipales = GetOUPrincipalesFromAD();

            // Cargar grupos del Active Directory
            var grupos = GetGruposFromAD();
            ViewBag.GruposAD = grupos.OrderBy(g => g).ToList();

            // Otros datos necesarios desde el servicio
            ViewBag.PortalEmpleado = _ouService.GetPortalEmpleado();
            ViewBag.Cuota = _ouService.GetCuota();

            return View("AltaUsuario");
        }
        catch (Exception ex)
        {
            // Lanzar la excepción para que el middleware la maneje
            throw new Exception("Error al cargar la página de alta de usuario: " + ex.Message, ex);
        }
    }

    private List<string> GetGruposFromAD()
    {
        var grupos = new List<string>();

        try
        {
            using (var entry = new DirectoryEntry("LDAP://DC=aytosa,DC=inet"))
            {
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = "(objectClass=group)";
                    searcher.PropertiesToLoad.Add("cn");
                    searcher.SearchScope = SearchScope.Subtree;
                    searcher.PageSize = 1000;

                    foreach (SearchResult result in searcher.FindAll())
                    {
                        if (result.Properties.Contains("cn"))
                        {
                            grupos.Add(result.Properties["cn"][0].ToString());
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception("Error al obtener los grupos del Active Directory: " + ex.Message, ex);
        }

        return grupos;
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

            // Construir el path LDAP para la OU principal seleccionada
            string ldapPath = $"LDAP://OU={ouPrincipal},OU=AREAS,DC=aytosa,DC=inet";

            using (var rootEntry = new DirectoryEntry(ldapPath))
            {
                // Buscar las sub-OUs dentro de la OU principal
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
            return Json(ouSecundarias);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error al obtener las OU secundarias para {ouPrincipal}: {ex.Message}", ex);
        }
    }

    [HttpPost]
    public IActionResult GetDepartamento([FromBody] Dictionary<string, string> requestData)
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
                // Si hay OU secundaria, buscamos el departamento en la OU secundaria
                ldapPath = $"LDAP://OU={ouSecundaria},OU={ouPrincipal},OU=AREAS,DC=aytosa,DC=inet";
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
                // Si hay OU secundaria, buscamos el lugar de envío en la OU secundaria
                ldapPath = $"LDAP://OU={ouSecundaria},OU={ouPrincipal},OU=AREAS,DC=aytosa,DC=inet";
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


    //Función propia para crear el usuario
    [HttpPost]
    public IActionResult CreateUser([FromBody] UserModelAltaUsuario user)
    {
        // Validar si los datos se recibieron correctamente
        if (user == null)
        {
            return Json(new { success = false, message = "No se recibieron datos válidos." });
        }

        // Validar los campos obligatorios (Dni ahora es obligatorio)
        if (string.IsNullOrEmpty(user.Nombre) || string.IsNullOrEmpty(user.Apellido1) ||
            string.IsNullOrEmpty(user.Username) || string.IsNullOrEmpty(user.OUPrincipal) ||
            string.IsNullOrEmpty(user.Departamento) || string.IsNullOrEmpty(user.FechaCaducidadOp) ||
            string.IsNullOrEmpty(user.Dni))
        {
            return Json(new { success = false, message = "Faltan campos obligatorios." });
        }

        try
        {
            // Convertir nombre y apellidos a mayúsculas y eliminar acentos
            string nombreUpper = RemoveAccents(user.Nombre).ToUpperInvariant();
            string apellido1Upper = RemoveAccents(user.Apellido1).ToUpperInvariant();
            string apellido2Upper = string.IsNullOrEmpty(user.Apellido2) ? "" : RemoveAccents(user.Apellido2).ToUpperInvariant();

            // Conformar el nombre completo
            string displayName = $"{nombreUpper} {apellido1Upper} {apellido2Upper}".Trim();

            // Construir el path LDAP (usar la OU principal directamente si no hay OU secundaria)
            string ldapPath;
            if (!string.IsNullOrEmpty(user.OUSecundaria))
            {
                // Si hay OU secundaria, creamos el usuario en la OU secundaria
                ldapPath = $"LDAP://OU={user.OUSecundaria},OU={user.OUPrincipal},OU=AREAS,DC=aytosa,DC=inet";
            }
            else
            {
                // Si no hay OU secundaria, creamos el usuario directamente en la OU principal
                ldapPath = $"LDAP://OU={user.OUPrincipal},OU=AREAS,DC=aytosa,DC=inet";
            }

            using (DirectoryEntry ouEntry = new DirectoryEntry(ldapPath))
            {
                if (ouEntry == null)
                {
                    return Json(new { success = false, message = "No se pudo conectar a la OU especificada." });
                }

                // Crear un nuevo usuario
                DirectoryEntry newUser = null;

                try
                {
                    newUser = ouEntry.Children.Add($"CN={displayName}", "user");

                    // Establecer atributos básicos del usuario
                    newUser.Properties["givenName"].Value = user.Nombre;
                    newUser.Properties["sn"].Value = user.Apellido1 + " " + user.Apellido2;
                    newUser.Properties["sAMAccountName"].Value = user.Username;
                    newUser.Properties["userPrincipalName"].Value = $"{user.Username}@aytosa.inet";
                    newUser.Properties["displayName"].Value = displayName;
                    // Solo asignar NFuncionario si no está vacío
                    if (!string.IsNullOrEmpty(user.NFuncionario))
                    {
                        newUser.Properties["employeeNumber"].Value = user.NFuncionario;
                    }
                    // Solo asignar NTelefono si no está vacío
                    if (!string.IsNullOrEmpty(user.NTelefono))
                    {
                        newUser.Properties["telephoneNumber"].Value = user.NTelefono;
                    }
                    newUser.Properties["physicalDeliveryOfficeName"].Value = user.Departamento;
                    newUser.Properties["l"].Value = user.LugarEnvio; // Establecer el atributo 'city' (l)
                    newUser.Properties["employeeID"].Value = user.Dni; // Establecer el atributo 'employeeID'

                    if (user.FechaCaducidadOp == "si")
                    {
                        if (user.FechaCaducidad <= DateTime.Now)
                        {
                            return Json(new { success = false, message = "La fecha de caducidad debe ser una fecha futura." });
                        }

                        try
                        {
                            long accountExpires = user.FechaCaducidad.ToFileTime();
                            newUser.Properties["accountExpires"].Value = accountExpires.ToString();
                        }
                        catch (ArgumentOutOfRangeException ex)
                        {
                            return Json(new { success = false, message = $"Error al convertir la fecha. {ex.Message}" });
                        }
                    }

                    newUser.CommitChanges();

                    // Configurar contraseña y activar cuenta
                    newUser.Invoke("SetPassword", new object[] { "Temporal2024" });
                    newUser.Properties["userAccountControl"].Value = 0x200;
                    newUser.Properties["pwdLastSet"].Value = 0;

                    newUser.CommitChanges();

                    // Añadir al usuario a los grupos seleccionados
                    if (user.Grupos != null && user.Grupos.Any())
                    {
                        foreach (string grupo in user.Grupos)
                        {
                            DirectoryEntry groupEntry = FindGroupByName(grupo);
                            if (groupEntry != null)
                            {
                                try
                                {
                                    // Agregar el usuario al grupo
                                    groupEntry.Invoke("Add", new object[] { newUser.Path });
                                    groupEntry.CommitChanges();
                                }
                                catch (Exception ex)
                                {
                                    // Ignorar errores al agregar a grupos (puedes registrar el error si lo deseas)
                                }
                                finally
                                {
                                    groupEntry.Dispose();
                                }
                            }
                            else
                            {
                                return Json(new { success = false, message = $"Grupo {grupo} no encontrado en el dominio." });
                            }
                        }
                    }

                    newUser.CommitChanges();
                }
                catch (Exception ex)
                {
                    return Json(new { success = false, message = $"Error al crear el usuario: {ex.Message}" });
                }
                finally
                {
                    newUser?.Dispose();
                }

                return Json(new { success = true, message = "Usuario creado exitosamente y añadido a los grupos seleccionados." });
            }
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al crear el usuario: {ex.Message}" });
        }
    }

    // Nuevo método para verificar si el DNI ya existe en el Active Directory
    [HttpPost]
    public IActionResult CheckDniExists([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("dni"))
        {
            return Json(new { success = false, message = "DNI no especificado." });
        }

        string dni = requestData["dni"];
        if (string.IsNullOrEmpty(dni))
        {
            return Json(new { success = false, message = "El DNI no puede estar vacío." });
        }

        try
        {
            using (var entry = new DirectoryEntry("LDAP://DC=aytosa,DC=inet"))
            {
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(&(objectClass=user)(employeeID={dni}))";
                    searcher.SearchScope = SearchScope.Subtree;
                    searcher.PropertiesToLoad.Add("employeeID");

                    var result = searcher.FindOne();
                    if (result != null)
                    {
                        return Json(new { success = true, exists = true });
                    }
                    else
                    {
                        return Json(new { success = true, exists = false });
                    }
                }
            }
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al verificar el DNI: {ex.Message}" });
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
}
