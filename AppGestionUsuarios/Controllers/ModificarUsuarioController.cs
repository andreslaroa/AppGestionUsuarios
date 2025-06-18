using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace AppGestionUsuarios.Controllers
{
    [Authorize]
    public class ModificarUsuarioController : Controller
    {
        // Cadena base para el LDAP
        private const string DomainPath = "LDAP://DC=aytosa,DC=inet";

        private readonly IConfiguration _config;

        public ModificarUsuarioController(IConfiguration config)
        {
            _config = config;
        }

        // GET: /ModificarUsuario
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

        // POST: /ModificarUsuario/GetOUSecundarias
        [HttpPost]
        public JsonResult GetOUSecundarias([FromBody] Dictionary<string, string> requestData)
        {
            if (requestData == null || !requestData.ContainsKey("ouPrincipal"))
                return Json(new List<string>());

            var ouPrincipal = requestData["ouPrincipal"];
            // Ruta donde buscar las OU secundarias
            var ldapPath = $"LDAP://OU=Usuarios,OU={ouPrincipal},{_config["ActiveDirectory:DomainBase"]}";

            try
            {
                using var entry = new DirectoryEntry(ldapPath);
                using var searcher = new DirectorySearcher(entry)
                {
                    Filter = "(objectClass=organizationalUnit)",
                    SearchScope = SearchScope.OneLevel
                };

                var list = searcher.FindAll()
                                   .Cast<SearchResult>()
                                   .Select(r => r.Properties["ou"][0].ToString())
                                   .ToList();
                return Json(list);
            }
            catch
            {
                return Json(new List<string>());
            }
        }

        // POST: /ModificarUsuario/GetDepartamento
        [HttpPost]
        public JsonResult GetDepartamento([FromBody] Dictionary<string, string> requestData)
        {
            if (requestData == null || !requestData.ContainsKey("ouPrincipal"))
                return Json(new { success = false, message = "Falta ouPrincipal." });

            var ouPrincipal = requestData["ouPrincipal"];
            var ouSecundaria = requestData.GetValueOrDefault("ouSecundaria");
            // El path cambia si hay OU secundaria o no
            var ldapPath = !string.IsNullOrEmpty(ouSecundaria)
                ? $"LDAP://OU={ouSecundaria},OU=Usuarios,OU={ouPrincipal},{_config["ActiveDirectory:DomainBase"]}"
                : $"LDAP://OU=Usuarios,OU={ouPrincipal},{_config["ActiveDirectory:DomainBase"]}";

            try
            {
                using var entry = new DirectoryEntry(ldapPath);
                var dep = entry.Properties[_config["GroupInformation:DepartmentAttr"]]?.Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(dep))
                    return Json(new { success = false, message = "Departamento no definido." });
                return Json(new { success = true, departamento = dep });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // POST: /ModificarUsuario/GetLugarEnvio
        [HttpPost]
        public JsonResult GetLugarEnvio([FromBody] Dictionary<string, string> requestData)
        {
            if (requestData == null || !requestData.ContainsKey("ouPrincipal"))
                return Json(new { success = false, message = "Falta ouPrincipal." });

            var ouPrincipal = requestData["ouPrincipal"];
            var ouSecundaria = requestData.GetValueOrDefault("ouSecundaria");
            var ldapPath = !string.IsNullOrEmpty(ouSecundaria)
                ? $"LDAP://OU={ouSecundaria},OU=Usuarios,OU={ouPrincipal},{_config["ActiveDirectory:DomainBase"]}"
                : $"LDAP://OU=Usuarios,OU={ouPrincipal},{_config["ActiveDirectory:DomainBase"]}";

            try
            {
                using var entry = new DirectoryEntry(ldapPath);
                var lugar = entry.Properties[_config["GroupInformation:SendPlaceAttr"]]?.Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(lugar))
                    return Json(new { success = false, message = "Lugar de envío no definido." });
                return Json(new { success = true, lugarEnvio = lugar });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // POST: /ModificarUsuario/ModifyUserOU
        [HttpPost]
        public JsonResult ModifyUserOU([FromBody] Dictionary<string, string> requestData)
        {
            // 1) Validación básica
            if (requestData == null
             || !requestData.ContainsKey("username")
             || !requestData.ContainsKey("ouPrincipal")
             || !requestData.ContainsKey("departamento"))
            {
                return Json(new { success = false, message = "Faltan datos obligatorios." });
            }

            var rawUser = requestData["username"];
            var username = ExtractUsername(rawUser);
            var ouPrincipal = requestData["ouPrincipal"];
            var ouSecundaria = requestData.GetValueOrDefault("ouSecundaria");
            var departamento = requestData["departamento"];

            if (username == null)
                return Json(new { success = false, message = "Formato de usuario inválido." });

            // 2) Construir nuevo path LDAP
            var newPath = !string.IsNullOrEmpty(ouSecundaria)
                ? $"LDAP://OU={ouSecundaria},OU=Usuarios,OU={ouPrincipal},{_config["ActiveDirectory:DomainBase"]}"
                : $"LDAP://OU=Usuarios,OU={ouPrincipal},{_config["ActiveDirectory:DomainBase"]}";

            try
            {
                using var ctx = new PrincipalContext(ContextType.Domain);
                using var user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, username);
                if (user == null)
                    return Json(new { success = false, message = $"Usuario '{username}' no encontrado en AD." });

                var de = (DirectoryEntry)user.GetUnderlyingObject();

                // 3) Mover a la nueva OU
                de.MoveTo(new DirectoryEntry(newPath));

                // 4) Actualizar departamento
                de.Properties["physicalDeliveryOfficeName"].Value = departamento;
                de.CommitChanges();

                return Json(new { success = true, message = "OU y departamento actualizados correctamente." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Error al modificar OU: {ex.Message}" });
            }
        }

       
        /// <summary>
        /// Extrae el sAMAccountName de un string tipo "DisplayName (sam)".
        /// </summary>
        private string ExtractUsername(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return null;
            var i1 = input.LastIndexOf('(');
            var i2 = input.LastIndexOf(')');
            if (i1 < 0 || i2 <= i1) return null;
            return input.Substring(i1 + 1, i2 - i1 - 1);
        }

        /// <summary>
        /// Obtiene la lista de OUs principales debajo de OU=AREAS.
        /// </summary>
        private List<string> GetOUPrincipales()
        {
            var list = new List<string>();
            try
            {
                string ldapPath = $"LDAP://{_config["ActiveDirectory:DomainBase"]}";
                using var entry = new DirectoryEntry(ldapPath);
                using var searcher = new DirectorySearcher(entry)
                {
                    Filter = "(objectClass=organizationalUnit)",
                    SearchScope = SearchScope.OneLevel
                };

                foreach (SearchResult r in searcher.FindAll())
                {
                    list.Add(r.Properties["ou"][0].ToString());
                }
            }
            catch
            {
                // ignorar errores: devolvemos lista vacía
            }
            return list;
        }

        /// <summary>
        /// Obtiene todos los grupos de AD (por SamAccountName).
        /// </summary>
        public List<string> GetGruposFromAD()
        {
            var grupos = new List<string>();
            string baseLdap = _config["ActiveDirectory:BaseLdapPrefix"]
                            + _config["ActiveDirectory:DomainComponents"];

            using (var entry = new DirectoryEntry(baseLdap))
            {
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = "(objectClass=group)";
                    searcher.PropertiesToLoad.Add("cn");
                    searcher.SearchScope = SearchScope.Subtree;
                    searcher.PageSize = 1000;

                    foreach (System.DirectoryServices.SearchResult result in searcher.FindAll())
                        if (result.Properties.Contains("cn"))
                            grupos.Add(result.Properties["cn"][0].ToString());
                }
            }
            return grupos;
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
    }
}
