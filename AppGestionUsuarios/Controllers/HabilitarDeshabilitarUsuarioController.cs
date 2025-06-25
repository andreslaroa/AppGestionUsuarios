using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.DirectoryServices;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Win32.SafeHandles;

namespace AppGestionUsuarios.Controllers
{
    [Authorize]
    public class HabilitarDeshabilitarUsuarioController : Controller
    {
        private const string DomainPath = "LDAP://DC=aytosa,DC=inet";

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
        private readonly IConfiguration _config;
        public HabilitarDeshabilitarUsuarioController(IConfiguration config, IDataProtectionProvider dp)
        {
            _config = config;
            _protector = dp.CreateProtector("CredencialesProtector");

        }

        // GET: /HabilitarDeshabilitarUsuario
        [HttpGet]
        public IActionResult HabilitarDeshabilitarUsuario()
        {
            var usuarios = new List<string>();
            try
            {
                using var root = new DirectoryEntry(DomainPath);
                using var searcher = new DirectorySearcher(root)
                {
                    Filter = "(objectClass=user)",
                    SearchScope = SearchScope.Subtree,
                    PageSize = 1000
                };
                searcher.PropertiesToLoad.Add("displayName");
                searcher.PropertiesToLoad.Add("sAMAccountName");

                foreach (SearchResult result in searcher.FindAll())
                {
                    if (result.Properties.Contains("displayName") &&
                        result.Properties.Contains("sAMAccountName"))
                    {
                        string dn = result.Properties["displayName"][0].ToString();
                        string sam = result.Properties["sAMAccountName"][0].ToString();
                        usuarios.Add($"{dn} ({sam})");
                    }
                }

                ViewBag.Usuarios = usuarios.OrderBy(u => u).ToList();
            }
            catch (Exception ex)
            {
                // En caso de error devolvemos lista vacía y logueamos
                ViewBag.Usuarios = new List<string>();
                Console.WriteLine($"Error cargando usuarios: {ex.Message}");
            }

            return View();
        }

        // POST: /HabilitarDeshabilitarUsuario/ManageUserStatus
        [HttpPost]
        [Produces("application/json")]
        public IActionResult ManageUserStatus([FromBody] Dictionary<string, string> requestData)
        {
            // 1) Validación preliminar
            if (requestData == null
             || !requestData.ContainsKey("username")
             || !requestData.ContainsKey("action"))
            {
                return Json(new { success = false, message = "Faltan datos: 'username' y 'action' son obligatorios." });
            }

            // 2) Recuperar credenciales de sesión y dominio
            string adminUsername = HttpContext.Session.GetString("adminUser");
            var encryptedPass = HttpContext.Session.GetString("adminPassword");
            var adminPassword = _protector.Unprotect(encryptedPass);

            var domainName = _config["ActiveDirectory:DomainName"];
            if (string.IsNullOrWhiteSpace(domainName))
            {
                return Json(new { success = false, message = "Configuración incorrecta: falta ActiveDirectory:DomainName" });
            }

            // 3) Impersonación
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
            IActionResult finalResult = Json(new { success = false, message = "No se completó la operación." });

            // 4) Ejecutar bajo impersonación
            WindowsIdentity.RunImpersonated(safeToken, () =>
            {
                string rawInput = requestData["username"];
                string action = requestData["action"].ToLower();
                string username = ExtractUsername(rawInput);
                if (username == null)
                {
                    finalResult = Json(new { success = false, message = "Formato de usuario inválido." });
                    return;
                }

                try
                {
                    using var root = new DirectoryEntry(domainName);
                    using var searcher = new DirectorySearcher(root)
                    {
                        Filter = $"(&(objectClass=user)(sAMAccountName={username}))",
                        SearchScope = SearchScope.Subtree
                    };
                    searcher.PropertiesToLoad.Add("userAccountControl");

                    var result = searcher.FindOne();
                    if (result == null)
                    {
                        finalResult = Json(new { success = false, message = $"Usuario '{username}' no encontrado." });
                        return;
                    }

                    using var userEntry = result.GetDirectoryEntry();
                    int uac = (int)userEntry.Properties["userAccountControl"].Value;

                    if (action == "enable")
                        uac &= ~0x2;   // quitar flag DISABLED
                    else if (action == "disable")
                        uac |= 0x2;    // añadir flag DISABLED
                    else
                    {
                        finalResult = Json(new { success = false, message = "Acción no válida: use 'enable' o 'disable'." });
                        return;
                    }

                    userEntry.Properties["userAccountControl"].Value = uac;
                    userEntry.CommitChanges();

                    string verb = action == "enable" ? "habilitado" : "deshabilitado";
                    finalResult = Json(new { success = true, message = $"Usuario '{username}' {verb} correctamente." });
                }
                catch (Exception ex)
                {
                    finalResult = Json(new { success = false, message = $"Error al {requestData["action"]} usuario: {ex.Message}" });
                }
            });

            // 5) Cerrar handle
            CloseHandle(userToken);

            // 6) Retornar siempre JSON
            return finalResult;
        }


        // POST: /HabilitarDeshabilitarUsuario/GetUserGroups
        [HttpPost]
        public IActionResult GetUserGroups([FromBody] Dictionary<string, string> requestData)
        {
            if (requestData == null || !requestData.ContainsKey("username"))
                return Json(new { success = false, message = "Se requiere 'username'." });

            string rawInput = requestData["username"];
            string username = ExtractUsername(rawInput);
            if (username == null)
                return Json(new { success = false, message = "Formato de usuario inválido." });

            try
            {
                using var root = new DirectoryEntry(DomainPath);
                using var searcher = new DirectorySearcher(root)
                {
                    Filter = $"(&(objectClass=user)(sAMAccountName={username}))",
                    SearchScope = SearchScope.Subtree
                };
                searcher.PropertiesToLoad.Add("memberOf");

                var result = searcher.FindOne();
                if (result == null)
                    return Json(new { success = false, message = $"Usuario '{username}' no encontrado." });

                var groups = new List<string>();
                if (result.Properties.Contains("memberOf"))
                {
                    foreach (var dn in result.Properties["memberOf"])
                    {
                        string cn = ExtractGroupName(dn.ToString());
                        if (cn != null) groups.Add(cn);
                    }
                }

                return Json(new { success = true, groups });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Error obteniendo grupos: {ex.Message}" });
            }
        }

        // --- Helpers privados ---

        /// <summary>
        /// Extrae el sAMAccountName de un string como "Display Name (sam)".
        /// </summary>
        private string ExtractUsername(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return null;
            int i1 = input.LastIndexOf('(');
            int i2 = input.LastIndexOf(')');
            if (i1 < 0 || i2 <= i1) return null;
            return input.Substring(i1 + 1, i2 - i1 - 1).Trim();
        }

        /// <summary>
        /// Busca un DirectoryEntry de grupo por CN.
        /// </summary>
        private DirectoryEntry FindGroupByName(string groupName)
        {
            try
            {
                using var root = new DirectoryEntry(DomainPath);
                using var searcher = new DirectorySearcher(root)
                {
                    Filter = $"(&(objectClass=group)(cn={groupName}))",
                    SearchScope = SearchScope.Subtree
                };
                searcher.PropertiesToLoad.Add("distinguishedName");

                var result = searcher.FindOne();
                return result?.GetDirectoryEntry();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// De un distinguishedName devuelve solo el valor CN.
        /// </summary>
        private string ExtractGroupName(string distinguishedName)
        {
            if (string.IsNullOrWhiteSpace(distinguishedName)) return null;
            var parts = distinguishedName.Split(',');
            foreach (var p in parts)
            {
                if (p.StartsWith("CN=", StringComparison.OrdinalIgnoreCase))
                    return p.Substring(3);
            }
            return null;
        }
    }
}
