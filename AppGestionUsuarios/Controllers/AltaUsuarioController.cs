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
using System.DirectoryServices.AccountManagement;
using static GestionUsuariosController;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Runtime.InteropServices; // Para COM


/*En esta clase encontramos todos los métodos que son concretos del alta de usuario*/
/*En el caso de métodos que puedan usar otros menús, se almacenan en el apartado de gestión de usuarios*/

[Authorize]
public class AltaUsuarioController : Controller
{

    private readonly IConfiguration _config;

    public AltaUsuarioController(IConfiguration config)
    {
        _config = config;
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
        public string FechaCaducidadOp { get; set; }
        public DateTime FechaCaducidad { get; set; }
        public string Cuota { get; set; }
        public List<string> Grupos { get; set; }
        public string NumeroLargoFijo { get; set; }
        public string ExtensionMovil { get; set; }
        public string NumeroLargoMovil { get; set; }
        public string TarjetaIdentificativa { get; set; }
        public string DNI { get; set; }
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

            // Mover la lógica de GetPortalEmpleado y GetCuota directamente aquí
            ViewBag.PortalEmpleado = new List<string> { "GA_R_PORTALDELEMPLEADO" };
            ViewBag.Cuota = new List<string> { "500MB", "1GB", "2GB" };

            return View("AltaUsuario");
        }
        catch (Exception ex)
        {
            throw new Exception("Error al cargar la página de alta de usuario: " + ex.Message, ex);
        }
    }

    private List<string> GetGruposFromAD()
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

                foreach (SearchResult result in searcher.FindAll())
                    if (result.Properties.Contains("cn"))
                        grupos.Add(result.Properties["cn"][0].ToString());
            }
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
    public IActionResult CreateUser([FromBody] UserModelAltaUsuario user)
    {
        // 1) Leemos la configuración de fichero
        string serverName = _config["FsConfig:ServerName"];
        string folderPathBase = _config["FsConfig:ShareBase"];
        string quotaPathBase = _config["FsConfig:QuotaPathBase"];

        // 2) Construimos rutas sin hard-codear ningún literal
        string uncUserFolder = Path.Combine(folderPathBase, user.Username);
        string localQuotaPath = Path.Combine(quotaPathBase, user.Username);


        // Validar si los datos se recibieron correctamente
        if (user == null)
        {
            return Json(new { success = false, message = "No se recibieron datos válidos." });
        }

        // Validar los campos obligatorios (los nuevos campos son opcionales)
        if (string.IsNullOrEmpty(user.Nombre) || string.IsNullOrEmpty(user.Apellido1) ||
            string.IsNullOrEmpty(user.Username) || string.IsNullOrEmpty(user.OUPrincipal) ||
            string.IsNullOrEmpty(user.Departamento) || string.IsNullOrEmpty(user.FechaCaducidadOp))
        {
            return Json(new { success = false, message = "Faltan campos obligatorios." });
        }

        // Lista para almacenar los errores que ocurran durante el proceso
        List<string> errors = new List<string>();
        bool userCreated = false; // Bandera para indicar si el usuario se creó
        bool addedToDepartmentGroup = false; // Bandera para indicar si se añadió al grupo del departamento

        string grupoDepartamento = null;

        try
        {
            // Convertir nombre y apellidos a mayúsculas y eliminar acentos
            string nombreUpper = RemoveAccents(user.Nombre).ToUpperInvariant();
            string apellido1Upper = RemoveAccents(user.Apellido1).ToUpperInvariant();
            string apellido2Upper = string.IsNullOrEmpty(user.Apellido2) ? "" : RemoveAccents(user.Apellido2).ToUpperInvariant();

            // Conformar el nombre completo
            string displayName = $"{nombreUpper} {apellido1Upper} {apellido2Upper}".Trim();

            // Construir el path LDAP
            string ldapPath;
            string ouPath; // Para obtener los atributos de la OU más inmediata
            if (!string.IsNullOrEmpty(user.OUSecundaria))
            {
                ldapPath = $"LDAP://OU={user.OUSecundaria},OU=Usuarios y Grupos,OU={user.OUPrincipal},OU=AREAS,DC=aytosa,DC=inet";
                ouPath = $"LDAP://OU={user.OUSecundaria},OU=Usuarios y Grupos,OU={user.OUPrincipal},OU=AREAS,DC=aytosa,DC=inet";
                errors.Add($"Usando OU secundaria: Path = {ouPath}");
            }
            else
            {
                ldapPath = $"LDAP://OU=Usuarios y Grupos,OU={user.OUPrincipal},OU=AREAS,DC=aytosa,DC=inet";
                ouPath = $"LDAP://OU={user.OUPrincipal},OU=AREAS,DC=aytosa,DC=inet";
                errors.Add($"Usando OU principal: Path = {ouPath}");
            }

            // Obtener los atributos "st", "description" y "street" de la OU más inmediata
            string estadoProvincia = null;
            string descripcion = null;

            try
            {
                using (var ouEntryForAttributes = new DirectoryEntry(ouPath))
                {
                    if (ouEntryForAttributes == null)
                    {
                        errors.Add("No se pudo conectar a la OU para obtener los atributos.");
                    }
                    else
                    {
                        // Obtener el atributo "st" (Estado o Provincia)
                        estadoProvincia = ouEntryForAttributes.Properties["st"]?.Value?.ToString();
                        if (string.IsNullOrEmpty(estadoProvincia))
                        {
                            estadoProvincia = "Sin estado o provincia";
                            errors.Add("Atributo 'st' no definido, usando valor predeterminado: 'Sin estado o provincia'.");
                        }
                        else
                        {
                            errors.Add($"Atributo 'st' encontrado: '{estadoProvincia}'.");
                        }

                        // Obtener el atributo "description" (Descripción)
                        descripcion = ouEntryForAttributes.Properties["description"]?.Value?.ToString();
                        if (string.IsNullOrEmpty(descripcion))
                        {
                            descripcion = "Sin descripción";
                            errors.Add("Atributo 'description' no definido, usando valor predeterminado: 'Sin descripción'.");
                        }
                        else
                        {
                            errors.Add($"Atributo 'description' encontrado: '{descripcion}'.");
                        }

                        // Obtener el atributo "street" (Calle), que será el nombre del grupo del departamento
                        grupoDepartamento = ouEntryForAttributes.Properties["street"]?.Value?.ToString();
                        if (string.IsNullOrEmpty(grupoDepartamento))
                        {
                            grupoDepartamento = null;
                            errors.Add($"El atributo 'street' no está definido en la OU (Path: {ouPath}).");
                        }
                        else
                        {
                            errors.Add($"Atributo 'street' encontrado: '{grupoDepartamento}' (Path: {ouPath}).");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                errors.Add($"Error al obtener los atributos de la OU: {ex.Message}");
            }

            using (DirectoryEntry ouEntry = new DirectoryEntry(ldapPath))
            {
                if (ouEntry == null)
                {
                    errors.Add("No se pudo conectar a la OU especificada para crear el usuario.");
                    return Json(new { success = false, message = "No se pudo conectar a la OU para crear el usuario." });
                }

                DirectoryEntry newUser = null;

                try
                {
                    // Crear el usuario
                    newUser = ouEntry.Children.Add($"CN={displayName}", "user");
                    errors.Add($"Usuario creado en la OU: CN={displayName}");

                    // Establecer atributos básicos del usuario
                    try
                    {
                        newUser.Properties["givenName"].Value = user.Nombre;
                        newUser.Properties["sn"].Value = user.Apellido1 + " " + user.Apellido2;
                        newUser.Properties["sAMAccountName"].Value = user.Username;
                        newUser.Properties["userPrincipalName"].Value = $"{user.Username}@aytosa.inet";
                        newUser.Properties["displayName"].Value = displayName;
                        newUser.Properties["department"].Value = estadoProvincia;
                        newUser.Properties["division"].Value = descripcion;
                        errors.Add("Atributos básicos del usuario establecidos correctamente.");
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Error al establecer los atributos básicos del usuario: {ex.Message}");
                    }

                    // Asignar campos opcionales
                    try
                    {
                        if (!string.IsNullOrEmpty(user.NFuncionario))
                        {
                            newUser.Properties["description"].Value = user.NFuncionario;
                            errors.Add($"Atributo 'description' (NFuncionario) establecido: {user.NFuncionario}");
                        }
                        if (!string.IsNullOrEmpty(user.NTelefono))
                        {
                            newUser.Properties["telephoneNumber"].Value = user.NTelefono;
                            errors.Add($"Atributo 'telephoneNumber' establecido: {user.NTelefono}");
                        }
                        if (!string.IsNullOrEmpty(user.NumeroLargoFijo))
                        {
                            newUser.Properties["otherTelephone"].Value = user.NumeroLargoFijo;
                            errors.Add($"Atributo 'otherTelephone' establecido: {user.NumeroLargoFijo}");
                        }
                        if (!string.IsNullOrEmpty(user.ExtensionMovil))
                        {
                            newUser.Properties["mobile"].Value = user.ExtensionMovil;
                            errors.Add($"Atributo 'mobile' establecido: {user.ExtensionMovil}");
                        }
                        if (!string.IsNullOrEmpty(user.NumeroLargoMovil))
                        {
                            newUser.Properties["otherMobile"].Value = user.NumeroLargoMovil;
                            errors.Add($"Atributo 'otherMobile' establecido: {user.NumeroLargoMovil}");
                        }
                        if (!string.IsNullOrEmpty(user.TarjetaIdentificativa))
                        {
                            newUser.Properties["serialNumber"].Value = user.TarjetaIdentificativa;
                            errors.Add($"Atributo 'serialNumber' establecido: {user.TarjetaIdentificativa}");
                        }
                        if (!string.IsNullOrEmpty(user.DNI))
                        {
                            newUser.Properties["employeeID"].Value = user.DNI;
                            errors.Add($"Atributo 'employeeID' establecido: {user.DNI}");
                        }
                        newUser.Properties["physicalDeliveryOfficeName"].Value = user.Departamento;
                        newUser.Properties["l"].Value = user.LugarEnvio;
                        errors.Add("Atributos opcionales del usuario establecidos correctamente.");
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Error al establecer los atributos opcionales del usuario: {ex.Message}");
                    }

                    // Establecer la fecha de caducidad si aplica
                    if (user.FechaCaducidadOp == "si")
                    {
                        try
                        {
                            if (user.FechaCaducidad <= DateTime.Now)
                            {
                                errors.Add("La fecha de caducidad debe ser una fecha futura.");
                            }
                            else
                            {
                                long accountExpires = user.FechaCaducidad.ToFileTime();
                                newUser.Properties["accountExpires"].Value = accountExpires.ToString();
                                errors.Add($"Fecha de caducidad establecida: {user.FechaCaducidad}");
                            }
                        }
                        catch (Exception ex)
                        {
                            errors.Add($"Error al establecer la fecha de caducidad: {ex.Message}");
                        }
                    }

                    // Guardar los cambios iniciales del usuario
                    try
                    {
                        newUser.CommitChanges();
                        errors.Add("Cambios iniciales del usuario guardados correctamente.");
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Error al guardar los cambios iniciales del usuario: {ex.Message}");
                        throw; // Si falla aquí, no podemos continuar
                    }

                    // Configurar contraseña y activar cuenta
                    try
                    {
                        newUser.Invoke("SetPassword", new object[] { "Temporal2024" });
                        newUser.Properties["userAccountControl"].Value = 0x200;
                        newUser.Properties["pwdLastSet"].Value = 0;
                        newUser.CommitChanges();
                        userCreated = true; // Marcamos que el usuario se creó exitosamente
                        errors.Add("Contraseña configurada y cuenta activada correctamente.");
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Error al configurar la contraseña o activar la cuenta: {ex.Message}");
                    }

                    // Añadir al usuario al grupo del departamento (obligatorio) basado en el atributo "street"
                    if (!string.IsNullOrEmpty(grupoDepartamento))
                    {
                        errors.Add($"Intentando añadir al usuario al grupo del departamento: '{grupoDepartamento}'");
                        DirectoryEntry groupEntry = FindGroupByName(grupoDepartamento);
                        if (groupEntry != null)
                        {
                            try
                            {
                                groupEntry.Invoke("Add", new object[] { newUser.Path });
                                groupEntry.CommitChanges();
                                addedToDepartmentGroup = true;
                                errors.Add($"Usuario añadido exitosamente al grupo del departamento '{grupoDepartamento}'.");
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"Error al añadir al grupo del departamento '{grupoDepartamento}': {ex.Message}");
                            }
                            finally
                            {
                                groupEntry?.Dispose();
                            }
                        }
                        else
                        {
                            errors.Add($"El grupo del departamento '{grupoDepartamento}' no existe en el Directorio Activo.");
                        }
                    }
                    else
                    {
                        errors.Add("No se pudo añadir al grupo del departamento porque el atributo 'street' no está definido.");
                    }

                    // Añadir al usuario a los grupos seleccionados (los que vienen del formulario)
                    if (user.Grupos != null && user.Grupos.Any())
                    {
                        errors.Add("Añadiendo usuario a los grupos seleccionados...");
                        foreach (string grupo in user.Grupos)
                        {
                            DirectoryEntry groupEntry = FindGroupByName(grupo);
                            if (groupEntry != null)
                            {
                                try
                                {
                                    groupEntry.Invoke("Add", new object[] { newUser.Path });
                                    groupEntry.CommitChanges();
                                    errors.Add($"Usuario añadido exitosamente al grupo '{grupo}'.");
                                }
                                catch (Exception ex)
                                {
                                    errors.Add($"Error al añadir al grupo '{grupo}': {ex.Message}");
                                }
                                finally
                                {
                                    groupEntry?.Dispose();
                                }
                            }
                            else
                            {
                                errors.Add($"Grupo '{grupo}' no encontrado en el dominio.");
                            }
                        }
                    }
                    else
                    {
                        errors.Add("No se seleccionaron grupos adicionales para el usuario.");
                    }

                    // Guardar los cambios finales del usuario en Active Directory
                    try
                    {
                        newUser.CommitChanges();
                        errors.Add("Cambios finales del usuario guardados correctamente.");
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Error al guardar los cambios finales del usuario: {ex.Message}");
                    }

                    // Configuración de la carpeta personal y cuota
                    if (userCreated && user.OUPrincipal != "OAGER")
                    {
                        string folderPath = Path.Combine(folderPathBase, user.Username);
                        errors.Add($"Verificando existencia de la carpeta: {folderPath}");

                        if (!Directory.Exists(folderPath))
                        {
                            try
                            {
                                // 1) Crear la carpeta vía UNC
                                Directory.CreateDirectory(folderPath);
                                errors.Add($"Carpeta creada: {folderPath}");

                                // 2) NTFS: permisos sobre \\LEONARDO\Home\<user>
                                DirectoryInfo di = new DirectoryInfo(folderPath);
                                var ds = new DirectorySecurity();
                                // FullControl a las cuentas de administración
                                ds.AddAccessRule(new FileSystemAccessRule(
                                    new NTAccount("aytosa\\adm_fs"),
                                    FileSystemRights.FullControl,
                                    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                    PropagationFlags.None,
                                    AccessControlType.Allow));
                                ds.AddAccessRule(new FileSystemAccessRule(
                                    new NTAccount("aytosa\\adm_ds"),
                                    FileSystemRights.FullControl,
                                    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                    PropagationFlags.None,
                                    AccessControlType.Allow));
                                // Permisos del propio usuario
                                ds.AddAccessRule(new FileSystemAccessRule(
                                    new NTAccount($"aytosa\\{user.Username}"),
                                    FileSystemRights.ReadAndExecute | FileSystemRights.Write | FileSystemRights.DeleteSubdirectoriesAndFiles,
                                    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                    PropagationFlags.None,
                                    AccessControlType.Allow));

                                ds.AddAccessRule(new FileSystemAccessRule(
                                    new NTAccount("AYTOSA\\adm_andres"),
                                    FileSystemRights.FullControl,
                                    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                    PropagationFlags.None,
                                    AccessControlType.Allow));
                                // Protegemos herencia y quitamos permisos preexistentes
                                ds.SetAccessRuleProtection(true, true);
                                di.SetAccessControl(ds);

                                // Limpieza de reglas BUILTIN si lo necesitas
                                var updatedDs = di.GetAccessControl();
                                updatedDs.RemoveAccessRuleAll(new FileSystemAccessRule(
                                    new NTAccount("BUILTIN\\Administradores"),
                                    FileSystemRights.FullControl,
                                    AccessControlType.Allow));
                                updatedDs.RemoveAccessRuleAll(new FileSystemAccessRule(
                                    new NTAccount("BUILTIN\\Usuarios"),
                                    FileSystemRights.FullControl,
                                    AccessControlType.Allow));
                                di.SetAccessControl(updatedDs);

                                errors.Add("Permisos NTFS en Leonardo configurados correctamente.");
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"Error creando carpeta o permisos NTFS: {ex.Message}");
                            }

                            // 3) Configuración de cuota FSRM en C:\Home\<user>
                            try
                            {
                                string quota = user.Cuota ?? "1GB";
                                string template = $"HOME-{quota}";
                                string quotaFolder = Path.Combine(quotaPathBase, user.Username);

                                errors.Add($"Configurando cuota en {quotaFolder} con plantilla {template} sobre {serverName}");

                                // Instanciamos FSRM remoto
                                var fsrmType = Type.GetTypeFromProgID("Fsrm.FsrmQuotaManager", serverName);
                                if (fsrmType == null)
                                    throw new Exception($"No instanciable FsrmQuotaManager en {serverName}.");

                                dynamic qm = Activator.CreateInstance(fsrmType);
                                try
                                {
                                    dynamic existing = null;
                                    try { existing = qm.GetQuota(quotaFolder); } catch { /* no existe */ }

                                    if (existing != null)
                                    {
                                        errors.Add("Cuota existente, actualizando…");
                                        existing.ApplyTemplate(template);
                                        existing.Commit();
                                        errors.Add("Cuota actualizada.");
                                    }
                                    else
                                    {
                                        dynamic q = qm.CreateQuota(quotaFolder);
                                        q.ApplyTemplate(template);
                                        q.Commit();
                                        errors.Add("Cuota creada.");
                                    }
                                }
                                finally
                                {
                                    Marshal.ReleaseComObject(qm);
                                }
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"Error configurando cuota FSRM: {ex.Message}");
                            }
                        }
                        else
                        {
                            errors.Add($"La carpeta {folderPath} ya existe, omitiendo creación y cuota.");
                        }
                    }
                    else if (user.OUPrincipal == "OAGER")
                    {
                        errors.Add("OU es 'OAGER', no se configura carpeta ni cuota.");
                    }
                }
                catch (Exception ex)
                {
                    errors.Add($"[FATAL] Excepción inesperada en CreateUser: {ex.Message}");
                }
                finally
                {
                    // **Aquí sólo escribimos el log, sin devolver nada**
                    try
                    {
                        System.IO.File.WriteAllText(@"C:\Temp\AltaUsuario.log",
                            string.Join(Environment.NewLine, errors));
                    }
                    catch
                    {
                        // ignorar fallos al escribir el log
                    }
                }
            }
        }
        catch (Exception ex)
        {
            errors.Add($"Error general en el proceso de creación del usuario: {ex.Message}");
        }


        // Construir el mensaje de respuesta
        string message;
        if (userCreated)
        {
            message = "Usuario creado exitosamente.";
            if (!string.IsNullOrEmpty(grupoDepartamento))
            {
                if (addedToDepartmentGroup)
                {
                    message += $" Añadido al grupo del departamento '{grupoDepartamento}' y a los grupos seleccionados.";
                }
                else
                {
                    message += $" Sin embargo, no se pudo añadir al grupo del departamento '{grupoDepartamento}'.";
                }
            }
            else
            {
                message += " No se especificó un grupo de departamento en la OU (atributo 'street').";
            }

            message += "\nDetalles del proceso:\n- " + (errors.Any() ? string.Join("\n- ", errors) : "No se registraron eventos adicionales.");
        }
        else
        {
            message = "No se pudo crear el usuario.";
            message += "\nErrores encontrados:\n- " + (errors.Any() ? string.Join("\n- ", errors) : "No se proporcionaron detalles adicionales sobre el error.");
        }

        return Json(new { success = userCreated, message });
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

    // Nuevo método para verificar si el DNI ya existe en el Active Directory
    [HttpPost]
    public IActionResult CheckDNIExists([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("dni"))
        {
            return Json(new { success = false, message = "DNI no especificado." });
        }

        string dni = requestData["dni"];
        if (string.IsNullOrEmpty(dni))
        {
            return Json(new { success = false, message = "DNI no puede estar vacío." });
        }

        try
        {
            using (var entry = new DirectoryEntry("LDAP://DC=aytosa,DC=inet"))
            {
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(&(objectClass=user)(employeeID={dni}))";
                    searcher.PropertiesToLoad.Add("employeeID");
                    searcher.SearchScope = SearchScope.Subtree;

                    var result = searcher.FindOne();
                    return Json(new { success = true, exists = result != null });
                }
            }
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al verificar el DNI: {ex.Message}" });
        }
    }



    [HttpPost]
    public IActionResult CheckOtherTelephoneExists([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("numeroLargoFijo"))
        {
            return Json(new { success = false, message = "Número largo fijo no especificado." });
        }

        string numeroLargoFijo = requestData["numeroLargoFijo"];
        if (string.IsNullOrEmpty(numeroLargoFijo))
        {
            return Json(new { success = false, message = "Número largo fijo no puede estar vacío." });
        }

        try
        {
            using (var entry = new DirectoryEntry("LDAP://DC=aytosa,DC=inet"))
            {
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(&(objectClass=user)(otherTelephone={numeroLargoFijo}))";
                    searcher.PropertiesToLoad.Add("otherTelephone");
                    searcher.SearchScope = SearchScope.Subtree;

                    var result = searcher.FindOne();
                    return Json(new { success = true, exists = result != null });
                }
            }
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al verificar el número largo fijo: {ex.Message}" });
        }
    }

    [HttpPost]
    public IActionResult CheckMobileExists([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("extensionMovil"))
        {
            return Json(new { success = false, message = "Extensión del móvil no especificada." });
        }

        string extensionMovil = requestData["extensionMovil"];
        if (string.IsNullOrEmpty(extensionMovil))
        {
            return Json(new { success = false, message = "Extensión del móvil no puede estar vacía." });
        }

        try
        {
            using (var entry = new DirectoryEntry("LDAP://DC=aytosa,DC=inet"))
            {
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(&(objectClass=user)(mobile={extensionMovil}))";
                    searcher.PropertiesToLoad.Add("mobile");
                    searcher.SearchScope = SearchScope.Subtree;

                    var result = searcher.FindOne();
                    return Json(new { success = true, exists = result != null });
                }
            }
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al verificar la extensión del móvil: {ex.Message}" });
        }
    }

    [HttpPost]
    public IActionResult CheckOtherMobileExists([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("numeroLargoMovil"))
        {
            return Json(new { success = false, message = "Número largo del móvil no especificado." });
        }

        string numeroLargoMovil = requestData["numeroLargoMovil"];
        if (string.IsNullOrEmpty(numeroLargoMovil))
        {
            return Json(new { success = false, message = "Número largo del móvil no puede estar vacío." });
        }

        try
        {
            using (var entry = new DirectoryEntry("LDAP://DC=aytosa,DC=inet"))
            {
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(&(objectClass=user)(otherMobile={numeroLargoMovil}))";
                    searcher.PropertiesToLoad.Add("otherMobile");
                    searcher.SearchScope = SearchScope.Subtree;

                    var result = searcher.FindOne();
                    return Json(new { success = true, exists = result != null });
                }
            }
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al verificar el número largo del móvil: {ex.Message}" });
        }
    }

    // Nuevo método para verificar si el DNI ya existe en el Active Directory
    [HttpPost]
    public IActionResult checkTarjetaIdentificativaExists([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData == null || !requestData.ContainsKey("tarjetaIdentificativa"))
        {
            return Json(new { success = false, message = "tarjeta identificativa no especificada." });
        }

        string tarjetaIdentificativa = requestData["tarjetaIdentificativa"];
        if (string.IsNullOrEmpty(tarjetaIdentificativa))
        {
            return Json(new { success = false, message = "Tarjeta Identificativa no puede estar vacío." });
        }

        try
        {
            using (var entry = new DirectoryEntry("LDAP://DC=aytosa,DC=inet"))
            {
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(&(objectClass=user)(employeeID={tarjetaIdentificativa}))";
                    searcher.PropertiesToLoad.Add("serialNumber");
                    searcher.SearchScope = SearchScope.Subtree;

                    var result = searcher.FindOne();
                    return Json(new { success = true, exists = result != null });
                }
            }
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al verificar la tarjeta identificativa: {ex.Message}" });
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