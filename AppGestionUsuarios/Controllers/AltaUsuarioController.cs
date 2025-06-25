using Microsoft.AspNetCore.Mvc;
using System.DirectoryServices;
using Microsoft.AspNetCore.Authorization;
using System.Globalization;
using System.Text;
using System.Management.Automation;
using System.DirectoryServices.AccountManagement;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Runtime.InteropServices;
using Microsoft.Graph.Models;
using Microsoft.Graph;
using Azure.Identity;
using System.Security;
using System.Management.Automation.Runspaces;
using System.ComponentModel;
using Microsoft.Win32.SafeHandles;
using Microsoft.AspNetCore.DataProtection;
using System.Diagnostics;



[Authorize]
public class AltaUsuarioController : Controller
{

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


    //Esto se utiliza para obtener las credenciales de usuario
    private readonly IDataProtector _protector;
    

    //Esto se utiliza para poder leer la información de appsettings
    private readonly IConfiguration _config;

    public AltaUsuarioController(IConfiguration config, IDataProtectionProvider dp)
    {
        _config = config;
        _protector = dp.CreateProtector("CredencialesProtector");
    }

    private GraphServiceClient? _graphClient = null;

    public class userInputModel
    {
        public string Nombre { get; set; }
        public string Apellido1 { get; set; }
        public string Apellido2 { get; set; }
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
            var gruposPorDefecto = _config
                .GetSection("GruposPorDefecto:Grupos")
                .Get<List<string>>()
                ?? new List<string> { "GA_R_PORTALEMPLEADO" };

            ViewBag.GruposPorDefecto = gruposPorDefecto;

            var cuotas = _config
                .GetSection("QuotaSettings:Templates")
                .Get<List<string>>()
                ?? new List<string> { "HOME-500MB", "HOME-1GB", "HOME-2GB" };

            ViewBag.Cuota = cuotas;

            return View("AltaUsuario");
        }
        catch (Exception ex)
        {
            throw new Exception("Error al cargar la página de alta de usuario: " + ex.Message, ex);
        }
    }

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

    private List<string> GetOUPrincipalesFromAD()
    {
        var ouPrincipales = new List<string>();

        try
        {
            using (var rootEntry = new DirectoryEntry(_config["ActiveDirectory:LDAPPath"]))
            {
                using (var searcher = new DirectorySearcher(rootEntry))
                {
                    // Buscar la OU "AREAS" como base
                    searcher.Filter = "(&(objectClass=organizationalUnit)(ou=AREAS))";
                    searcher.SearchScope = SearchScope.Subtree;
                    searcher.PropertiesToLoad.Add("distinguishedName");

                    System.DirectoryServices.SearchResult areasResult = searcher.FindOne();
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
            string ldapPath = $"LDAP://OU=Usuarios,OU={ouPrincipal},{_config["ActiveDirectory:DomainBase"]}";

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
                // Si hay OU secundaria, buscamos el departamento en la OU secundaria dentro de "Usuarios"
                ldapPath = $"LDAP://OU={ouSecundaria},OU=Usuarios,OU={ouPrincipal},{_config["ActiveDirectory:DomainBase"]}";
            }
            else
            {
                // Si no hay OU secundaria, buscamos el departamento en la OU principal
                ldapPath = $"LDAP://OU={ouPrincipal},{_config["ActiveDirectory:DomainBase"]}";
            }

            using (var ouEntry = new DirectoryEntry(ldapPath))
            {
                if (ouEntry == null)
                {
                    return Json(new { success = false, message = "No se pudo conectar a la OU especificada.", ldapPath });
                }

                // Obtener el atributo 'description' de la OU
                string departamento = ouEntry.Properties[_config["GroupInformation:DepartmentAttr"]]?.Value?.ToString();
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
                // Si hay OU secundaria, buscamos el lugar de envío en la OU secundaria dentro de "Usuarios"
                ldapPath = $"LDAP://OU={ouSecundaria},OU=Usuarios,OU={ouPrincipal},{_config["ActiveDirectory:DomainBase"]}";
            }
            else
            {
                // Si no hay OU secundaria, buscamos el lugar de envío en la OU principal
                ldapPath = $"LDAP://OU={ouPrincipal},{_config["ActiveDirectory:DomainBase"]}";
            }

            using (var ouEntry = new DirectoryEntry(ldapPath))
            {
                if (ouEntry == null)
                {
                    return Json(new { success = false, message = "No se pudo conectar a la OU especificada.", ldapPath });
                }

                // Obtener el atributo 'city' de la OU
                string lugarEnvio = ouEntry.Properties[_config["GroupInformation:SendPlaceAttr"]]?.Value?.ToString(); // El atributo 'city' en AD es 'l' (lowercase L)
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
            string.IsNullOrEmpty(user.Username) || string.IsNullOrEmpty(user.Apellido2) ||
            string.IsNullOrEmpty(user.OUPrincipal) || string.IsNullOrEmpty(user.FechaCaducidadOp))
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
                ldapPath = $"LDAP://OU=Usuarios, OU={user.OUSecundaria},OU=Usuarios,OU={user.OUPrincipal},{_config["ActiveDirectory:DomainBase"]}";
                ouPath = $"LDAP://OU={user.OUSecundaria},OU=Usuarios,OU={user.OUPrincipal},{_config["ActiveDirectory:DomainBase"]}";
                errors.Add($"Usando OU secundaria: Path = {ouPath}");
            }
            else
            {
                ldapPath = $"LDAP://OU=Usuarios,OU={user.OUPrincipal},{_config["ActiveDirectory:DomainBase"]}";
                ouPath = $"LDAP://OU={user.OUPrincipal},{_config["ActiveDirectory:DomainBase"]}";
                errors.Add($"Usando OU principal: Path = {ouPath}");
            }

            string departamento = null;
            string area = null;
            string lugarEnvio = null;

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
                        //Obtiene el nombre del departamento según la ou
                        departamento = ouEntryForAttributes.Properties[_config["GroupInformation:DepartmentAttr"]]?.Value?.ToString();
                        if (string.IsNullOrEmpty(departamento))
                        {
                            departamento = "";
                            errors.Add("Atributo 'st' no definido, usando valor predeterminado: 'sin departamento'.");
                        }
                        else
                        {
                            errors.Add($"Atributo 'st' encontrado: '{departamento}'.");
                        }


                        // obtiene el nombre del área
                        area = ouEntryForAttributes.Properties[_config["GroupInformation:AreaAttr"]]?.Value?.ToString();
                        if (string.IsNullOrEmpty(area))
                        {
                            area = "Sin area";
                            errors.Add("Atributo 'description' no definido, usando valor predeterminado: 'Sin Área'.");
                        }
                        else
                        {
                            errors.Add($"Atributo 'description' encontrado: '{area}'.");
                        }

                        // Obtener el grupo de usuarios asociado al departamento concreto
                        grupoDepartamento = ouEntryForAttributes.Properties[_config["GroupInformation:DepartmentGroupAttr"]]?.Value?.ToString();
                        if (string.IsNullOrEmpty(grupoDepartamento))
                        {
                            grupoDepartamento = "";
                            errors.Add($"El atributo 'street' no está definido en la OU (Path: {ouPath}).");
                        }
                        else
                        {
                            errors.Add($"Atributo 'street' encontrado: '{grupoDepartamento}' (Path: {ouPath}).");
                        }

                        // Obtener el lugar de envío
                        lugarEnvio = ouEntryForAttributes.Properties[_config["GroupInformation:SendPlaceAttr"]]?.Value?.ToString();
                        if (string.IsNullOrEmpty(lugarEnvio))
                        {
                            lugarEnvio = "";
                            errors.Add($"El atributo 'l' no está definido en la OU (Path: {ouPath}).");
                        }
                        else
                        {
                            errors.Add($"Atributo 'l' encontrado: '{lugarEnvio}' (Path: {ouPath}).");
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
                    errors.Add($"Usuario creado en la OU: CN={ldapPath}");

                    // Establecer atributos básicos del usuario
                    try
                    {
                        newUser.Properties[_config["ADAttributes:NameAttr"]].Value = user.Nombre;
                        newUser.Properties[_config["ADAttributes:SurnameAttr"]].Value = user.Apellido1 + " " + user.Apellido2;
                        newUser.Properties[_config["ADAttributes:UsernameAttr"]].Value = user.Username;
                        newUser.Properties[_config["ADAttributes:NameAndDomainAttr"]].Value = $"{user.Username}@aytosa.inet";
                        newUser.Properties[_config["ADAttributes:DisplayNameAttr"]].Value = displayName;
                        newUser.Properties[_config["ADAttributes:DepartmentAttr"]].Value = departamento;
                        newUser.Properties[_config["ADAttributes:AreaAttr"]].Value = area;
                        errors.Add("Atributos básicos del usuario establecidos correctamente.");
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Error al establecer los atributos básicos del usuario: {ex.Message}");
                    }

                    // Asignar otros campos
                    try
                    {
                        if (!string.IsNullOrEmpty(user.NFuncionario))
                        {
                            newUser.Properties[_config["ADAttributes:NFuncionarioAttr"]].Value = user.NFuncionario;
                            errors.Add($"Atributo 'NFuncionarioAttr' (NFuncionario) establecido: {user.NFuncionario}");
                        }
                        else
                        {
                            errors.Add($"Atributo 'NFuncionarioAttr' se recibió como null o vacio");

                        }
                        //NTelefono se refiere a la extensión del número fijo
                        if (!string.IsNullOrEmpty(user.NTelefono))
                        {
                            newUser.Properties[_config["ADAttributes:TelephoneNumberAttr"]].Value = user.NTelefono;
                            errors.Add($"Atributo 'TelephoneNumberAttr' establecido: {user.NTelefono}");
                        }
                        else
                        {
                            errors.Add($"Atributo 'TelephoneNumberAttr' se recibió como null o vacio");

                        }
                        //Se refiere al número fijo completo
                        if (!string.IsNullOrEmpty(user.NumeroLargoFijo))
                        {
                            newUser.Properties[_config["ADAttributes:OtherTelephoneAttr"]].Value = user.NumeroLargoFijo;
                            errors.Add($"Atributo 'OtherTelephoneAttr' establecido: {user.NumeroLargoFijo}");
                        }
                        else
                        {
                            errors.Add($"Atributo 'OtherTelephoneAttr' se recibió como null o vacio");

                        }
                        if (!string.IsNullOrEmpty(user.ExtensionMovil))
                        {
                            newUser.Properties[_config["ADAttributes:MobileExtensionAttr"]].Value = user.ExtensionMovil;
                            errors.Add($"Atributo 'MobileExtensionAttr' establecido: {user.ExtensionMovil}");
                        }
                        else
                        {
                            errors.Add($"Atributo 'MobileextensionAttr' se recibió como null o vacio");

                        }
                        if (!string.IsNullOrEmpty(user.NumeroLargoMovil))
                        {
                            newUser.Properties[_config["ADAttributes:LargeMobileNumberAttr"]].Value = user.NumeroLargoMovil;
                            errors.Add($"Atributo 'LargeMobileNumberAttr' establecido: {user.NumeroLargoMovil}");
                        }
                        else
                        {
                            errors.Add($"Atributo 'MobileextensionAttr' se recibió como null o vacio");

                        }
                        if (!string.IsNullOrEmpty(user.TarjetaIdentificativa))
                        {
                            newUser.Properties[_config["ADAttributes:IDCardAttr"]].Value = user.TarjetaIdentificativa;
                            errors.Add($"Atributo 'IDCardAttr' establecido: {user.TarjetaIdentificativa}");
                        }
                        else
                        {
                            errors.Add($"Atributo 'IDCardAttr' se recibió como null o vacio");

                        }
                        if (!string.IsNullOrEmpty(user.DNI))
                        {
                            newUser.Properties[_config["ADAttributes:DNIAttr"]].Value = user.DNI;
                            errors.Add($"Atributo 'DNIAttr' establecido: {user.DNI}");
                        }
                        else
                        {
                            errors.Add($"Atributo 'DNIAttr' se recibió como null o vacio");

                        }
                        newUser.Properties[_config["ADAttributes:LocationAttr"]].Value = user.LugarEnvio;
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
                                newUser.Properties[_config["ADAttributes:AccountExpiresAttr"]].Value = accountExpires.ToString();
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
                        newUser.Invoke("SetPassword", new object[] { _config["ActiveDirectory:TemporalPassword"] });
                        newUser.Properties[_config["ADAttributes:EnableAccountAttr"]].Value = 0x200;
                        newUser.Properties[_config["ADAttributes:ChangePassNextLoginAttr"]].Value = 0;
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


                                string adminUsername = _config["ActiveDirectory:AppAdministrator"];
                                string adminFileSystem = _config["ActiveDirectory:FileSystemAdministrator"];
                                string adminDomainSystem = _config["ActiveDirectory:DomainSystemAdministrator"];
                                string quotaDomain = _config["ActiveDirectory:QuotaDomain"];

                                // 2) NTFS: permisos sobre \\LEONARDO\Home\<user>
                                DirectoryInfo di = new DirectoryInfo(folderPath);
                                var ds = new DirectorySecurity();
                                // FullControl a las cuentas de administración
                                ds.AddAccessRule(new FileSystemAccessRule(
                                    new NTAccount($"{quotaDomain}\\{adminFileSystem}"),
                                    FileSystemRights.FullControl,
                                    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                    PropagationFlags.None,
                                    AccessControlType.Allow));
                                ds.AddAccessRule(new FileSystemAccessRule(
                                    new NTAccount($"{quotaDomain}\\{adminDomainSystem}"),
                                    FileSystemRights.FullControl,
                                    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                    PropagationFlags.None,
                                    AccessControlType.Allow));

                                // Permisos del propio usuario
                                ds.AddAccessRule(new FileSystemAccessRule(
                                    new NTAccount($"{quotaDomain}\\{user.Username}"),
                                    FileSystemRights.ReadAndExecute | FileSystemRights.Write | FileSystemRights.DeleteSubdirectoriesAndFiles,
                                    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                    PropagationFlags.None,
                                    AccessControlType.Allow));


                                ds.AddAccessRule(new FileSystemAccessRule(
                                    new NTAccount($"{quotaDomain} \\{adminUsername}"),
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
                                string quota = user.Cuota ?? "HOME-1GB";
                                string template = quota;
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

    
    //Método para crear el alta complta de usuario con correo electrónico
    //[HttpPost]
    //public async Task<IActionResult> AltaCompleta([FromBody] UserModelAltaUsuario user)
    //{
    //    var log = new List<string>();
    //    bool ok = true;

    //    if (user == null || string.IsNullOrEmpty(user.Username))
    //        return Json(new { success = false, message = "No se recibieron datos válidos." });

    //    if (string.IsNullOrWhiteSpace(user.adminUser) || string.IsNullOrWhiteSpace(user.adminPassword))
    //        return Json(new { success = false, message = "Faltan credenciales de administrador Exchange." });

    //    log.Add("Validación de credenciales del administrador...");
    //    if (!ValidateCredentials(user.adminUser, user.adminPassword))
    //        return Json(new { success = false, message = "Credenciales de administrador exchange incorrectas" });
    //    log.Add("[OK] Credenciales validadas correctamente.");

    //    log.Add("Inicializando GraphServiceClient...");
    //    var tenantId = _config["AzureAd:TenantId"] ?? throw new InvalidOperationException("Falta AzureAd:TenantId");
    //    var clientId = _config["AzureAd:ClientId"] ?? throw new InvalidOperationException("Falta AzureAd:ClientId");
    //    var clientSecret = _config["AzureAd:ClientSecret"] ?? throw new InvalidOperationException("Falta AzureAd:ClientSecret");
    //    var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    //    _graphClient = new GraphServiceClient(credential);
    //    log.Add("[OK] GraphServiceClient inicializado.");

    //    var samAccountName = user.Username;
    //    var username = samAccountName;

    //    try
    //    {
    //        log.Add("=== Alta Completa iniciado: " + DateTime.Now + " ===");

    //        // Paso 1: Crear usuario en AD
    //        log.Add("Paso 1: Intentando crear el usuario en AD...");
    //        var createResult = CreateUser(user) as JsonResult;
    //        dynamic createData = createResult?.Value;
    //        if (createData == null || !(bool)createData.success)
    //        {
    //            log.Add("[ERROR] Fallo al crear el usuario en AD: " + (createData?.message ?? "sin mensaje"));
    //            return Json(new
    //            {
    //                success = false,
    //                message = "Alta completa abortada: fallo al crear el usuario en AD.",
    //                log
    //            });
    //        }
    //        log.Add("[OK] Usuario creado en AD correctamente.");


    //        // Paso 2: Añadir a grupo de licencias
    //        log.Add("Paso 2: Añadiendo al grupo de licencias...");
    //        var grupoResult = ModifyUserGroup(samAccountName) as JsonResult;
    //        dynamic grupoData = grupoResult?.Value;
    //        if (grupoData == null || !(bool)grupoData.success)
    //        {
    //            log.Add("[ERROR] No se pudo añadir al grupo: " + (grupoData?.message ?? "sin mensaje"));
    //            ok = false;
    //        }
    //        else
    //        {
    //            log.Add("[OK] Usuario añadido al grupo de licencias.");
    //        }

    //        // Paso 3: Sincronizar con Azure AD
    //        log.Add("Paso 3: Lanzando sincronización Delta con Azure AD Connect...");
    //        var (syncOk, syncErr) = SyncDeltaOnVanGogh();
    //        if (!syncOk)
    //        {
    //            log.Add("[ERROR] Error en la sincronización: " + syncErr);
    //            ok = false;
    //            throw new InvalidOperationException(syncErr);
    //        }
    //        log.Add("[OK] Sincronización Delta completada.");

    //        // Paso 4: Esperar a que aparezca en Azure AD
    //        log.Add("Paso 4: Esperando aparición del usuario en Azure AD...");
    //        var exists = await WaitForAzureUser(samAccountName);
    //        if (exists)
    //            log.Add("[OK] Usuario encontrado en Azure AD.");
    //        else
    //            log.Add("[WARN] Timeout esperando al usuario en Azure AD.");

    //        // Paso 5: Crear buzón on-prem
    //        log.Add("Paso 5: Habilitando buzón on-prem...");
    //        EnableOnPremMailbox(username, user.adminUser, user.adminPassword);
    //        log.Add("[OK] Buzón on-prem habilitado correctamente.");

    //        // Paso 6: Actualizar proxyAddresses
    //        log.Add("Paso 6: Actualizando proxyAddresses...");
    //        UpdateProxyAddresses(samAccountName);
    //        log.Add("[OK] proxyAddresses actualizadas.");

    //        // Paso 7: Crear lote de migración
    //        log.Add("Paso 7: Creando y lanzando lote de migración...");
    //        CreateMigrationBatch(new[] { username });
    //        log.Add("[OK] Lote de migración lanzado.");

    //        log.Add("=== Alta Completa finalizada con éxito ===");
    //    }
    //    catch (Exception ex)
    //    {
    //        ok = false;
    //        log.Add("[ERROR] " + ex.Message);
    //        log.Add("=== Alta Completa abortada ===");
    //    }

    //    // Guardar log en disco
    //    System.IO.File.WriteAllLines(
    //        Path.Combine(Path.GetTempPath(), $"AltaCompleta_{samAccountName}.log"),
    //        log);

    //    var message = ok
    //        ? "Alta completa realizada con éxito.\n" + string.Join("\n", log)
    //        : "Se produjeron errores en la Alta Completa:\n" + string.Join("\n", log);

    //    return Json(new { success = ok, message });
    //}



    //Función encargada de comvertir el username recibido de una vista en string y pasarlo a la función que lo busca en AD


    [HttpPost]
    public async Task<IActionResult> AltaCompleta([FromBody] UserModelAltaUsuario user)
    {
        

        string domain = _config["ActiveDirectory:DomainName"];
        string adminUsername = HttpContext.Session.GetString("adminUser");
        var encryptedPass = HttpContext.Session.GetString("adminPassword");
        var adminPassword = _protector.Unprotect(encryptedPass);

        // 3) LogonUser → token Windows
        if (!LogonUser(
                adminUsername,
                domain,
                adminPassword,
                LOGON32_LOGON_NEW_CREDENTIALS,
                LOGON32_PROVIDER_DEFAULT,
                out var userToken))
        {
            var err = new Win32Exception(Marshal.GetLastWin32Error()).Message;
            return Json(new { success = false, message = $"No se pudo loguear: {err}" });
        }

        try
        {
            // 4) Envolver TODO en impersonación
            using var safeToken = new SafeAccessTokenHandle(userToken);
            IActionResult finalResult = null;

            await WindowsIdentity.RunImpersonated(safeToken, async () =>
            {
                var log = new List<string>();
                bool ok = true;


                log.Add("Inicializando GraphServiceClient...");
                var tenantId = _config["AzureAd:TenantId"] ?? throw new InvalidOperationException("Falta AzureAd:TenantId");
                var clientId = _config["AzureAd:ClientId"] ?? throw new InvalidOperationException("Falta AzureAd:ClientId");
                var clientSecret = _config["AzureAd:ClientSecret"] ?? throw new InvalidOperationException("Falta AzureAd:ClientSecret");
                var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
                _graphClient = new GraphServiceClient(credential);
                log.Add("[OK] GraphServiceClient inicializado.");

                var samAccountName = user.Username;

                try
                {
                    log.Add("=== Alta Completa iniciado: " + DateTime.Now + " ===");

                    // Paso 1: Crear usuario en AD
                    log.Add("Paso 1: Intentando crear el usuario en AD...");
                    var createResult = CreateUser(user) as JsonResult;
                    dynamic createData = createResult?.Value;
                    if (createData == null || !(bool)createData.success)
                    {
                        log.Add("[ERROR] Fallo al crear el usuario en AD: " + (createData?.message ?? "sin mensaje"));
                        finalResult = Json(new
                        {
                            success = false,
                            message = "Alta completa abortada: fallo al crear el usuario en AD." ,
                            log
                        });
                        return;
                    }
                    log.Add("[OK] Usuario creado en AD correctamente.");

                    // Paso 2: Añadir a grupo de licencias
                    log.Add("Paso 2: Añadiendo al grupo de licencias...");
                    var grupoResult = ModifyUserGroup(samAccountName) as JsonResult;
                    dynamic grupoData = grupoResult?.Value;
                    if (grupoData == null || !(bool)grupoData.success)
                    {
                        log.Add("[ERROR] No se pudo añadir al grupo: " + (grupoData?.message ?? "sin mensaje"));
                        ok = false;
                    }
                    else
                    {
                        log.Add("[OK] Usuario añadido al grupo de licencias.");
                    }

                    // Paso 3: Sincronizar con Azure AD Connect
                    log.Add("Paso 3: Lanzando sincronización Delta con Azure AD Connect...");
                    var (syncOk, syncErr) = SyncDeltaOnVanGogh();
                    if (!syncOk)
                    {
                        log.Add("[ERROR] Error en la sincronización: " + syncErr);
                        ok = false;
                        throw new InvalidOperationException(syncErr);
                    }
                    log.Add("[OK] Sincronización Delta completada.");

                    // Paso 4: Esperar a que aparezca en Azure AD
                    log.Add("Paso 4: Esperando aparición del usuario en Azure AD...");
                    var exists = await WaitForAzureUser(samAccountName);
                    if (exists)
                        log.Add("[OK] Usuario encontrado en Azure AD.");
                    else
                        log.Add("[WARN] Timeout esperando al usuario en Azure AD.");

                    // Paso 5: Crear buzón on-prem
                    log.Add("Paso 5: Habilitando buzón on-prem...");
                    EnableOnPremMailbox(samAccountName, adminUsername, adminPassword);
                    log.Add("[OK] Buzón on-prem habilitado correctamente.");

                    // Paso 6: Actualizar proxyAddresses
                    log.Add("Paso 6: Actualizando proxyAddresses...");
                    UpdateProxyAddresses(samAccountName);
                    log.Add("[OK] proxyAddresses actualizadas.");

                    // Paso 7: Crear lote de migración
                    log.Add("Paso 7: Creando y lanzando lote de migración...");
                    CreateMigrationBatch(new[] { samAccountName });
                    log.Add("[OK] Lote de migración lanzado.");

                    log.Add("=== Alta Completa finalizada con éxito ===");
                }
                catch (Exception ex)
                {
                    ok = false;
                    log.Add("[ERROR] " + ex.Message);
                    log.Add("=== Alta Completa abortada ===");
                }

                
                // Preparar resultado final
                var message = ok
                    ? "Alta completa realizada con éxito.\n" + string.Join("\n", log)
                    : "Se produjeron errores en la Alta Completa:\n" + string.Join("\n", log);

                finalResult = Json(new { success = ok, message });
            });

            return finalResult;
        }
        finally
        {
            CloseHandle(userToken);
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
                string domain = _config["ActiveDirectory:DomainName"];
                string attributeName = "description";

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
                        System.DirectoryServices.SearchResult result = searcher.FindOne();

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
                string domain = _config["ActiveDirectory:DomainName"];
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
                        System.DirectoryServices.SearchResult result = searcher.FindOne();

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
    public DirectoryEntry FindGroupByName(string groupName)
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

                    System.DirectoryServices.SearchResult result = searcher.FindOne();
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

    public (bool Success, string OutputOrError) SyncDeltaOnVanGogh()
    {
        // 0) Recuperar usuario y clave de sesión
        string adminUsername = HttpContext.Session.GetString("adminUser");
        var encryptedPass = HttpContext.Session.GetString("adminPassword");
        var adminPassword = _protector.Unprotect(encryptedPass);

        // 1) Servidor ADSync desde configuración
        var server = _config["AzureAdSync:SyncServer"]
                     ?? throw new InvalidOperationException("Falta AzureAdSync:SyncServer en config");

        // 2) Construimos el script PowerShell completo
        //    Creamos un SecureString y un PSCredential, luego Invoke-Command con -Credential.
        var remoteScript = $@"
      $sec = ConvertTo-SecureString '{adminPassword}' -AsPlainText -Force
      $cred = New-Object System.Management.Automation.PSCredential('{adminUsername}', $sec)
      Invoke-Command -ComputerName '{server}' -Credential $cred -ScriptBlock {{
          Import-Module ADSync -ErrorAction Stop
          Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
      }} -ErrorAction Stop
    ";

        // 3) Preparamos el proceso externo de PowerShell
        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command \"{remoteScript}\"",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        try
        {
            using var proc = System.Diagnostics.Process.Start(psi);
            if (proc == null)
                return (false, "No se pudo iniciar powershell.exe");

            // 4) Capturamos stdout y stderr
            string stdout = proc.StandardOutput.ReadToEnd();
            string stderr = proc.StandardError.ReadToEnd();
            proc.WaitForExit();

            // 5) Evaluamos el código de salida
            if (proc.ExitCode != 0)
            {
                var errorText = !string.IsNullOrWhiteSpace(stderr) ? stderr : stdout;
                return (false, errorText.Trim());
            }

            // 6) Éxito: devolvemos la salida
            return (true, stdout.Trim());
        }
        catch (Exception ex)
        {
            // 7) Cualquier fallo al invocar powershell.exe
            return (false, ex.Message);
        }
    }


    /// <summary>
    /// Espera 30 s, luego intenta hasta 4 veces comprobar en Graph si el usuario existe,
    /// con 10 s de espera entre cada intento. Cualquier excepción cuenta como “no existe”.
    /// </summary>
    public async Task<bool> WaitForAzureUser(string upn)
    {
        const int maxAttempts = 4;
        const int initialDelaySeconds = 30;
        const int retryDelaySeconds = 10;

        // 1) Delay inicial
        await Task.Delay(TimeSpan.FromSeconds(initialDelaySeconds));

        for (int attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                // Si no lanza excepción, el usuario existe
                await _graphClient.Users[upn].GetAsync();
                return true;
            }
            catch
            {
                // Si es el último intento, devolvemos false
                if (attempt == maxAttempts)
                    return false;

                // Si no, esperamos y reintentamos
                await Task.Delay(TimeSpan.FromSeconds(retryDelaySeconds));
            }
        }

        return false;
    }


    //Código antiguo de enable onpremMailBox, funciona solo en el entorno de pruebas
    //public void EnableOnPremMailbox(string username, string adminRunAs, string adminPassword)
    //{
    //    // 1) Lee sólo los parámetros que sigan en config
    //    var server = _config["Exchange:Server"]
    //                 ?? throw new InvalidOperationException("Falta Exchange:Server");
    //    var dbName = _config["Exchange:Database"]
    //                 ?? throw new InvalidOperationException("Falta Exchange:Database");
    //    var domain = _config["ActiveDirectory:DomainName"]
    //                 ?? throw new InvalidOperationException("Falta ActiveDirectory:DomainName");

    //    // 2) Construye el SecureString a partir de la password recibida
    //    var securePwd = new SecureString();
    //    foreach (var c in adminPassword)
    //        securePwd.AppendChar(c);

    //    var cred = new PSCredential(adminRunAs, securePwd);

    //    // 3) Crea el script bloque
    //    var identity = $"{domain}\\{username}";
    //    var script = $@"
    //    Import-Module ExchangeOnlineManagement
    //    Enable-Mailbox -Identity '{identity}' -Alias '{username}' -Database '{dbName}'
    //";

    //    // 4) Ejecuta Invoke-Command con esas credenciales
    //    using var ps = PowerShell.Create();
    //    ps.AddCommand("Invoke-Command")
    //      .AddParameter("ComputerName", server)
    //      .AddParameter("Credential", cred)
    //      .AddParameter("ScriptBlock", ScriptBlock.Create(script));

    //    var results = ps.Invoke();
    //    if (ps.Streams.Error.Count > 0)
    //    {
    //        var errs = string.Join(";\n", ps.Streams.Error.ReadAll().Select(e => e.ToString()));
    //        throw new InvalidOperationException($"Error al habilitar buzón on-prem: {errs}");
    //    }
    //}

    //Nuevo enable OnPreMailBox se conecta a la directamente a la shell de exchange antes de ejecutar el comando remoto
    public void EnableOnPremMailbox(string username, string adminRunAs, string adminPassword)
    {
        // 1) Parámetros de configuración
        var server = _config["Exchange:Server"]
                     ?? throw new InvalidOperationException("Falta Exchange:Server");
        var dbName = _config["Exchange:Database"]
                     ?? throw new InvalidOperationException("Falta Exchange:Database");
        var domain = _config["ActiveDirectory:DomainName"]
                     ?? throw new InvalidOperationException("Falta ActiveDirectory:DomainName");

        // 2) Impersonación para obtener ticket Kerberos
        if (!LogonUser(
                adminRunAs,
                domain,
                adminPassword,
                LOGON32_LOGON_NEW_CREDENTIALS,
                LOGON32_PROVIDER_DEFAULT,
                out var userToken))
        {
            var err = new Win32Exception(Marshal.GetLastWin32Error()).Message;
            throw new InvalidOperationException($"Imposible impersonar: {err}");
        }

        using var safeToken = new SafeAccessTokenHandle(userToken);

        // 3) Ejecutar todo bajo la identidad impersonada
        WindowsIdentity.RunImpersonated(safeToken, () =>
        {
            // 3.1) Configurar conexión WinRM a la EMS de Exchange
            var uri = new Uri($"http://{server}/PowerShell/");
            var connectionInfo = new WSManConnectionInfo(
                uri,
                "http://schemas.microsoft.com/powershell/Microsoft.Exchange",
                credential: null  // usa Kerberos delegado
            )
            {
                AuthenticationMechanism = AuthenticationMechanism.Negotiate,
                OperationTimeout = 4 * 60 * 1000,  // 4 minutos
                OpenTimeout = 1 * 60 * 1000   // 1 minuto
            };

            // 3.2) Abrir runspace remoto
            using var runspace = RunspaceFactory.CreateRunspace(connectionInfo);
            runspace.Open();

            // 3.3) Invocar Enable-Mailbox directamente en ese runspace
            using var ps = PowerShell.Create();
            ps.Runspace = runspace;
            ps.AddCommand("Enable-Mailbox")
              .AddParameter("Identity", $"{domain}\\{username}")
              .AddParameter("Alias", username)
              .AddParameter("Database", dbName);

            ps.Invoke();
            if (ps.Streams.Error.Count > 0)
            {
                var errs = string.Join(";\n", ps.Streams.Error
                                               .ReadAll()
                                               .Select(e => e.ToString()));
                throw new InvalidOperationException($"Error al habilitar buzón on-prem via Kerberos: {errs}");
            }
        });
    }

    private static void ThrowIfErrors(PowerShell ps, string paso)
    {
        if (!ps.HadErrors) return;
        var msg = string.Join(" | ", ps.Streams.Error.Select(e => e.ToString()));
        throw new InvalidOperationException($"Error al {paso}: {msg}");
    }


    public void UpdateProxyAddresses(string samAccountName)
    {
        // 1) Crear contexto y buscar usuario
        using var ctx = new PrincipalContext(ContextType.Domain, _config["ActiveDirectory:DomainName"]);
        using var user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, samAccountName);
        if (user == null)
            throw new InvalidOperationException($"Usuario '{samAccountName}' no encontrado en AD.");

        // 2) Obtener su DirectoryEntry
        var de = user.GetUnderlyingObject() as DirectoryEntry
                 ?? throw new InvalidOperationException("No se pudo leer DirectoryEntry del usuario.");

        // 3) Leer dominios
        var oldDomain = _config["ActiveDirectory:OldDomain"]
                        ?? throw new InvalidOperationException("Falta ActiveDirectory:OldDomain");
        var newDomain = _config["ActiveDirectory:NewDomainNAme"]
                        ?? throw new InvalidOperationException("Falta ActiveDirectory:NewDomainNAme");

        // 4) Acceder a proxyAddresses
        var proxies = de.Properties["proxyAddresses"];
                
        // 6) Añadir la nueva primaria
        var newProxy = $"SMTP:{samAccountName}@{newDomain}";
        if (!proxies.Cast<string>().Any(p => p.Equals(newProxy, StringComparison.OrdinalIgnoreCase)))
        {
            proxies.Add(newProxy);
        }

        // 7) Guardar cambios
        de.CommitChanges();
    }


    public void AddUserToGroup(string samAccountName)
    {
        // 1) Crear contexto y buscar usuario
        using var ctx = new PrincipalContext(ContextType.Domain, _config["ActiveDirectory:DomainName"]);
        using var user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, samAccountName);
        if (user == null)
            throw new InvalidOperationException($"Usuario '{samAccountName}' no encontrado en AD.");

        Console.WriteLine(" ++++++++++++++++ se ha encontrado " + samAccountName);

        // 2) Obtener su DirectoryEntry
        var de = user.GetUnderlyingObject() as DirectoryEntry
                 ?? throw new InvalidOperationException("No se pudo leer DirectoryEntry del usuario.");
        var userDn = de.Properties["distinguishedName"].Value.ToString();
        Console.WriteLine("No sé que es el Directory Entry, pero está");

        // 3) Abrir el entry del grupo
        var groupDn = _config["ActiveDirectory:LicenseGroupDn"]
                      ?? throw new InvalidOperationException("Falta ActiveDirectory:LicenseGroupDn");
        Console.WriteLine("+++++++++++Se Buscagrupos" + groupDn);

        using var grp = new DirectoryEntry("LDAP://" + groupDn);
        var members = grp.Properties["member"];

        Console.WriteLine("+++++++++++Se han encontrado grupos" + grp.Username);

        // 4) Añadirlo al grupo de licencias si aún no es miembro
        if (!members.Cast<string>().Any(m => m.Equals(userDn, StringComparison.OrdinalIgnoreCase)))
        {
            try
            {
                members.Add(userDn);
                grp.CommitChanges();
            }
            catch (Exception ex)
            {
                Console.WriteLine("!!!!!!!!!!!!!!!!!!!!!!" + ex.ToString());
                throw ex;
            }
            Console.WriteLine("+++++++++++Se ha añadido el grupo  y se ha metido en el grupo");
        }

        Console.WriteLine("----------- Se ha incluido el usuario al grupo");

    }

    /// <summary>
    /// Crea un lote de migración híbrida en Exchange Online utilizando el módulo ExchangeOnlineManagement.
    /// Si el módulo no está instalado, se instala en el perfil del usuario.
    /// </summary>
    /// //Método de migración antiguo sin usar la exchange shell


    //Nueva creación de lote de migración
    public void CreateMigrationBatch(string[] upns)
    {
        // Config
        var appId = _config["AzureAd:ClientId"]
                     ?? throw new InvalidOperationException("Falta AzureAd:ClientId");
        var tenantId = _config["AzureAd:TenantId"]
                     ?? throw new InvalidOperationException("Falta AzureAd:TenantId");
        var secret = _config["AzureAd:ClientSecret"]
                     ?? throw new InvalidOperationException("Falta AzureAd:ClientSecret");

        var endpoint = _config["Exchange:Endpoint"] ?? "saura";
        var tgtDomain = _config["Exchange:TargetDeliveryDomain"] ?? "aytosalamanca.mail.onmicrosoft.com";
        var batchName = $"Migra_{DateTime.UtcNow:yyyyMMdd_HHmmss}";

        // 1) Runspace con módulo EMS-EXO precargado
        var iss = InitialSessionState.CreateDefault();
        iss.ImportPSModule(new[] { "ExchangeOnlineManagement" });
        using var run = RunspaceFactory.CreateRunspace(iss);
        run.Open();

        using var ps = PowerShell.Create();
        ps.Runspace = run;

        // 2) Conectar a EXO con App-only + Client Secret
        ps.AddCommand("Connect-ExchangeOnline")
          .AddParameter("AppId", appId)
          .AddParameter("TenantId", tenantId)   // ← NO “Organization”
          .AddParameter("ClientSecret", secret)
          .AddParameter("ShowBanner", false);
        ps.Invoke();
        ThrowIfErrors(ps, "conectar a Exchange Online");
        ps.Commands.Clear();

        // 3) Crear el Migration Batch
        ps.AddCommand("New-MigrationBatch")
          .AddParameter("Name", batchName)
          .AddParameter("MigrationType", "RemoteMove")
          .AddParameter("SourceEndpoint", endpoint)
          .AddParameter("TargetDeliveryDomain", tgtDomain)
          .AddParameter("Users", upns)          // array directo
          .AddParameter("AutoStart", true)
          .AddParameter("AutoComplete", true);
        ps.Invoke();
        ThrowIfErrors(ps, "crear el lote de migración");
        ps.Commands.Clear();

        // 4) Desconectar sesión
        ps.AddCommand("Disconnect-ExchangeOnline")
          .AddParameter("Confirm", false)
          .Invoke();
    }

    public bool ValidateCredentials(string usuario, string password)
    {
        string dominio = _config["ActiveDirectory:DomainName"]
                        ?? throw new InvalidOperationException("Falta ActiveDirectory:DomainName");
        using (var ctx = new PrincipalContext(
                   ContextType.Domain,
                   dominio))
        {

            // El método devuelve true solo si el usuario y la contraseña son correctos.
            return ctx.ValidateCredentials(usuario, password);
        }
    }

    [HttpPost]
    public IActionResult ModifyUserGroup(string SAMAccountName)
    {
        if (string.IsNullOrEmpty(SAMAccountName))
        {

        }
        string username = SAMAccountName; // Extrae el nombre de usuario
        string groupName = _config["ActiveDirectory:LicenseGroupDn"];
        string action = "add";

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

                    System.DirectoryServices.SearchResult result = searcher.FindOne();

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


}