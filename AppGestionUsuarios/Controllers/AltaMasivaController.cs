using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Text;
using OfficeOpenXml;
using System.Globalization;
using System.Text.Json;
using Azure.Identity;
using Microsoft.Graph;
using static AltaUsuarioController;
using Microsoft.AspNetCore.DataProtection;
using System.Runtime.Intrinsics.Arm;
using System.Runtime.InteropServices;
using System.ComponentModel;
using Microsoft.Win32.SafeHandles;
using System.Security.Principal;
using Microsoft.AspNetCore.Http.HttpResults;

[Authorize]
public class AltaMasivaController : Controller
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


    private readonly AltaUsuarioController _altaUsuarioController;

    //Esto se utiliza para obtener las credenciales de usuario
    private readonly IDataProtector _protector;

    private GraphServiceClient? _graphClient = null;
    private readonly IConfiguration _config;

    public AltaMasivaController(AltaUsuarioController altaUsuarioController, IConfiguration config, IDataProtectionProvider dp)
    {
        _protector = dp.CreateProtector("CredencialesProtector");
        _altaUsuarioController = altaUsuarioController
            ?? throw new ArgumentNullException(nameof(altaUsuarioController));
        _config = config;
    }



    // Clase para el resultado de las verificaciones
    private class CheckResult
    {
        public bool Success { get; set; }
        public bool Exists { get; set; }
        public string Message { get; set; }
    }

    
    public class ProcessUsersRequest
    {
        public List<Dictionary<string, object>> UsersRaw { get; set; }
    }



    [HttpGet]
    public IActionResult AltaMasiva()
    {
        try
        {
            // traemos todos los grupos de AD para poblar los <select> en la vista
            ViewBag.GruposAD = _altaUsuarioController
                       .GetGruposFromAD()
                       .OrderBy(g => g)
                       .ToList();
            return View();
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al cargar la página de alta masiva: {ex.Message}" });
        }
    }

    // POST: /AltaMasiva/LoadFile
    // Sólo lee el Excel y devuelve usersData
    [HttpPost]
    public async Task<JsonResult> LoadFile(IFormFile file)
    {
        if (file == null || file.Length == 0)
            return Json(new { success = false, message = "No se ha proporcionado un archivo Excel válido." });

        var users = new List<Dictionary<string, string>>();
        try
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial; ;
            using var ms = new MemoryStream();
            await file.CopyToAsync(ms);
            using var pkg = new ExcelPackage(ms);
            var ws = pkg.Workbook.Worksheets.First();
            int rows = ws.Dimension.End.Row;
            int cols = ws.Dimension.End.Column;

            var headers = Enumerable.Range(1, cols)
                                    .Select(c => ws.Cells[1, c].Text.Trim())
                                    .ToList();

            for (int r = 2; r <= rows; r++)
            {
                var dict = new Dictionary<string, string>();
                for (int c = 1; c <= cols; c++)
                    dict[headers[c - 1]] = ws.Cells[r, c].Text.Trim();
                users.Add(dict);
            }
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al procesar el Excel: {ex.Message}" });
        }

        return Json(new { success = true, users });
    }


    public async Task<JsonResult> ProcessUsers([FromBody] ProcessUsersRequest request)
    {

        // 1) Resultado por defecto, por si algo va muy mal
        JsonResult finalResult = Json(new
        {
            success = false,
            message = "No se ejecutó la operación de alta masiva."
        });

        var summaryMessages = new List<string>();
        var createdUsernames = new List<string>();
        bool overallSuccess = true;

        string adminUsername = HttpContext.Session.GetString("adminUser");
        var encryptedPass = HttpContext.Session.GetString("adminPassword");
        var adminPassword = _protector.Unprotect(encryptedPass);
        string domain = _config["ActiveDirectory:DomainName"]
                                  ?? throw new InvalidOperationException("Falta ActiveDirectory:DomainName");

        // 1) Impersonación
        if (!LogonUser(
                adminUsername,
                domain,
                adminPassword,
                LOGON32_LOGON_NEW_CREDENTIALS,
                LOGON32_PROVIDER_DEFAULT,
                out var userToken))
        {
            var err = new Win32Exception(Marshal.GetLastWin32Error()).Message;
            return Json(new { success = false, message = $"Imposible impersonar: {err}" });
        }


        try
        {
            using var safeToken = new SafeAccessTokenHandle(userToken);

            // TODO: dentro de esta lambda se ejecuta TODO bajo impersonación
            await WindowsIdentity.RunImpersonated(safeToken, async () =>
            {

                try 
                { 
                    //// --- inicializar GraphServiceClient ---
                    //var tenantId = _config["AzureAd:TenantId"]
                    //                   ?? throw new InvalidOperationException("Falta AzureAd:TenantId");
                    //var clientId = _config["AzureAd:ClientId"]
                    //                   ?? throw new InvalidOperationException("Falta AzureAd:ClientId");
                    //var clientSecret = _config["AzureAd:ClientSecret"]
                    //                   ?? throw new InvalidOperationException("Falta AzureAd:ClientSecret");
                    //_graphClient = new GraphServiceClient(
                    //                       new ClientSecretCredential(tenantId, clientId, clientSecret)
                    //                   );

                    // --- validación inicial ---
                    if (request.UsersRaw == null || request.UsersRaw.Count == 0)
                    {
                        finalResult = Json(new
                        {
                            success = false,
                            message = "No se han enviado usuarios para procesar.",
                            messages = summaryMessages
                        });
                        return; // ← SALIMOS de la lambda, no devolvemos Json directamente
                    }
                    summaryMessages.Add($"▶ ProcessUsers: recibidas {request.UsersRaw.Count} filas.");

                    // --- tu bucle de filas (idéntico al original) ---
                    for (int i = 0; i < request.UsersRaw.Count; i++)
                    {
                        int rowNumber = i + 2;
                        summaryMessages.Add($"-- Fila {rowNumber} --");

                        var dict = request.UsersRaw[i];
                        string get(string key) =>
                            dict.TryGetValue(key, out var v) && v != null ? v.ToString().Trim() : "";

                        var model = new AltaUsuarioController.UserModelAltaUsuario
                        {
                            Nombre = get("Nombre"),
                            Apellido1 = get("Apellido1"),
                            Apellido2 = get("Apellido2"),
                            DNI = get("DNI"),
                            OUPrincipal = get("OUPrincipal"),
                            OUSecundaria = get("OUSecundaria"),
                            Cuota = get("Cuota"),
                            NTelefono = string.IsNullOrEmpty(get("nTelefono")) ? null : get("nTelefono"),
                            ExtensionMovil = string.IsNullOrEmpty(get("MobileExt")) ? null : get("MobileExt"),
                            NumeroLargoMovil = string.IsNullOrEmpty(get("MobileNumber")) ? null : get("MobileNumber"),
                            TarjetaIdentificativa = string.IsNullOrEmpty(get("TarjetaId")) ? null : get("TarjetaId"),
                            NFuncionario = string.IsNullOrEmpty(get("nFuncionario")) ? null : get("nFuncionario")
                        };

                        // 2.2) Validar OU principal
                        if (!OuPrincipalExiste(model.OUPrincipal))
                        {
                            summaryMessages.Add($"Fila {rowNumber}: Error: la OU principal '{model.OUPrincipal}' no existe.");
                            overallSuccess = false;
                            continue;
                        }

                        // 2.3) Validar OU secundaria (si se indicó)
                        if (!string.IsNullOrEmpty(model.OUSecundaria)
                            && !OuSecundariaExiste(model.OUPrincipal, model.OUSecundaria))
                        {
                            summaryMessages.Add($"Fila {rowNumber}: Error: la OU secundaria '{model.OUSecundaria}' no existe bajo '{model.OUPrincipal}'.");
                            overallSuccess = false;
                            continue;
                        }

                        // 2.4) Unicidad: DNI
                        if (!string.IsNullOrEmpty(model.DNI))
                        {
                            dynamic dniData = (_altaUsuarioController.CheckDNIExists(
                                new Dictionary<string, string> { ["dni"] = model.DNI }) as JsonResult)?.Value;
                            if (dniData?.success == true && (bool)dniData.exists)
                            {
                                summaryMessages.Add($"Fila {rowNumber}: Error: el DNI '{model.DNI}' ya existe.");
                                overallSuccess = false;
                                continue;
                            }
                        }

                        // 2.5) Unicidad: Teléfono fijo
                        if (!string.IsNullOrEmpty(model.NTelefono))
                        {
                            dynamic telData = (_altaUsuarioController.CheckTelephoneExists(
                                new Dictionary<string, string> { ["nTelefono"] = model.NTelefono }) as JsonResult)?.Value;
                            if (telData?.success == true && (bool)telData.exists)
                            {
                                summaryMessages.Add($"Fila {rowNumber}: Error: el teléfono '{model.NTelefono}' ya existe.");
                                overallSuccess = false;
                                continue;
                            }
                        }

                        // 2.6) Unicidad: Nº Funcionario
                        if (!string.IsNullOrEmpty(model.NFuncionario))
                        {
                            dynamic funData = (_altaUsuarioController.CheckNumberIdExists(
                                new Dictionary<string, string> { ["nFuncionario"] = model.NFuncionario }) as JsonResult)?.Value;
                            if (funData?.success == true && (bool)funData.exists)
                            {
                                summaryMessages.Add($"Fila {rowNumber}: Error: el número de funcionario '{model.NFuncionario}' ya existe.");
                                overallSuccess = false;
                                continue;
                            }
                        }

                        // 2.7) Leer directamente el array enviado desde el cliente
                        List<string> grupos = new List<string>();
                        if (dict.TryGetValue("Grupos", out var grObj) && grObj is JsonElement grEl && grEl.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var item in grEl.EnumerateArray())
                                if (item.ValueKind == JsonValueKind.String)
                                    grupos.Add(item.GetString());
                        }
                        model.Grupos = grupos;


                        // 2.8) Parseo de FechaCaducidad (dd/MM/yyyy o ISO yyyy-MM-dd)
                        var rawFecha = get("FechaCaducidad");
                        if (!string.IsNullOrEmpty(rawFecha))
                        {
                            DateTime parsedDate;
                            if (DateTime.TryParseExact(rawFecha, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate)
                             || DateTime.TryParse(rawFecha, out parsedDate))
                            {
                                model.FechaCaducidadOp = "si";
                                model.FechaCaducidad = parsedDate;
                            }
                            else
                            {
                                model.FechaCaducidadOp = "no";
                            }
                        }
                        else
                        {
                            model.FechaCaducidadOp = "no";
                        }

                        // 2.9) Generar Username
                        dynamic genData = (_altaUsuarioController.GenerateUsername(
                            new userInputModel
                            {
                                Nombre = model.Nombre,
                                Apellido1 = model.Apellido1,
                                Apellido2 = model.Apellido2
                            }) as JsonResult)?.Value;
                        model.Username = genData?.success == true ? genData.username : null;
                        if (string.IsNullOrEmpty(model.Username))
                        {
                            summaryMessages.Add($"Fila {rowNumber}: Error: no se pudo generar nombre de usuario.");
                            overallSuccess = false;
                            continue;
                        }

                        // 2.10) Obtener Departamento y LugarEnvio
                        dynamic depData = (_altaUsuarioController.GetDepartamento(
                            new Dictionary<string, string>
                            {
                                ["ouPrincipal"] = model.OUPrincipal,
                                ["ouSecundaria"] = model.OUSecundaria
                            }) as JsonResult)?.Value;
                        model.Departamento = depData?.success == true ? depData.departamento : null;

                        dynamic lugData = (_altaUsuarioController.GetLugarEnvio(
                            new Dictionary<string, string>
                            {
                                ["ouPrincipal"] = model.OUPrincipal,
                                ["ouSecundaria"] = model.OUSecundaria
                            }) as JsonResult)?.Value;
                        model.LugarEnvio = lugData?.success == true ? lugData.lugarEnvio : null;

                        // --- creación ---
                        dynamic createData = (_altaUsuarioController.CreateUser(model) as JsonResult)?.Value;
                        bool created = createData?.success ?? false;
                        if (created)
                        {
                            summaryMessages.Add($"Fila {rowNumber}: Éxito");
                            createdUsernames.Add(model.Username);
                        }
                        else
                        {
                            summaryMessages.Add($"Fila {rowNumber}: Error: {createData?.message ?? "desconocido"}");
                            overallSuccess = false;
                        }
                    }

                    if (createdUsernames.Any())
                    {
                        summaryMessages.Add("▶ Asignando licencias de Exchange a usuarios creados...");
                        // 3) Asignar licencias (iteramos sobre una copia)
                        var usersForLicenses = createdUsernames.ToList();
                        var failedLicenseUsers = new List<string>();

                        foreach (var user in usersForLicenses)
                        {
                            try
                            {
                                _altaUsuarioController.AddUserToGroup(user);
                                summaryMessages.Add($"Licencia asignada a '{user}'.");
                            }
                            catch (Exception ex)
                            {
                                summaryMessages.Add($"Error asignando licencia a '{user}': {ex.Message}");
                                failedLicenseUsers.Add(user);
                                overallSuccess = false;
                            }
                        }

                        // 2) Forzar sincronización Azure AD Connect
                        //var (syncOk, syncErr) = _altaUsuarioController.SyncDeltaOnVanGogh();
                        //if (!syncOk)
                        //{
                        //    summaryMessages.Add("Error al sincronizar Azure AD Connect: " + syncErr);
                        //    throw new InvalidOperationException(syncErr);
                        //}

                        // 3) Comprobar si existe el último usuario de la lista en Exchange, lo que querría decir que existen todos
                        //var lastUser = createdUsernames.LastOrDefault();

                        //var exists = await _altaUsuarioController.WaitForAzureUser(lastUser);
                        //if (!exists)
                        //{
                        //    summaryMessages.Add($"Timeout esperando al usuario '{lastUser}' en Azure AD. Abortando creación de buzón y correo.");
                        //    return Json(new
                        //    {
                        //        success = false,
                        //        message = "Aborto: el usuario no apareció en Azure AD antes del timeout.",
                        //        messages = summaryMessages,
                        //        created = createdUsernames
                        //    });
                        //}

                        // 4) Crear buzón on prem
                        //foreach (var user in usersForLicenses)
                        //{
                        //    try
                        //    {

                        //        _altaUsuarioController.EnableOnPremMailbox(user, adminUsername, adminPassword);
                        //        summaryMessages.Add($"Buzón creado para '{user}'.");
                        //    }
                        //    catch (Exception ex)
                        //    {
                        //        summaryMessages.Add($"Error creando buzón para '{user}': {ex.Message}");
                        //        createdUsernames.Remove(user);
                        //        overallSuccess = false;
                        //    }
                        //}

                        // 5) Actualizar proxyaddresses en AD
                        foreach (var user in usersForLicenses)
                        {
                            try
                            {
                                _altaUsuarioController.UpdateProxyAddresses(user);
                                summaryMessages.Add($"Proxy addresses actualizadas para '{user}'.");
                            }
                            catch (Exception ex)
                            {
                                summaryMessages.Add($"Error actualizando proxy addresses para '{user}': {ex.Message}");
                                createdUsernames.Remove(user);
                                overallSuccess = false;
                            }
                        }

                        //Crear batch de migración

                        //try
                        //{
                        //    string[] listaUsuarios = createdUsernames.ToArray();
                        //    _altaUsuarioController.CreateMigrationBatch(listaUsuarios);
                        //    summaryMessages.Add("Batch de migración creado con éxito.");
                        //}
                        //catch (Exception ex)
                        //{
                        //    summaryMessages.Add($"Error creando batch de migración: {ex.Message}");
                        //    overallSuccess = false;
                        //}

                        // Capturar errores de RunImpersonated o de inicialización
                        overallSuccess = true;
                        finalResult = Json(new
                        {
                            success = true,
                            message = $"Éxito en altas masivas: " + summaryMessages,
                            log = summaryMessages
                        });
                    }
                
                }
                catch (Exception exInner)
                {
                    // Capturar cualquier excepción no esperada dentro de la impersonación
                    overallSuccess = false;
                    summaryMessages.Add($"[ERROR FATAL]: {exInner.Message}");
                    finalResult = Json(new
                    {
                        success = false,
                        message = $"Error inesperado en alta masiva: {exInner.Message}",
                        log = summaryMessages
                    });
                }
            });

        }
        catch (Exception exOuter)
        {
            // Capturar errores de RunImpersonated o de inicialización
            overallSuccess = false;
            summaryMessages.Add($"[ERROR EXTERNO]: {exOuter.Message}");
            finalResult = Json(new
            {
                success = false,
                message = $"Error exterior en alta masiva: {exOuter.Message}",
                log = summaryMessages
            });
        }
        finally
        {
            CloseHandle(userToken);
        }

        // 5) Devolver SIEMPRE un JsonResult válido
        return finalResult;
    }


    
    /// Divide la cadena cruda de grupos separada por ';' y devuelve la lista limpia.
    /// Si la cadena está vacía, devuelve lista vacía.
    /// </summary>
    private List<string> ParseGrupos(string rawGrupos)
    {
        if (string.IsNullOrWhiteSpace(rawGrupos))
            return new List<string>();

        return rawGrupos
            .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(g => g.Trim())
            .Where(g => !string.IsNullOrEmpty(g))
            .ToList();
    }



    private async Task<string> GetOUPath(string ouName, string parentOU)
    {
        string domainPath = _config["ActiveDirectory:DomainComponents"];
            try
        {
            using (var rootEntry = new DirectoryEntry($"LDAP://{domainPath}"))
            {
                using (var searcher = new DirectorySearcher(rootEntry))
                {
                    string filter;
                    if (!string.IsNullOrEmpty(parentOU))
                    {
                        // Buscar OU secundaria dentro de OU principal
                        filter = $"(&(objectClass=organizationalUnit)(ou={ouName})(distinguishedName=OU={ouName},OU=Usuarios,OU={parentOU},{_config["ActiveDirectory:DomainBase"]}";
                    }
                    else
                    {
                        // Buscar OU principal bajo AREAS
                        filter = $"(&(objectClass=organizationalUnit)(ou={ouName})(distinguishedName=OU={ouName},{_config["ActiveDirectory:DomainBase"]}))";
                    }

                    searcher.Filter = filter;
                    searcher.SearchScope = SearchScope.Subtree;
                    searcher.PropertiesToLoad.Add("distinguishedName");

                    var result = searcher.FindOne();
                    if (result != null && result.Properties["distinguishedName"].Count > 0)
                    {
                        return result.Properties["distinguishedName"][0]?.ToString();
                    }
                    return null;
                }
            }
        }
        catch
        {
            return null;
        }
    }

    private async Task<string> GetOUAttribute(string ouPath, string attribute)
    {
        try
        {
            using (var ouEntry = new DirectoryEntry($"LDAP://{ouPath}"))
            {
                using (var searcher = new DirectorySearcher(ouEntry))
                {
                    searcher.Filter = "(objectClass=organizationalUnit)";
                    searcher.PropertiesToLoad.Add(attribute);
                    var result = searcher.FindOne();
                    if (result != null && result.Properties[attribute].Count > 0)
                    {
                        return result.Properties[attribute][0]?.ToString();
                    }
                    return null;
                }
            }
        }
        catch
        {
            return null;
        }
    }

    private async Task<string> GenerateUsername(string nombre, string apellido1, string apellido2)
    {
        try
        {
            // Normalizar entradas
            nombre = RemoveAccents(nombre.Trim().ToLower());
            apellido1 = RemoveAccents(apellido1.Trim().ToLower());
            apellido2 = RemoveAccents(apellido2.Trim().ToLower());

            // Construir candidatos
            List<string> candidatos = new List<string>();

            // 1. Inicial del nombre + apellido1 + inicial del apellido2
            string candidato1 = $"{nombre[0]}{apellido1}{apellido2[0]}";
            candidatos.Add(candidato1.Substring(0, Math.Min(12, candidato1.Length)));

            // 2. Inicial del nombre + apellido1
            string candidato2 = $"{nombre[0]}{apellido1}";
            candidatos.Add(candidato2.Substring(0, Math.Min(12, candidato2.Length)));

            // 3. Apellido1 + inicial del apellido2
            string candidato3 = $"{apellido1}{apellido2[0]}";
            candidatos.Add(candidato3.Substring(0, Math.Min(12, candidato3.Length)));

            // Verificar existencia
            foreach (string candidato in candidatos)
            {
                if (!await CheckUserInActiveDirectory(candidato))
                {
                    return candidato;
                }
            }

            // Intentar con sufijos numéricos
            for (int i = 1; i <= 100; i++)
            {
                string candidatoConNumero = $"{candidatos[0]}{i}";
                if (!await CheckUserInActiveDirectory(candidatoConNumero))
                {
                    return candidatoConNumero.Substring(0, Math.Min(12, candidatoConNumero.Length));
                }
            }

            return null;
        }
        catch
        {
            return null;
        }
    }

    private async Task<bool> CheckUserInActiveDirectory(string username)
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
            return true; // Asumir que existe en caso de error
        }
    }

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

    /// <summary>
    /// Comprueba si existe la OU principal bajo {_config["ActiveDirectory:DomainBase"]}
    /// </summary>
    private bool OuPrincipalExiste(string ouPrincipal)
    {
        if (string.IsNullOrWhiteSpace(ouPrincipal))
            return false;

        // Base de búsqueda: {_config["ActiveDirectory:DomainBase"]}
        string ldapPath = $"LDAP://{_config["ActiveDirectory:DomainBase"]}";
        using var entry = new DirectoryEntry(ldapPath);
        using var searcher = new DirectorySearcher(entry)
        {
            Filter = $"(&(objectClass=organizationalUnit)(ou={ouPrincipal}))",
            SearchScope = SearchScope.OneLevel
        };

        return searcher.FindOne() != null;
    }

    /// <summary>
    /// Comprueba si existe la OU secundaria bajo
    ///     OU=Usuarios,OU={ouPrincipal},{_config["ActiveDirectory:DomainBase"]}
    /// </summary>
    private bool OuSecundariaExiste(string ouPrincipal, string ouSecundaria)
    {
        if (string.IsNullOrWhiteSpace(ouSecundaria))
            return true;  // no es obligatorio

        // Base de búsqueda: OU=Usuarios bajo la OU principal
        string ldapPath = $"LDAP://OU=Usuarios,OU={ouPrincipal},{_config["ActiveDirectory:DomainBase"]}";
        using var entry = new DirectoryEntry(ldapPath);
        using var searcher = new DirectorySearcher(entry)
        {
            Filter = $"(&(objectClass=organizationalUnit)(ou={ouSecundaria}))",
            SearchScope = SearchScope.OneLevel
        };

        return searcher.FindOne() != null;
    }

    /// <summary>
    /// Escapa los caracteres especiales según RFC2254 para usar en filtros LDAP.
    /// </summary>
    private string EscapeLdapSearchFilter(string input)
    {
        if (input == null) return null;
        return input
            .Replace("\\", "\\5c")
            .Replace("*", "\\2a")
            .Replace("(", "\\28")
            .Replace(")", "\\29")
            .Replace("\0", "\\00");
    }


    /// <summary>
    /// Comprueba si existe un grupo en AD usando PrincipalContext en el dominio aytosa.inet
    /// </summary>
    private bool GroupExists(string groupName)
    {
        if (string.IsNullOrWhiteSpace(groupName))
            return false;

        using var ctx = new PrincipalContext(ContextType.Domain, "aytosa.inet");

        // Primero comprueba por SamAccountName
        var grp = GroupPrincipal.FindByIdentity(
            ctx,
            IdentityType.SamAccountName,
            groupName
        );
        if (grp != null) return true;

        // Luego por CN
        grp = GroupPrincipal.FindByIdentity(
            ctx,
            IdentityType.Name,
            groupName
        );
        return grp != null;
    }


    /// <summary>
    /// Dada una lista de nombres de grupo, devuelve aquellos que NO existen en AD.
    /// </summary>
    private List<string> GetMissingGroups(IEnumerable<string> grupos)
    {
        var missing = new List<string>();
        foreach (var g in grupos)
        {
            if (!GroupExists(g))
                missing.Add(g);
        }
        return missing;
    }

}