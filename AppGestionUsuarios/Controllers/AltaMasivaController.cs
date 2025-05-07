using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Globalization;
using static GestionUsuariosController;

[Authorize]
public class AltaMasivaController : Controller
{

    private readonly AltaUsuarioController _altaUsuarioController;

    public AltaMasivaController(AltaUsuarioController altaUsuarioController)
    {
        Console.WriteLine("Ctor AltaMasivaController invocado");
        _altaUsuarioController = altaUsuarioController
            ?? throw new ArgumentNullException(nameof(altaUsuarioController));
    }


    private readonly string domainPath = "DC=aytosa,DC=inet";

    // Clase para el resultado de las verificaciones
    private class CheckResult
    {
        public bool Success { get; set; }
        public bool Exists { get; set; }
        public string Message { get; set; }
    }

    [HttpGet]
    public IActionResult AltaMasiva()
    {
        try
        {
            return View("AltaMasiva");
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
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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

    // POST: /AltaMasiva/ProcessUsers
    // Mapea cada fila a UserModelAltaUsuario, inyecta Departamento y LugarEnvio y llama a CreateUser
    [HttpPost]
    public JsonResult ProcessUsers([FromBody] List<Dictionary<string, object>> usersRaw)
    {
        var summaryMessages = new List<string>();
        bool overallSuccess = true;

        // 1) Validación inicial
        if (usersRaw == null || usersRaw.Count == 0)
        {
            return Json(new
            {
                success = false,
                message = "No se han enviado usuarios para procesar.",
                messages = summaryMessages
            });
        }
        summaryMessages.Add($"▶ ProcessUsers: recibidas {usersRaw.Count} filas.");

        // 2) Iterar cada fila
        for (int i = 0; i < usersRaw.Count; i++)
        {
            int rowNumber = i + 2;
            summaryMessages.Add($"-- Fila {rowNumber} --");

            var dict = usersRaw[i];
            string get(string key) =>
                dict.TryGetValue(key, out var v) && v != null
                    ? v.ToString().Trim()
                    : "";

            // 2.1) Mapear modelo básico
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

            // 2.7) Parseo y validación de Grupos
            var grupos = get("Grupos")
                .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => s.Length > 0)
                .ToList();
            model.Grupos = grupos;

            var missing = GetMissingGroups(grupos);
            if (missing.Any())
            {
                summaryMessages.Add($"Fila {rowNumber}: Error: no existen en AD los grupos [{string.Join(", ", missing)}].");
                overallSuccess = false;
                continue;
            }

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

            // 2.11) Crear usuario
            dynamic createData = (_altaUsuarioController.CreateUser(model) as JsonResult)?.Value;
            bool created = createData?.success ?? false;
            if (created)
            {
                summaryMessages.Add($"Fila {rowNumber}: Éxito");
            }
            else
            {
                summaryMessages.Add($"Fila {rowNumber}: Error: {createData?.message ?? "desconocido"}");
                overallSuccess = false;
            }
        }

        // 3) Devolver resultado
        return Json(new
        {
            success = overallSuccess,
            message = overallSuccess
                       ? "Alta masiva completada con éxito."
                       : "Se produjeron errores en algunas filas.",
            messages = summaryMessages
        });
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
                        filter = $"(&(objectClass=organizationalUnit)(ou={ouName})(distinguishedName=OU={ouName},OU=Usuarios y Grupos,OU={parentOU},OU=AREAS,DC=aytosa,DC=inet))";
                    }
                    else
                    {
                        // Buscar OU principal bajo AREAS
                        filter = $"(&(objectClass=organizationalUnit)(ou={ouName})(distinguishedName=OU={ouName},OU=AREAS,DC=aytosa,DC=inet))";
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
    /// Comprueba si existe la OU principal bajo OU=AREAS,DC=aytosa,DC=inet
    /// </summary>
    private bool OuPrincipalExiste(string ouPrincipal)
    {
        if (string.IsNullOrWhiteSpace(ouPrincipal))
            return false;

        // Base de búsqueda: OU=AREAS,DC=aytosa,DC=inet
        using var entry = new DirectoryEntry("LDAP://OU=AREAS,DC=aytosa,DC=inet");
        using var searcher = new DirectorySearcher(entry)
        {
            Filter = $"(&(objectClass=organizationalUnit)(ou={ouPrincipal}))",
            SearchScope = SearchScope.OneLevel
        };

        return searcher.FindOne() != null;
    }

    /// <summary>
    /// Comprueba si existe la OU secundaria bajo
    ///     OU=Usuarios y Grupos,OU={ouPrincipal},OU=AREAS,DC=aytosa,DC=inet
    /// </summary>
    private bool OuSecundariaExiste(string ouPrincipal, string ouSecundaria)
    {
        if (string.IsNullOrWhiteSpace(ouSecundaria))
            return true;  // no es obligatorio

        // Base de búsqueda: OU=Usuarios y Grupos bajo la OU principal
        string ldapPath = $"LDAP://OU=Usuarios y Grupos,OU={ouPrincipal},OU=AREAS,DC=aytosa,DC=inet";
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
    /// Devuelve la lista de grupos que NO se encontraron en el AD.
    /// </summary>
    private List<string> GetMissingGroups(IEnumerable<string> grupos)
    {
        var missing = new List<string>();
        if (grupos == null || !grupos.Any())
            return missing;

        // 1) Obtén el contexto de naming (dominio) dinámicamente
        string namingContext;
        using (var rootDse = new DirectoryEntry("LDAP://RootDSE"))
        {
            namingContext = rootDse.Properties["defaultNamingContext"][0].ToString();
        }

        // 2) Crea el DirectoryEntry raíz
        using var domainEntry = new DirectoryEntry($"LDAP://{namingContext}");
        using var searcher = new DirectorySearcher(domainEntry)
        {
            SearchScope = SearchScope.Subtree
        };

        // 3) Solo necesitamos buscar por cn y objectCategory=group
        searcher.PropertiesToLoad.Clear();
        searcher.PropertiesToLoad.Add("cn");

        foreach (var grupo in grupos)
        {
            if (string.IsNullOrWhiteSpace(grupo))
                continue;

            var esc = EscapeLdapSearchFilter(grupo);
            searcher.Filter = $"(&(objectCategory=group)(cn={esc}))";

            var result = searcher.FindOne();
            if (result == null)
                missing.Add(grupo);
        }

        return missing;
    }
}