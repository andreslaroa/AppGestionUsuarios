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
        var messages = new List<string>();
        bool overallSuccess = true;

        // 1) Validación inicial
        if (usersRaw == null || usersRaw.Count == 0)
        {
            return Json(new
            {
                success = false,
                message = "No se han enviado usuarios para procesar.",
                messages
            });
        }
        messages.Add($"▶ ProcessUsers: recibidas {usersRaw.Count} filas.");

        // 2) Iteramos cada fila del Excel
        for (int i = 0; i < usersRaw.Count; i++)
        {
            int fila = i + 2;
            messages.Add($"-- Fila {fila} --");

            var dict = usersRaw[i];
            string get(string key) =>
                dict.TryGetValue(key, out var v) && v != null
                    ? v.ToString().Trim()
                    : "";

            try
            {
                // 2.1) Mapear campos
                var model = new AltaUsuarioController.UserModelAltaUsuario
                {
                    Nombre = get("Nombre"),
                    Apellido1 = get("Apellido1"),
                    Apellido2 = get("Apellido2"),
                    DNI = get("DNI"),
                    OUPrincipal = get("OUPrincipal"),
                    OUSecundaria = get("OUSecundaria"),
                    Cuota = get("Cuota"),
                    // Opcionales
                    NTelefono = string.IsNullOrEmpty(get("nTelefono")) ? null : get("nTelefono"),
                    ExtensionMovil = string.IsNullOrEmpty(get("MobileExt")) ? null : get("MobileExt"),
                    NumeroLargoMovil = string.IsNullOrEmpty(get("MobileNumber")) ? null : get("MobileNumber"),
                    TarjetaIdentificativa = string.IsNullOrEmpty(get("TarjetaId")) ? null : get("TarjetaId"),
                    NFuncionario = string.IsNullOrEmpty(get("nFuncionario")) ? null : get("nFuncionario"),
                    // Grupos separados por espacio
                    Grupos = get("Grupos")
                                                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                                                .ToList()
                };
                messages.Add($"  Modelo mapeado: {model.Nombre} {model.Apellido1} {model.Apellido2}");

                // 2.2) Parsear fecha de caducidad dd/MM/yyyy
                var rawFecha = get("FechaCaducidad");
                if (DateTime.TryParseExact(rawFecha, "dd/MM/yyyy",
                                           CultureInfo.InvariantCulture,
                                           DateTimeStyles.None,
                                           out var fechaCad))
                {
                    model.FechaCaducidadOp = "si";
                    model.FechaCaducidad = fechaCad;
                    messages.Add($"  FechaCaducidad parseada: {fechaCad:dd/MM/yyyy}");
                }
                else
                {
                    model.FechaCaducidadOp = "no";
                    messages.Add("  Sin FechaCaducidad");
                }

                // 2.3) Generar Username automáticamente
                messages.Add("  Generando Username...");
                var userInput = new userInputModel
                {
                    Nombre = model.Nombre,
                    Apellido1 = model.Apellido1,
                    Apellido2 = model.Apellido2
                };
                var genJson = _altaUsuarioController.GenerateUsername(userInput) as JsonResult;
                dynamic genData = genJson?.Value;
                model.Username = (genData?.success == true) ? genData.username : null;
                messages.Add($"  Username generado: '{model.Username}'");
                if (string.IsNullOrEmpty(model.Username))
                    messages.Add("  ⚠️ Username vacío, CreateUser podría fallar.");

                // 2.4) Obtener Departamento desde AltaUsuarioController
                messages.Add("  Obteniendo Departamento...");
                var depJson = _altaUsuarioController.GetDepartamento(
                    new Dictionary<string, string>
                    {
                        ["ouPrincipal"] = model.OUPrincipal,
                        ["ouSecundaria"] = model.OUSecundaria
                    }) as JsonResult;
                dynamic depData = depJson?.Value;
                model.Departamento = (depData?.success == true) ? depData.departamento : null;
                messages.Add($"  Departamento: '{model.Departamento}'");

                // 2.5) Obtener LugarEnvio
                messages.Add("  Obteniendo LugarEnvio...");
                var lugJson = _altaUsuarioController.GetLugarEnvio(
                    new Dictionary<string, string>
                    {
                        ["ouPrincipal"] = model.OUPrincipal,
                        ["ouSecundaria"] = model.OUSecundaria
                    }) as JsonResult;
                dynamic lugData = lugJson?.Value;
                model.LugarEnvio = (lugData?.success == true) ? lugData.lugarEnvio : null;
                messages.Add($"  LugarEnvio: '{model.LugarEnvio}'");

                // 2.6) Llamar a CreateUser
                messages.Add("  Llamando a CreateUser...");
                var createJson = _altaUsuarioController.CreateUser(model) as JsonResult;
                dynamic createData = createJson?.Value;
                bool created = createData?.success ?? false;
                string msg = createData?.message ?? "[sin mensaje]";
                messages.Add($"  CreateUser → success={created}, message='{msg}'");
                if (!created) overallSuccess = false;
            }
            catch (Exception ex)
            {
                messages.Add($"  ❌ Excepción fila {fila}: {ex.GetType().Name}: {ex.Message}");
                overallSuccess = false;
            }
        }

        messages.Add("▶ Fin de ProcessUsers");

        return Json(new
        {
            success = overallSuccess,
            message = overallSuccess
                       ? "Alta masiva completada con éxito (ver detalles)."
                       : "Se produjeron errores en la alta masiva (ver detalles).",
            messages
        });
    }



    // ---------------------------------------------------
    // Helper para convertir la cadena de grupos en List<string>
    // ---------------------------------------------------
    private List<string> ParseGrupos(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return new List<string>();
        return raw
            .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(g => g.Trim())
            .Where(g => g.Length > 0)
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
}