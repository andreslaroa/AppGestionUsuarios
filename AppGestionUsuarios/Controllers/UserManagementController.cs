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

public class UserManagementController : Controller
{
    private readonly OUService _ouService;
    private readonly ILogger<UserManagementController> _logger;

    public class UserModel
    {
        public string Nombre { get; set; }
        public string Apellido1 { get; set; }
        public string Apellido2 { get; set; }
        public string Username { get; set; }
        public string NFuncionario { get; set; }
        public string OUPrincipal { get; set; }
        public string OUSecundaria { get; set; }
        public string Departamento { get; set; }
    }


    public UserManagementController()
    {
        string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Resources", "ArchivoDePruebasOU.xlsx");
        _ouService = new OUService(filePath);
    }

    [HttpGet]
    public IActionResult LoginSuccess()
    {
        var ouPrincipales = _ouService.GetOUPrincipales();
        ViewBag.OUPrincipales = ouPrincipales;
        var portalEmpleado = _ouService.GetPortalEmpleado();
        ViewBag.portalEmpleado = portalEmpleado;
        var cuota = _ouService.GetCuota();
        ViewBag.cuota = cuota;
        return View();
    }

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

    [HttpPost]
    public IActionResult CreateUser([FromBody] UserModel user)
    {
        // Validar si los datos se recibieron correctamente
        if (user == null)
        {
            return Json(new { success = false, message = "No se recibieron datos válidos." });
        }

        // Validar los campos obligatorios
        if (string.IsNullOrEmpty(user.Nombre) || string.IsNullOrEmpty(user.Apellido1) || string.IsNullOrEmpty(user.Username) ||
            string.IsNullOrEmpty(user.OUPrincipal) || string.IsNullOrEmpty(user.OUSecundaria) || string.IsNullOrEmpty(user.Departamento))
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

            // Construir el path LDAP
            string ldapPath = $"LDAP://OU={user.OUSecundaria},OU=Usuarios y Grupos,OU={user.OUPrincipal},DC=aytosa,DC=inet";
            Console.WriteLine($"Intentando conectar a: {ldapPath}");

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
                    newUser.Properties["sAMAccountName"].Value = user.Username; // Nombre de usuario corto
                    newUser.Properties["userPrincipalName"].Value = $"{user.Username}@aytosa.inet"; // Dominio
                    newUser.Properties["displayName"].Value = displayName; // Nombre completo
                    newUser.Properties["description"].Value = $"Nº Funcionario: {user.NFuncionario}"; // Descripción

                    // Guardar cambios iniciales
                    newUser.CommitChanges();

                    // Establecer la contraseña
                    newUser.Invoke("SetPassword", new object[] { "Temporal2024" });

                    // Activar la cuenta
                    newUser.Properties["userAccountControl"].Value = 0x200;

                    // Forzar el cambio de contraseña al primer inicio de sesión
                    newUser.Properties["pwdLastSet"].Value = 0;

                    // Guardar cambios finales
                    newUser.CommitChanges();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error al crear el usuario: {ex.Message}");
                    return Json(new { success = false, message = $"Error al crear el usuario: {ex.Message}" });
                }
                finally
                {
                    if (newUser != null)
                    {
                        newUser.Dispose();
                    }
                }

                return Json(new { success = true, message = "Usuario creado exitosamente." });
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}\nStackTrace: {ex.StackTrace}");
            return Json(new { success = false, message = $"Error al crear el usuario: {ex.Message}" });
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



}
