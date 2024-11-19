using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.IO;
using System.DirectoryServices.AccountManagement;
using System;
using System.DirectoryServices;
using System.Globalization;
using System.Linq;
using System.Text;

public class UserManagementController : Controller
{
    private readonly OUService _ouService;

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

    // Nuevo método para verificar la existencia de un nombre de usuario en el Directorio Activo
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
            // Manejo de errores (opcional)
            return true; // Asumir que el usuario existe si hay un error
        }
    }

    [HttpPost]
    public IActionResult CreateUser(string nombre, string apellido1, string apellido2, string username, string nFuncionario, string ouPrincipal, string ouSecundaria, string departamento)
    {
        try
        {
            // Convertir nombre y apellidos a mayúsculas y eliminar acentos
            string nombreUpper = RemoveAccents(nombre).ToUpperInvariant();
            string apellido1Upper = RemoveAccents(apellido1).ToUpperInvariant();
            string apellido2Upper = RemoveAccents(apellido2).ToUpperInvariant();

            // Conformar el nombre completo
            string displayName = $"{nombreUpper} {apellido1Upper} {apellido2Upper}".Trim();

            // Ruta a la OU en el directorio activo
            string ldapPath = $"LDAP://OU={ouSecundaria},OU={ouPrincipal},DC=midominio,DC=local"; // Ajusta 'midominio' y 'local' según tu entorno

            using (DirectoryEntry ouEntry = new DirectoryEntry(ldapPath))
            {
                using (DirectoryEntry newUser = ouEntry.Children.Add($"CN={displayName}", "user"))
                {
                    // Establecer atributos del usuario
                    newUser.Properties["sAMAccountName"].Value = username;
                    newUser.Properties["userPrincipalName"].Value = $"{username}@midominio.local"; // Ajusta el dominio
                    newUser.Properties["displayName"].Value = displayName;
                    newUser.Properties["description"].Value = $"Nº Funcionario: {nFuncionario}";
                    newUser.Properties["physicalDeliveryOfficeName"].Value = departamento; // Atributo oficina (departamento)

                    // Commit de los cambios
                    newUser.CommitChanges();

                    // Establecer la contraseña
                    newUser.Invoke("SetPassword", new object[] { "Temporal2024" });

                    // Habilitar al usuario
                    newUser.Properties["userAccountControl"].Value = 0x200; // Habilitar cuenta

                    // Forzar el cambio de contraseña al primer inicio de sesión
                    newUser.Properties["pwdLastSet"].Value = 0;

                    // Guardar cambios finales
                    newUser.CommitChanges();
                }
            }

            return Json(new { success = true, message = "Usuario creado exitosamente." });
        }
        catch (Exception ex)
        {
            // Manejo de errores
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
