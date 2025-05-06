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

[Authorize]
public class AltaMasivaController : Controller
{
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

    [HttpPost]
    public async Task<IActionResult> ProcessUsers(IFormFile file)
    {
        if (file == null || file.Length == 0)
            return Json(new { success = false, message = "No se ha proporcionado un archivo Excel válido." });

        List<Dictionary<string, object>> users = new List<Dictionary<string, object>>();
        List<string> messages = new List<string>();
        bool overallSuccess = true;

        try
        {
            // Configurar EPPlus para uso no comercial
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    // Obtener encabezados
                    var headers = new List<string>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        headers.Add(worksheet.Cells[1, col].Text.Trim());
                    }

                    // Validar encabezados requeridos
                    if (!headers.Contains("Nombre") || !headers.Contains("Apellido1") || !headers.Contains("Apellido2") ||
                        !headers.Contains("DNI") || !headers.Contains("OUPrincipal"))
                    {
                        return Json(new { success = false, message = "El archivo Excel debe contener las columnas 'Nombre', 'Apellido1', 'Apellido2', 'DNI' y 'OUPrincipal'." });
                    }

                    // Leer datos
                    for (int row = 2; row <= rowCount; row++)
                    {
                        var user = new Dictionary<string, object>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            user[headers[col - 1]] = worksheet.Cells[row, col].Text.Trim();
                        }
                        users.Add(user);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al procesar el archivo Excel: {ex.Message}" });
        }

        for (int i = 0; i < users.Count; i++)
        {
            var user = users[i];
            int lineNumber = i + 2;

            try
            {
                // Extraer datos
                string nombre = user.GetValueOrDefault("Nombre", "").ToString();
                string apellido1 = user.GetValueOrDefault("Apellido1", "").ToString();
                string apellido2 = user.GetValueOrDefault("Apellido2", "").ToString();
                string dni = user.GetValueOrDefault("DNI", "").ToString();
                string nTelefono = user.GetValueOrDefault("nTelefono", "").ToString();
                string ddi = user.GetValueOrDefault("DDI", "").ToString();
                string mobileExt = user.GetValueOrDefault("MobileExt", "").ToString();
                string mobileNumber = user.GetValueOrDefault("MobileNumber", "").ToString();
                string tarjetaId = user.GetValueOrDefault("TarjetaId", "").ToString();
                string nFuncionario = user.GetValueOrDefault("nFuncionario", "").ToString();
                string ouPrincipalName = user.GetValueOrDefault("OUPrincipal", "").ToString();
                string ouSecundariaName = user.GetValueOrDefault("OUSecundaria", "").ToString();
                string fechaCaducidad = user.GetValueOrDefault("FechaCaducidad", "").ToString();

                // Validaciones obligatorias
                if (string.IsNullOrEmpty(nombre))
                {
                    messages.Add($"Fila {lineNumber}: El campo 'Nombre' es obligatorio.");
                    overallSuccess = false;
                    continue;
                }
                if (string.IsNullOrEmpty(apellido1))
                {
                    messages.Add($"Fila {lineNumber}: El campo 'Apellido1' es obligatorio.");
                    overallSuccess = false;
                    continue;
                }
                if (string.IsNullOrEmpty(apellido2))
                {
                    messages.Add($"Fila {lineNumber}: El campo 'Apellido2' es obligatorio.");
                    overallSuccess = false;
                    continue;
                }
                if (string.IsNullOrEmpty(dni))
                {
                    messages.Add($"Fila {lineNumber}: El campo 'DNI' es obligatorio.");
                    overallSuccess = false;
                    continue;
                }
                if (string.IsNullOrEmpty(ouPrincipalName))
                {
                    messages.Add($"Fila {lineNumber}: El campo 'OUPrincipal' es obligatorio.");
                    overallSuccess = false;
                    continue;
                }

                // Validar OUs y construir path LDAP
                string ouPrincipalPath = await GetOUPath(ouPrincipalName, null);
                if (string.IsNullOrEmpty(ouPrincipalPath))
                {
                    messages.Add($"Fila {lineNumber}: La OU principal '{ouPrincipalName}' no existe.");
                    overallSuccess = false;
                    continue;
                }

                string ouSecundariaPath = string.Empty;
                if (!string.IsNullOrEmpty(ouSecundariaName))
                {
                    ouSecundariaPath = await GetOUPath(ouSecundariaName, ouPrincipalName);
                    if (string.IsNullOrEmpty(ouSecundariaPath))
                    {
                        messages.Add($"Fila {lineNumber}: La OU secundaria '{ouSecundariaName}' no existe dentro de '{ouPrincipalName}'.");
                        overallSuccess = false;
                        continue;
                    }
                }

                // Obtener atributos de la OU (st y department)
                string targetOUPath = !string.IsNullOrEmpty(ouSecundariaPath) ? ouSecundariaPath : ouPrincipalPath;
                string physicalDeliveryOfficeName = await GetOUAttribute(targetOUPath, "st");
                string division = await GetOUAttribute(targetOUPath, "department");

                string ldapPath = !string.IsNullOrEmpty(ouSecundariaPath)
                    ? $"LDAP://OU={ouSecundariaName},OU=Usuarios y Grupos,OU={ouPrincipalName},OU=AREAS,DC=aytosa,DC=inet"
                    : $"LDAP://OU=Usuarios y Grupos,OU={ouPrincipalName},OU=AREAS,DC=aytosa,DC=inet";

                // Generar username
                string username = await GenerateUsername(nombre, apellido1, apellido2);
                if (string.IsNullOrEmpty(username))
                {
                    messages.Add($"Fila {lineNumber}: No se pudo generar un nombre de usuario único para '{nombre} {apellido1} {apellido2}'.");
                    overallSuccess = false;
                    continue;
                }

                // Validar FechaCaducidad
                long? accountExpires = null;
                if (!string.IsNullOrEmpty(fechaCaducidad))
                {
                    if (DateTime.TryParseExact(fechaCaducidad, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime expiryDate))
                    {
                        // Convertir a formato de Active Directory (FILETIME)
                        accountExpires = expiryDate.ToFileTimeUtc();
                    }
                    else
                    {
                        messages.Add($"Fila {lineNumber}: El formato de 'FechaCaducidad' debe ser dd/mm/aaaa.");
                        overallSuccess = false;
                        continue;
                    }
                }

                bool userCreated = false;

                using (DirectoryEntry ouEntry = new DirectoryEntry(ldapPath))
                {
                    if (ouEntry == null)
                    {
                        messages.Add($"Fila {lineNumber}: No se pudo conectar a la OU especificada.");
                        overallSuccess = false;
                        continue;
                    }

                    DirectoryEntry newUser = null;

                    try
                    {
                        // Normalizar nombre
                        string nombreCompleto = $"{nombre} {apellido1} {apellido2}";
                        string displayName = RemoveAccents(nombreCompleto).ToUpperInvariant();

                        // Crear el usuario
                        newUser = ouEntry.Children.Add($"CN={displayName}", "user");
                        messages.Add($"Fila {lineNumber}: Usuario creado en la OU: CN={displayName}");

                        // Establecer atributos
                        try
                        {
                            newUser.Properties["givenName"].Value = nombre;
                            newUser.Properties["sn"].Value = $"{apellido1} {apellido2}";
                            newUser.Properties["sAMAccountName"].Value = username;
                            newUser.Properties["userPrincipalName"].Value = $"{username}@aytosa.inet";
                            newUser.Properties["displayName"].Value = displayName;
                            newUser.Properties["employeeNumber"].Value = dni; // Usar employeeNumber para DNI

                            if (!string.IsNullOrEmpty(nTelefono))
                                newUser.Properties["telephoneNumber"].Value = nTelefono;
                            if (!string.IsNullOrEmpty(mobileNumber))
                                newUser.Properties["mobile"].Value = $"{mobileExt} {mobileNumber}".Trim();
                            if (!string.IsNullOrEmpty(nFuncionario))
                                newUser.Properties["employeeID"].Value = nFuncionario;
                            if (!string.IsNullOrEmpty(ddi))
                                newUser.Properties["extensionAttribute1"].Value = ddi;
                            if (!string.IsNullOrEmpty(tarjetaId))
                                newUser.Properties["extensionAttribute2"].Value = tarjetaId;
                            if (!string.IsNullOrEmpty(physicalDeliveryOfficeName))
                                newUser.Properties["physicalDeliveryOfficeName"].Value = physicalDeliveryOfficeName;
                            if (!string.IsNullOrEmpty(division))
                                newUser.Properties["division"].Value = division;
                            if (accountExpires.HasValue)
                                newUser.Properties["accountExpires"].Value = accountExpires.Value;

                            messages.Add($"Fila {lineNumber}: Atributos establecidos - givenName: '{nombre}', sn: '{apellido1} {apellido2}', sAMAccountName: '{username}', userPrincipalName: '{username}@aytosa.inet', displayName: '{displayName}', employeeNumber: '{dni}'" +
                                (string.IsNullOrEmpty(nTelefono) ? "" : $", telephoneNumber: '{nTelefono}'") +
                                (string.IsNullOrEmpty(mobileNumber) ? "" : $", mobile: '{mobileExt} {mobileNumber}'") +
                                (string.IsNullOrEmpty(nFuncionario) ? "" : $", employeeID: '{nFuncionario}'") +
                                (string.IsNullOrEmpty(ddi) ? "" : $", extensionAttribute1: '{ddi}'") +
                                (string.IsNullOrEmpty(tarjetaId) ? "" : $", extensionAttribute2: '{tarjetaId}'") +
                                (string.IsNullOrEmpty(physicalDeliveryOfficeName) ? "" : $", physicalDeliveryOfficeName: '{physicalDeliveryOfficeName}'") +
                                (string.IsNullOrEmpty(division) ? "" : $", division: '{division}'") +
                                (accountExpires.HasValue ? $", accountExpires: '{fechaCaducidad}'" : ""));
                        }
                        catch (Exception ex)
                        {
                            messages.Add($"Fila {lineNumber}: Error al establecer atributos: {ex.Message}");
                            throw;
                        }

                        // Guardar cambios iniciales
                        try
                        {
                            newUser.CommitChanges();
                            messages.Add($"Fila {lineNumber}: Cambios iniciales guardados.");
                        }
                        catch (Exception ex)
                        {
                            messages.Add($"Fila {lineNumber}: Error al guardar cambios iniciales: {ex.Message}");
                            throw;
                        }

                        // Configurar contraseña y activar cuenta
                        try
                        {
                            newUser.Invoke("SetPassword", new object[] { "Temporal2024" });
                            newUser.Properties["userAccountControl"].Value = 0x200; // Cuenta normal activada
                            newUser.Properties["pwdLastSet"].Value = 0; // Forzar cambio de contraseña
                            newUser.CommitChanges();
                            userCreated = true;
                            messages.Add($"Fila {lineNumber}: Contraseña configurada y cuenta activada.");
                        }
                        catch (Exception ex)
                        {
                            messages.Add($"Fila {lineNumber}: Error al configurar contraseña: {ex.Message}");
                            throw;
                        }
                    }
                    catch (Exception ex)
                    {
                        messages.Add($"Fila {lineNumber}: Error al crear el usuario: {ex.Message}");
                    }
                    finally
                    {
                        newUser?.Dispose();
                    }
                }

                messages.Add(userCreated
                    ? $"Fila {lineNumber}: Usuario '{username}' creado exitosamente en la OU especificada."
                    : $"Fila {lineNumber}: No se pudo crear el usuario '{username}'.");
            }
            catch (Exception ex)
            {
                messages.Add($"Fila {lineNumber}: Error general: {ex.Message}");
                overallSuccess = false;
            }
        }

        string finalMessage = overallSuccess
            ? "Alta masiva completada con éxito."
            : "Alta masiva completada con errores.";
        finalMessage += "\nDetalles:\n" + string.Join("\n", messages);

        return Json(new { success = overallSuccess, messages = string.Join("\n", messages), message = finalMessage });
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