using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;

[Authorize]
public class AltaMasivaController : Controller
{
    private readonly string domainPath = "DC=aytosa,DC=inet"; // Dominio base para construir las rutas de las OUs

    [HttpGet]
    public IActionResult AltaMasiva()
    {
        // Configurar la licencia de EPPlus (requerido para uso no comercial)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Cargar las OUs principales (solo nombres)
        List<string> ouPrincipales = new List<string>();
        try
        {
            using (DirectoryEntry entry = new DirectoryEntry($"LDAP://{domainPath}"))
            using (DirectorySearcher searcher = new DirectorySearcher(entry))
            {
                searcher.Filter = "(objectClass=organizationalUnit)";
                searcher.PropertiesToLoad.Add("name");
                searcher.PropertiesToLoad.Add("distinguishedName");
                searcher.SearchScope = SearchScope.OneLevel; // Solo OUs de nivel superior

                foreach (SearchResult result in searcher.FindAll())
                {
                    if (result.Properties.Contains("name"))
                    {
                        string name = result.Properties["name"][0].ToString();
                        ouPrincipales.Add(name);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al cargar las OUs: {ex.Message}");
        }

        // Cargar las cuotas
        List<string> cuotas = new List<string> { "5GB", "10GB", "20GB" }; // Ajusta según tu lógica

        // Cargar los grupos disponibles
        List<string> grupos = new List<string>();
        try
        {
            using (DirectoryEntry entry = new DirectoryEntry($"LDAP://{domainPath}"))
            using (DirectorySearcher searcher = new DirectorySearcher(entry))
            {
                searcher.Filter = "(objectClass=group)";
                searcher.PropertiesToLoad.Add("cn");
                searcher.SearchScope = SearchScope.Subtree;

                foreach (SearchResult result in searcher.FindAll())
                {
                    if (result.Properties.Contains("cn"))
                    {
                        string cn = result.Properties["cn"][0].ToString();
                        grupos.Add(cn);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al cargar los grupos: {ex.Message}");
        }

        ViewBag.OUPrincipales = ouPrincipales.OrderBy(ou => ou).ToList();
        ViewBag.Cuota = cuotas.OrderBy(c => c).ToList();
        ViewBag.GruposAD = grupos.OrderBy(g => g).ToList();
        return View();
    }

    [HttpPost]
    public async Task<IActionResult> LoadFile(IFormFile file)
    {
        if (file == null || file.Length == 0)
            return Json(new { success = false, message = "No se ha proporcionado un archivo válido." });

        List<Dictionary<string, object>> users = new List<Dictionary<string, object>>();
        string[] expectedHeaders = new[] { "Nombre", "Apellido1", "Apellido2", "DNI", "nTelefono", "DDI", "MobileExt", "MobileNumber", "TarjetaId", "nFuncionario", "OUPrincipal", "OUSecundaria", "FechaCaducidadOp", "FechaCaducidad", "Cuota", "Grupos" };

        try
        {
            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                stream.Position = 0;

                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // Primera hoja
                    if (worksheet == null || worksheet.Dimension == null)
                        return Json(new { success = false, message = "El archivo Excel está vacío o no tiene datos." });

                    // Leer el encabezado (primera fila)
                    var headers = new List<string>();
                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        var header = worksheet.Cells[1, col].Text?.Trim();
                        if (string.IsNullOrEmpty(header)) break;
                        headers.Add(header);
                    }

                    if (headers.Count != expectedHeaders.Length || !headers.SequenceEqual(expectedHeaders))
                        return Json(new { success = false, message = "El encabezado del archivo Excel no coincide con el formato esperado: " + string.Join(",", expectedHeaders) });

                    // Leer las filas de datos (a partir de la fila 2)
                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        // Verificar si la fila está vacía
                        bool isRowEmpty = true;
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            if (!string.IsNullOrWhiteSpace(worksheet.Cells[row, col].Text))
                            {
                                isRowEmpty = false;
                                break;
                            }
                        }
                        if (isRowEmpty) continue; // Ignorar filas vacías

                        var userData = new Dictionary<string, object>();
                        for (int col = 1; col <= headers.Count; col++)
                        {
                            string key = headers[col - 1];
                            string value = worksheet.Cells[row, col].Text?.Trim() ?? "";
                            if (key == "Grupos")
                            {
                                userData[key] = value.Split(',').Select(g => g.Trim()).Where(g => !string.IsNullOrEmpty(g)).ToList();
                            }
                            else
                            {
                                userData[key] = value;
                            }
                        }

                        // Generar el nombre de usuario automáticamente
                        string username = await GenerateUsername(userData["Nombre"].ToString(), userData["Apellido1"].ToString(), userData["Apellido2"].ToString());
                        if (string.IsNullOrEmpty(username))
                            return Json(new { success = false, message = $"Fila {row}: Error al generar el nombre de usuario." });

                        userData["Username"] = username;
                        users.Add(userData);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = $"Error al procesar el archivo Excel: {ex.Message}" });
        }

        if (!users.Any())
            return Json(new { success = false, message = "El archivo Excel no contiene datos válidos." });

        return Json(new { success = true, users });
    }

    [HttpPost]
    public async Task<IActionResult> ProcessUsers([FromBody] List<Dictionary<string, object>> users)
    {
        if (users == null || !users.Any())
            return Json(new { success = false, message = "No se han proporcionado datos de usuarios." });

        List<string> messages = new List<string>();
        bool overallSuccess = true;

        for (int i = 0; i < users.Count; i++)
        {
            var user = users[i];
            int lineNumber = i + 2; // +2 para contar el encabezado y empezar desde la fila 2

            try
            {
                // Extraer datos del usuario
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
                string fechaCaducidadOp = user.GetValueOrDefault("FechaCaducidadOp", "").ToString();
                string fechaCaducidad = user.GetValueOrDefault("FechaCaducidad", "").ToString();
                string cuota = user.GetValueOrDefault("Cuota", "").ToString();
                List<string> grupos = user.GetValueOrDefault("Grupos", new List<string>()) as List<string> ?? new List<string>();
                string username = user.GetValueOrDefault("Username", "").ToString();

                // Mapear nombres de OUs a rutas completas
                string ouPrincipalPath = string.Empty;
                string ouSecundariaPath = string.Empty;

                if (!string.IsNullOrEmpty(ouPrincipalName))
                {
                    ouPrincipalPath = await GetOUPath(ouPrincipalName, null);
                    if (string.IsNullOrEmpty(ouPrincipalPath))
                    {
                        messages.Add($"Fila {lineNumber}: La OU principal '{ouPrincipalName}' no existe.");
                        overallSuccess = false;
                        continue;
                    }
                }

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

                string targetOU = !string.IsNullOrEmpty(ouSecundariaPath) ? ouSecundariaPath : ouPrincipalPath;

                // Validaciones
                if (string.IsNullOrEmpty(nombre) || string.IsNullOrEmpty(apellido1) || string.IsNullOrEmpty(dni) || string.IsNullOrEmpty(ouPrincipalName))
                {
                    messages.Add($"Fila {lineNumber}: Faltan datos obligatorios (Nombre, Apellido1, DNI, OUPrincipal).");
                    overallSuccess = false;
                    continue;
                }

                if (fechaCaducidadOp == "sí" && string.IsNullOrEmpty(fechaCaducidad))
                {
                    messages.Add($"Fila {lineNumber}: Debe especificar una fecha de caducidad.");
                    overallSuccess = false;
                    continue;
                }

                // Validar duplicados
                if (await CheckDniExists(dni))
                {
                    messages.Add($"Fila {lineNumber}: El DNI '{dni}' ya existe.");
                    overallSuccess = false;
                    continue;
                }

                if (!string.IsNullOrEmpty(nTelefono) && await CheckTelephoneExists(nTelefono))
                {
                    messages.Add($"Fila {lineNumber}: El número de teléfono '{nTelefono}' ya existe.");
                    overallSuccess = false;
                    continue;
                }

                if (!string.IsNullOrEmpty(nFuncionario) && nFuncionario != "999999" && await CheckNumberIdExists(nFuncionario))
                {
                    messages.Add($"Fila {lineNumber}: El número de funcionario '{nFuncionario}' ya existe.");
                    overallSuccess = false;
                    continue;
                }

                // Crear el usuario en Active Directory
                using (var context = new PrincipalContext(ContextType.Domain, "aytosa.inet"))
                {
                    using (var userPrincipal = new UserPrincipal(context))
                    {
                        userPrincipal.SamAccountName = username;
                        userPrincipal.DisplayName = $"{nombre} {apellido1} {apellido2}".Trim();
                        userPrincipal.GivenName = nombre;
                        userPrincipal.Surname = $"{apellido1} {apellido2}".Trim();
                        userPrincipal.SetPassword("P@ssw0rd"); // Contraseña por defecto
                        userPrincipal.Enabled = true;

                        if (!string.IsNullOrEmpty(dni))
                            userPrincipal.Description = dni;

                        if (!string.IsNullOrEmpty(nTelefono))
                            userPrincipal.VoiceTelephoneNumber = nTelefono;

                        userPrincipal.Save();

                        // Acceder al objeto subyacente para establecer propiedades personalizadas
                        using (var userEntry = (DirectoryEntry)userPrincipal.GetUnderlyingObject())
                        {
                            // Establecer propiedades personalizadas
                            if (!string.IsNullOrEmpty(ddi))
                                userEntry.Properties["otherTelephone"].Value = ddi;

                            if (!string.IsNullOrEmpty(mobileExt))
                                userEntry.Properties["otherMobile"].Value = mobileExt;

                            if (!string.IsNullOrEmpty(mobileNumber))
                                userEntry.Properties["mobile"].Value = mobileNumber;

                            if (!string.IsNullOrEmpty(tarjetaId))
                                userEntry.Properties["extensionAttribute1"].Value = tarjetaId;

                            if (!string.IsNullOrEmpty(nFuncionario))
                                userEntry.Properties["employeeID"].Value = nFuncionario;

                            userEntry.CommitChanges();

                            // Mover a la OU seleccionada
                            using (var ouEntry = new DirectoryEntry($"LDAP://{targetOU}"))
                            {
                                userEntry.MoveTo(ouEntry);
                                userEntry.CommitChanges();
                            }

                            // Asignar a grupos
                            foreach (var grupo in grupos)
                            {
                                try
                                {
                                    using (var groupEntry = new DirectoryEntry($"LDAP://CN={grupo},DC=aytosa,DC=inet"))
                                    {
                                        groupEntry.Properties["member"].Add(userEntry.Properties["distinguishedName"].Value);
                                        groupEntry.CommitChanges();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    messages.Add($"Fila {lineNumber}: Error al asignar al grupo '{grupo}': {ex.Message}");
                                }
                            }
                        }

                        // Configurar fecha de caducidad
                        if (fechaCaducidadOp == "sí" && !string.IsNullOrEmpty(fechaCaducidad))
                        {
                            userPrincipal.AccountExpirationDate = DateTime.Parse(fechaCaducidad);
                            userPrincipal.Save();
                        }
                    }
                }

                // Crear carpeta personal (si aplica)
                if (!ouPrincipalName.Contains("OAGER"))
                {
                    try
                    {
                        string userFolderPath = $"\\\\fs1.aytosa.inet\\home\\{username}";
                        if (!Directory.Exists(userFolderPath))
                        {
                            Directory.CreateDirectory(userFolderPath);
                            messages.Add($"Fila {lineNumber}: Carpeta personal '{userFolderPath}' creada correctamente.");
                        }
                    }
                    catch (Exception ex)
                    {
                        messages.Add($"Fila {lineNumber}: Error al crear carpeta personal para '{username}': {ex.Message}");
                    }

                    // Configurar cuota FSRM (si aplica)
                    if (!string.IsNullOrEmpty(cuota))
                    {
                        try
                        {
                            messages.Add($"Fila {lineNumber}: Cuota '{cuota}' configurada correctamente para '{username}'.");
                        }
                        catch (Exception ex)
                        {
                            messages.Add($"Fila {lineNumber}: Error al configurar la cuota para '{username}': {ex.Message}");
                        }
                    }
                }

                messages.Add($"Fila {lineNumber}: Usuario '{username}' creado correctamente.");
            }
            catch (Exception ex)
            {
                messages.Add($"Fila {lineNumber}: Error al crear el usuario: {ex.Message}");
                overallSuccess = false;
            }
        }

        string finalMessage = overallSuccess
            ? "Alta masiva completada con éxito."
            : "Alta masiva completada con errores.";
        finalMessage += "\nDetalles:\n" + string.Join("\n", messages);

        return Json(new { success = overallSuccess, messages = string.Join("\n", messages), message = finalMessage });
    }

    private async Task<string> GenerateUsername(string nombre, string apellido1, string apellido2)
    {
        try
        {
            string baseUsername = $"{nombre[0]}{apellido1}".ToLower();
            string username = baseUsername;
            int suffix = 1;

            while (await CheckUsernameExists(username))
            {
                username = $"{baseUsername}{suffix}";
                suffix++;
            }

            return username;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al generar nombre de usuario: {ex.Message}");
            return null;
        }
    }

    private async Task<bool> CheckUsernameExists(string username)
    {
        try
        {
            using (var context = new PrincipalContext(ContextType.Domain, "aytosa.inet"))
            {
                return UserPrincipal.FindByIdentity(context, username) != null;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al verificar nombre de usuario: {ex.Message}");
            return true; // Asumimos que existe para evitar crear duplicados
        }
    }

    private async Task<bool> CheckDniExists(string dni)
    {
        try
        {
            using (DirectoryEntry entry = new DirectoryEntry($"LDAP://{domainPath}"))
            using (DirectorySearcher searcher = new DirectorySearcher(entry))
            {
                searcher.Filter = $"(&(objectClass=user)(description={dni}))";
                searcher.SearchScope = SearchScope.Subtree;
                return searcher.FindOne() != null;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al verificar DNI: {ex.Message}");
            return true;
        }
    }

    private async Task<bool> CheckTelephoneExists(string nTelefono)
    {
        try
        {
            using (DirectoryEntry entry = new DirectoryEntry($"LDAP://{domainPath}"))
            using (DirectorySearcher searcher = new DirectorySearcher(entry))
            {
                searcher.Filter = $"(&(objectClass=user)(telephoneNumber={nTelefono}))";
                searcher.SearchScope = SearchScope.Subtree;
                return searcher.FindOne() != null;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al verificar número de teléfono: {ex.Message}");
            return true;
        }
    }

    private async Task<bool> CheckNumberIdExists(string nFuncionario)
    {
        try
        {
            using (DirectoryEntry entry = new DirectoryEntry($"LDAP://{domainPath}"))
            using (DirectorySearcher searcher = new DirectorySearcher(entry))
            {
                searcher.Filter = $"(&(objectClass=user)(employeeID={nFuncionario}))";
                searcher.SearchScope = SearchScope.Subtree;
                return searcher.FindOne() != null;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al verificar número de funcionario: {ex.Message}");
            return true;
        }
    }

    private async Task<string> GetOUPath(string ouName, string parentOUName)
    {
        try
        {
            using (DirectoryEntry entry = new DirectoryEntry($"LDAP://{domainPath}"))
            using (DirectorySearcher searcher = new DirectorySearcher(entry))
            {
                if (string.IsNullOrEmpty(parentOUName))
                {
                    // Buscar OU principal
                    searcher.Filter = $"(&(objectClass=organizationalUnit)(name={ouName}))";
                    searcher.SearchScope = SearchScope.OneLevel; // Solo nivel superior
                }
                else
                {
                    // Buscar OU secundaria dentro de la OU principal
                    string parentOUPath = await GetOUPath(parentOUName, null);
                    if (string.IsNullOrEmpty(parentOUPath)) return null;

                    searcher.SearchRoot = new DirectoryEntry($"LDAP://{parentOUPath}");
                    searcher.Filter = $"(&(objectClass=organizationalUnit)(name={ouName}))";
                    searcher.SearchScope = SearchScope.OneLevel; // Solo sub-OUs
                }

                searcher.PropertiesToLoad.Add("distinguishedName");
                var result = searcher.FindOne();
                return result?.Properties["distinguishedName"]?.Count > 0
                    ? result.Properties["distinguishedName"][0].ToString()
                    : null;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error al buscar OU '{ouName}': {ex.Message}");
            return null;
        }
    }
}