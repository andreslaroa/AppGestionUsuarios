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
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.Extensions.Logging.Abstractions;
using System.Collections.ObjectModel;
using System.Management.Automation;


[Authorize]
public class GestionUsuariosController : Controller
{
    private readonly OUService _ouService;
    private readonly ILogger<GestionUsuariosController> _logger;

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
        public string FechaCaducidadOp {  get; set; }
        public DateTime FechaCaducidad { get; set; }
        public string Cuota { get; set; }
    }

    public class userInputModel
    {
        public string Nombre { set; get; }
        public string Apellido1 { set; get; }
        public string Apellido2 { set; get; }
    }



        public GestionUsuariosController()
    {
        string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Resources", "ArchivoDePruebasOU.xlsx");
        _ouService = new OUService(filePath);
    }  
    

    [HttpGet]
    public IActionResult AltaUsuario()
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



    [HttpPost]
    public IActionResult CreateUser([FromBody] UserModelAltaUsuario user)
    {
        // Validar si los datos se recibieron correctamente
        if (user == null)
        {
            return Json(new { success = false, message = "No se recibieron datos válidos." });
        }

        // Validar los campos obligatorios
        if (string.IsNullOrEmpty(user.Nombre) || string.IsNullOrEmpty(user.Apellido1) || string.IsNullOrEmpty(user.NTelefono) || string.IsNullOrEmpty(user.Username) ||
            string.IsNullOrEmpty(user.OUPrincipal) || string.IsNullOrEmpty(user.OUSecundaria) || string.IsNullOrEmpty(user.Departamento) || string.IsNullOrEmpty(user.FechaCaducidadOp))           
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
                    newUser.Properties["givenName"].Value = user.Nombre;
                    newUser.Properties["sn"].Value = user.Apellido1 + user.Apellido2;
                    newUser.Properties["sAMAccountName"].Value = user.Username; // Nombre de usuario corto
                    newUser.Properties["userPrincipalName"].Value = $"{user.Username}@aytosa.inet"; // Dominio
                    newUser.Properties["displayName"].Value = displayName; // Nombre completo
                    newUser.Properties["description"].Value = user.NFuncionario; // Descripción
                    newUser.Properties["telephoneNumber"].Value = user.NTelefono; //Extenión de teléfono
                    newUser.Properties["physicalDeliveryOfficeName"].Value = user.Departamento; //Departamento del usuario

                    //Comando para meter el nombre del correo en el directorio activo
                    //newUser.Properties["email"].Value = $"{user.Username}@aytosalamanca.es";

                    
                    if (user.FechaCaducidadOp == "si")
                    {
                        
                        // Validar que la fecha no sea anterior al momento actual
                        if (user.FechaCaducidad <= DateTime.Now)
                        {
                            return Json(new { success = false, message = "La fecha de caducidad debe ser una fecha futura." });
                        }

                        try
                        {
                            // Convertir la fecha a Windows File Time (Int64)
                            long accountExpires = user.FechaCaducidad.ToFileTime();

                            newUser.Properties["accountExpires"].Value = accountExpires.ToString();
                        }
                        catch (ArgumentOutOfRangeException ex)
                        {
                            return Json(new { success = false, message = $"Error al convertir la fecha. {ex.Message}" });
                        }
                    }
                    //Si decimos que no queremos fecha de caducidad, la creación de usuario por defecto pone a nunca la fecha de expiración

                          
                    //Cuando se realicen las pruebas reales descomentar esta zona de abajo que es la que crea el directorio de usuario en ribera y le asigna la cuota

                    //int cuotaMB = ObtenerCuotaEnMB(user.Cuota);

                    //try
                    //{
                        

                    //    var (success, message) = ConfigurarDirectorioYCuotaRemoto(user.Username, cuotaMB.ToString());

                    //    if (!success)
                    //    {
                    //        // Devolver el error desde la configuración de la cuota
                    //        return Json(new { success = false, message = $"Error al configurar el directorio: {message}" });
                    //    }
                    //}
                    //catch (Exception ex)
                    //{
                    //    return Json(new { success = false, message = $"Error al crear el directorio propio del usuario: {ex.Message}" });
                    //}

                    newUser.CommitChanges();

                    newUser.Invoke("SetPassword", new object[] { "Temporal2024" });

                    // Activar la cuenta
                    newUser.Properties["userAccountControl"].Value = 0x200;

                    // Forzar el cambio de contraseña al primer inicio de sesión
                    newUser.Properties["pwdLastSet"].Value = 0;

                    newUser.CommitChanges();
                }
                catch (Exception ex)
                {
                    return Json(new { success = false, message = $"Error al crear el usuario: {ex.Message}" });
                }
                finally
                {
                    if (newUser != null)
                    {
                        newUser.Dispose();
                    }
                }

                
                return Json(new { success = true, message = "Usuario creado exitosamente con fecha de expiración: ." });
                
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}\nStackTrace: {ex.StackTrace}");
            return Json(new { success = false, message = $"Error al crear el usuario: {ex.Message}" });
        }
    }

    private (bool success, string message) ConfigurarDirectorioYCuotaRemoto( string username, string quota)
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
