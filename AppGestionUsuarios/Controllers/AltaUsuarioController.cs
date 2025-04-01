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


/*En esta clase encontramos todos los métodos que son concretos del alta de usuario*/
/*En el caso de métodos que puedan usar otros menús, se almacenan en el apartado de gestión de usuarios*/

[Authorize]
public class AltaUsuarioController : Controller
{
    private readonly OUService _ouService;

    public AltaUsuarioController()
    {
        string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Resources", "ArchivoDePruebasOU.xlsx");
        _ouService = new OUService(filePath);
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
        public string FechaCaducidadOp { get; set; }
        public DateTime FechaCaducidad { get; set; }
        public string Cuota { get; set; }
        public List<string> Grupos { get; set; }
    }

   
    //Función para recibir la petición get
    [HttpGet]
    public IActionResult AltaUsuario()
    {
        ViewBag.OUPrincipales = _ouService.GetOUPrincipales();
        ViewBag.portalEmpleado = _ouService.GetPortalEmpleado();
        ViewBag.cuota = _ouService.GetCuota();

        // Grupos de AD
        try
        {
            using (var entry = new DirectoryEntry("LDAP://DC=aytosa,DC=inet"))
            using (var searcher = new DirectorySearcher(entry))
            {
                searcher.Filter = "(objectClass=group)";
                searcher.PropertiesToLoad.Add("cn");
                searcher.SearchScope = SearchScope.Subtree;

                var grupos = new List<string>();
                foreach (SearchResult result in searcher.FindAll())
                {
                    if (result.Properties.Contains("cn"))
                        grupos.Add(result.Properties["cn"][0].ToString());
                }

                ViewBag.GruposAD = grupos.OrderBy(g => g).ToList();
            }
        }
        catch
        {
            ViewBag.GruposAD = new List<string>();
        }

        return View("AltaUsuario"); // Asegúrate de que tu vista esté en /Views/AltaUsuario/AltaUsuario.cshtml
    }


    //Función propia para crear el usuario
    [HttpPost]
    public IActionResult CreateUser([FromBody] UserModelAltaUsuario user)
    {
        // Validar si los datos se recibieron correctamente
        if (user == null)
        {
            return Json(new { success = false, message = "No se recibieron datos válidos." });
        }

        // Validar los campos obligatorios
        if (string.IsNullOrEmpty(user.Nombre) || string.IsNullOrEmpty(user.Apellido1) ||
            string.IsNullOrEmpty(user.NTelefono) || string.IsNullOrEmpty(user.Username) ||
            string.IsNullOrEmpty(user.OUPrincipal) || string.IsNullOrEmpty(user.OUSecundaria) ||
            string.IsNullOrEmpty(user.Departamento) || string.IsNullOrEmpty(user.FechaCaducidadOp))
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
                    newUser.Properties["sn"].Value = user.Apellido1 + " " + user.Apellido2;
                    newUser.Properties["sAMAccountName"].Value = user.Username;
                    newUser.Properties["userPrincipalName"].Value = $"{user.Username}@aytosa.inet";
                    newUser.Properties["displayName"].Value = displayName;
                    newUser.Properties["description"].Value = user.NFuncionario;
                    newUser.Properties["telephoneNumber"].Value = user.NTelefono;
                    newUser.Properties["physicalDeliveryOfficeName"].Value = user.Departamento;

                    if (user.FechaCaducidadOp == "si")
                    {
                        if (user.FechaCaducidad <= DateTime.Now)
                        {
                            return Json(new { success = false, message = "La fecha de caducidad debe ser una fecha futura." });
                        }

                        try
                        {
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

                    // Configurar contraseña y activar cuenta
                    newUser.Invoke("SetPassword", new object[] { "Temporal2024" });
                    newUser.Properties["userAccountControl"].Value = 0x200;
                    newUser.Properties["pwdLastSet"].Value = 0;

                    newUser.CommitChanges();


                    // Añadir al usuario a los grupos seleccionados
                    if (user.Grupos != null && user.Grupos.Any())
                    {
                        foreach (string grupo in user.Grupos)
                        {
                            DirectoryEntry groupEntry = FindGroupByName(grupo);
                            if (groupEntry != null)
                            {
                                try
                                {
                                    // Agregar el usuario al grupo
                                    groupEntry.Invoke("Add", new object[] { newUser.Path });
                                    groupEntry.CommitChanges();
                                }
                                catch (Exception ex)
                                {

                                }
                                finally
                                {
                                    groupEntry.Dispose();
                                }
                            }
                            else
                            {
                                return Json(new { success = false, message = $"Grupo {grupo} no encontrado en el dominio." });
                            }
                        }
                    }

                    newUser.CommitChanges();

                    //Falta la creación del correo electrónico

                }
                catch (Exception ex)
                {
                    return Json(new { success = false, message = $"Error al crear el usuario: {ex.Message}" });
                }
                finally
                {
                    newUser?.Dispose();
                }

                return Json(new { success = true, message = "Usuario creado exitosamente y añadido a los grupos seleccionados." });
            }
        }
        catch (Exception ex)
        {
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
