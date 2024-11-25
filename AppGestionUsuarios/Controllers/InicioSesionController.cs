using Microsoft.AspNetCore.Mvc;
using System;
using System.DirectoryServices.AccountManagement;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication;
using System.Security.Claims;

namespace AppGestionUsuarios.Controllers
{
    public class InicioSesionController : Controller
    {
        [HttpGet]
        public IActionResult Login()
        {
            return View();
        }


        // Redirección a una página de éxito si las credenciales son correctas
        public IActionResult MenuPrincipal()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> LoginAsync(string username, string password)
        {
            
            string domainName = "aytosa.inet";
            string groupName = "GADM_SISTEMAS";

            // Validar las credenciales ingresadas por el usuario contra el dominio especificado
            string isAuthenticated = ValidateUserCredentialsAndGroup(domainName, username, password, groupName);

            if (isAuthenticated == "correcto")
            {
                // Crear las claims del usuario autenticado
                var claims = new List<Claim>
            {
                new Claim(ClaimTypes.Name, username),
                new Claim(ClaimTypes.Role, "User") // Rol de ejemplo
            };

                var claimsIdentity = new ClaimsIdentity(claims, "CookieAuth");

                // En esta línea le indicamos al hilo que ejecuta el método que debe permanecer abierto hasta que terminen las tareas del usuario
                //La explicación sencilla consiste en que guardara el código propio del usuario para identificarlo en todo momento
                await HttpContext.SignInAsync("CookieAuth", new ClaimsPrincipal(claimsIdentity)); 

                ViewBag.Message = "Credenciales correctas.";
                return RedirectToAction("index", "MenuPrincipal");
            }
            else
            {
                ViewBag.Message = isAuthenticated;
                return View();
            }
        }

        // Método para validar las credenciales del usuario contra el dominio especificado
        private string ValidateUserCredentialsAndGroup(string domain, string username, string password, string groupName)
        {
            try
            {
                // Crear un contexto del dominio usando solo el nombre del dominio proporcionado
                using (var context = new PrincipalContext(ContextType.Domain, domain))
                {
                    // Validar las credenciales del usuario
                    if (!context.ValidateCredentials(username, password))
                    {
                        return "Credenciales inválidas.";
                    }

                    // Buscar al usuario en el dominio
                    using (var userPrincipal = UserPrincipal.FindByIdentity(context, username))
                    {
                        if (userPrincipal == null)
                        {
                            return "Usuario no encontrado en el dominio";
                        }

                        // Verificar si el usuario pertenece al grupo especificado
                        using (var groupPrincipal = GroupPrincipal.FindByIdentity(context, groupName))
                        {
                            if (groupPrincipal == null)
                            {
                                return $"Error buscando el grupo de permisos";
                            }

                            if (!userPrincipal.IsMemberOf(groupPrincipal))
                            {
                                return $"El usuario no tiene los permisos necesarios";
                            }

                            // El usuario pertenece al grupo
                            return "correcto";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Manejo de errores
                return $"Error al validar las credenciales o grupo: {ex.Message}";
            }
        }

    }
}
