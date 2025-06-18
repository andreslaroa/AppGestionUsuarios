using Microsoft.AspNetCore.Mvc;
using System.DirectoryServices.AccountManagement;
using Microsoft.AspNetCore.Authentication;
using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;

namespace AppGestionUsuarios.Controllers
{
    public class InicioSesionController : Controller
    {

        private readonly IConfiguration _configuration;

        public InicioSesionController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

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
            // Validación de credenciales y grupo leyendo la configuración
            var validationResult = ValidateUserCredentialsAndGroup(username, password);

            if (validationResult == "correcto")
            {
                var claims = new List<Claim>
                {
                    new Claim(ClaimTypes.Name, username),
                    new Claim(ClaimTypes.Role, "User")
                };

                var claimsIdentity = new ClaimsIdentity(claims, "CookieAuth");

                await HttpContext.SignInAsync("CookieAuth", new ClaimsPrincipal(claimsIdentity));

                return RedirectToAction("Index", "MenuPrincipal");
            }

            ViewBag.Message = validationResult;
            return View();
        }

        [Authorize]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Logout()
        {
            // Destruye la cookie de autenticación
            await HttpContext.SignOutAsync("CookieAuth");
            // Redirige al formulario de login
            return RedirectToAction("Login", "InicioSesion");
        }

        // Método para validar las credenciales del usuario contra el dominio especificado
        private string ValidateUserCredentialsAndGroup(string username, string password)
        {
            var domain = _configuration["ActiveDirectory:DomainName"];
            var groupName = _configuration["ActiveDirectory:AdminGroup"];

            try
            {
                using (var context = new PrincipalContext(ContextType.Domain, domain))
                {
                    if (!context.ValidateCredentials(username, password))
                    {
                        return "Credenciales inválidas.";
                    }

                    using (var userPrincipal = UserPrincipal.FindByIdentity(context, username))
                    {
                        if (userPrincipal == null)
                        {
                            return "Usuario no encontrado en el dominio";
                        }

                        using (var groupPrincipal = GroupPrincipal.FindByIdentity(context, groupName))
                        {
                            if (groupPrincipal == null)
                            {
                                return "Error buscando el grupo de permisos";
                            }

                            if (!userPrincipal.IsMemberOf(groupPrincipal))
                            {
                                return "El usuario no tiene los permisos necesarios";
                            }

                            return "correcto";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return $"Error al validar las credenciales o grupo: {ex.Message}";
            }
        }

    }
}
