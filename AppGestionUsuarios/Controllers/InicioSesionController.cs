using Microsoft.AspNetCore.Mvc;
using System.DirectoryServices.AccountManagement;
using Microsoft.AspNetCore.Authentication;
using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.AspNetCore.Http;

namespace AppGestionUsuarios.Controllers
{
    public class InicioSesionController : Controller
    {
        private readonly IConfiguration _configuration;
        private readonly IDataProtector _protector;

        public InicioSesionController(
            IConfiguration configuration,
            IDataProtectionProvider dataProtectionProvider)
        {
            _configuration = configuration;
            // Creamos un protector con un propósito único
            _protector = dataProtectionProvider.CreateProtector("CredencialesProtector");
        }

        [HttpGet]
        public IActionResult Login()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> LoginAsync(string username, string password)
        {
            // 1) Validar credenciales contra AD y grupo
            var validationResult = ValidateUserCredentialsAndGroup(username, password);
            if (validationResult != "correcto")
            {
                ViewBag.Message = validationResult;
                return View();
            }

            // 2) Guardar en Session (cifrar la contraseña)
            HttpContext.Session.SetString("adminUser", username);
            HttpContext.Session.SetString("adminPassword", _protector.Protect(password));

            // 3) Crear la cookie de autenticación
            var claims = new List<Claim>
            {
                new Claim(ClaimTypes.Name, username),
                new Claim(ClaimTypes.Role, "User")
            };
            var claimsIdentity = new ClaimsIdentity(claims, "CookieAuth");
            await HttpContext.SignInAsync("CookieAuth", new ClaimsPrincipal(claimsIdentity));

            return RedirectToAction("index", "MenuPrincipal");
        }

        [Authorize]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Logout()
        {
            await HttpContext.SignOutAsync("CookieAuth");
            return RedirectToAction("Login");
        }

        public IActionResult MenuPrincipal()
        {
            return View();
        }

        // Ya no cambia: sigue validando contra AD
        private string ValidateUserCredentialsAndGroup(string username, string password)
        {
            var domain    = _configuration["ActiveDirectory:DomainName"];
            var groupName = _configuration["ActiveDirectory:AdminGroup"];

            try
            {
                using var context = new PrincipalContext(ContextType.Domain, domain);
                if (!context.ValidateCredentials(username, password))
                    return "Credenciales inválidas.";

                using var userPrincipal = UserPrincipal.FindByIdentity(context, username);
                if (userPrincipal == null)
                    return "Usuario no encontrado en el dominio";

                using var groupPrincipal = GroupPrincipal.FindByIdentity(context, groupName);
                if (groupPrincipal == null)
                    return "Error buscando el grupo de permisos";

                if (!userPrincipal.IsMemberOf(groupPrincipal))
                    return "El usuario no tiene los permisos necesarios";

                return "correcto";
            }
            catch (Exception ex)
            {
                return $"Error al validar credenciales o grupo: {ex.Message}";
            }
        }
    }
}
