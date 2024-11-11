
using Microsoft.AspNetCore.Mvc;
using System.DirectoryServices.AccountManagement;

namespace LoginApp.Controllers
{
    public class AccountController : Controller
    {
        [HttpGet]
        public IActionResult Login()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Login(string username, string password)
        {
            try
            {
                // Valida las credenciales del usuario en el Directorio Activo
                bool isAuthenticated = ValidateUserCredentials("aytosa.inet", username, password); // Reemplaza "miDominio" con tu dominio real

                if (isAuthenticated)
                {
                    ViewBag.Message = "Credenciales correctas.";
                    return View();
                }
                else
                {
                    ViewBag.Message = "Credenciales incorrectas. Intente nuevamente.";
                    return View();
                }
            }
            catch
            {
                ViewBag.Message = "Ha ocurrido un error al intentar validar las credenciales.";
                return View();
            }
        }

        // Método para validar las credenciales contra el Directorio Activo
        private bool ValidateUserCredentials(string domain, string username, string password)
        {
            try
            {
                using (var context = new PrincipalContext(ContextType.Domain, domain))
                {
                    // Validar las credenciales
                    return context.ValidateCredentials(username, password);
                }
            }
            catch (Exception ex)
            {
                // Manejo de errores (registro opcional)
                Console.WriteLine($"Error al validar las credenciales: {ex.Message}");
                return false;
            }
        }

        public IActionResult LoginSuccess()
        {
            return View();
        }
    }
}
