using Microsoft.AspNetCore.Mvc;
using System;
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

        string hola = "hola mundo";

        // Redirección a una página de éxito si las credenciales son correctas
        public IActionResult LoginSuccess()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Login(string username, string password)
        {
            
            string domainName = "aytosa.inet"; 

            // Validar las credenciales ingresadas por el usuario contra el dominio especificado
            bool isAuthenticated = ValidateUserCredentials(domainName, username, password);

            if (isAuthenticated)
            {
                ViewBag.Message = "Credenciales correctas.";
                return RedirectToAction("LoginSuccess");
            }
            else
            {
                ViewBag.Message = "Credenciales incorrectas. Intente nuevamente.";
                return View();
            }
        }

        // Método para validar las credenciales del usuario contra el dominio especificado
        private bool ValidateUserCredentials(string domain, string username, string password)
        {
            try
            {
                // Crear un contexto del dominio usando solo el nombre del dominio proporcionado
                using (var context = new PrincipalContext(ContextType.Domain, domain))
                {
                    // Validar las credenciales del usuario
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
    }
}
