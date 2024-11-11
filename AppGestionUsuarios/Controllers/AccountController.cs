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

        //Las redirecciones de estos métodos de IActionResult se consiguen gracias al nombre del método que coincide con la vista de la carpeta Account
        public IActionResult LoginSuccess()
        {
            return View();
        }


        [HttpPost]
        public IActionResult Login(string username, string password)
        {
            // Nombre del dominio de tu organización
            string domainName = "aytosa.inet"; // Cambia esto por tu dominio real
            string domainController = "leonardo"; // Cambia esto por el nombre de tu controlador de dominio si es necesario

            // Validar las credenciales ingresadas por el usuario contra el dominio especificado
            bool isAuthenticated = ValidateUserCredentials(domainController, domainController, username, password);

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

        // Método para validar las credenciales del usuario contra el dominio especificado usando credenciales explícitas
        private bool ValidateUserCredentials(string domainController, string domain, string username, string password)
        {
            try
            {
                // Credenciales explícitas para el contexto (si es necesario)
                string serviceAccountUser = "dominio\\usuarioServicio"; // Cambia esto por tu usuario de servicio
                string serviceAccountPassword = "contraseña"; // Cambia esto por la contraseña del usuario de servicio

                using (var context = new PrincipalContext(ContextType.Domain, domainController, domain, ContextOptions.Negotiate, serviceAccountUser, serviceAccountPassword))
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
}  puedes hacerme este código pero quitando el controlador de dominio, y utilizando unicamente el nombre de dominio?
