using Microsoft.AspNetCore.Mvc;
using System.Management.Automation;

namespace AppGestionUsuarios.Controllers
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
            if (CheckUserExistsInAD(username))
            {
                ViewBag.Message = "El usuario existe en el Directorio Activo.";
            }
            else
            {
                ViewBag.Message = "El usuario no existe en el Directorio Activo.";
            }

            return View();
        }

        private bool CheckUserExistsInAD(string username)
        {
            try
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.AddCommand("Get-ADUser")
                      .AddParameter("Identity", username);

                    var results = ps.Invoke();

                    // Si se encuentra un resultado, el usuario existe
                    return results.Count > 0;
                }
            }
            catch
            {
                // Si hay algún error (p.ej., falta de permisos o el comando falla), manejamos esto devolviendo falso
                return false;
            }
        }
    }
}

