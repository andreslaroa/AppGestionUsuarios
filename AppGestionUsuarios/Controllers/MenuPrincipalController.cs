using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace AppGestionUsuarios.Controllers
{
    [Authorize]
    public class MenuPrincipalController : Controller
    {
        [HttpGet]
        public IActionResult Index()
        {
            // Puedes agregar lógica aquí si necesitas pasar datos a la vista.
            ViewBag.Message = "Bienvenido al panel de control.";
            return View();
        }

        [HttpGet]
        public IActionResult AltaUsuario()
        {
            return RedirectToAction("AltaUsuario", "AltaUsuario"); // Redirige al método que maneja la creación
        }

        [HttpGet]
        public IActionResult EditUser()
        {
            return RedirectToAction("EditUser", "UserManagement"); // Asegúrate de tener este método en UserManagementController
        }

        [HttpGet]
        public IActionResult DeleteUser()
        {
            return RedirectToAction("BajaUsuario", "BajaUsuario"); // Asegúrate de tener este método en UserManagementController
        }

        [HttpGet]
        public IActionResult AltaMasiva()
        {
            return RedirectToAction("AltaMasiva", "AltaMasiva"); // Asegúrate de tener este método en UserManagementController
        }
    }
}
