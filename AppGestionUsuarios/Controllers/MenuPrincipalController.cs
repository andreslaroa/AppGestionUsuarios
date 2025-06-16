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
            return RedirectToAction("AltaUsuario", "AltaUsuario"); 
        }

        [HttpGet]
        public IActionResult HabilitarDeshabilitarUsuario()
        {
            return RedirectToAction("HabilitarDeshabilitarUsuario", "HabilitarDeshabilitarUsuario"); 
        }

        [HttpGet]
        public IActionResult ModificarUsuario()
        {
            return RedirectToAction("ModificarUsuario", "ModificarUsuario");
        }


        [HttpGet]
        public IActionResult DeleteUser()
        {
            return RedirectToAction("BajaUsuario", "BajaUsuario"); 
        }

        [HttpGet]
        public IActionResult AltaMasiva()
        {
            return RedirectToAction("AltaMasiva", "AltaMasiva"); 
        }
    }
}
