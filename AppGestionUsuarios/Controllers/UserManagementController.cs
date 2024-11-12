using Microsoft.AspNetCore.Mvc;

namespace AppGestionUsuarios.Controllers
{

    public class UserManagementController : Controller
    {

        [HttpGet]
        public IActionResult LoginSuccess()
        {
            // Lógica para inicializar datos o manejar el formulario en la vista
            return View();
        }



    }
}
