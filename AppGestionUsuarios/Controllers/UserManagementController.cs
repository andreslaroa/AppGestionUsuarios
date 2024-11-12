using Microsoft.AspNetCore.Mvc;

namespace AppGestionUsuarios.Controllers
{
    public class UserManagementController : Controller
    {
        [HttpGet]
        public IActionResult LoginSuccess()
        {
            // No cargar datos ni utilizar OUService; simplemente devolver la vista
            return View();
        }
    }
}
