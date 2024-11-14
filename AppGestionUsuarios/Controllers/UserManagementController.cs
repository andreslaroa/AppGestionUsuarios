using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;

namespace AppGestionUsuarios.Controllers
{
    public class UserManagementController : Controller
    {
        private readonly OUService _ouService;

        public UserManagementController()
        {
            // Ruta al archivo Excel
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Resources", "ArchivoDePruebasOU.xlsx");
            _ouService = new OUService(filePath);
        }

        [HttpGet]
        public IActionResult LoginSuccess()
        {
            // Obtener OUs principales y secundarias
            var ouPrincipales = _ouService.GetOUPrincipales();
            var ouSecundarias = _ouService.GetOUSecundarias();

            // Pasar los datos a la vista usando ViewBag
            ViewBag.OUPrincipales = ouPrincipales;
            ViewBag.OUSecundarias = ouSecundarias;

            return View();
        }
    }
}
