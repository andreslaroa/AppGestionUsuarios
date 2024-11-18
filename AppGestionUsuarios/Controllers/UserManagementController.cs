using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.IO;

public class UserManagementController : Controller
{
    private readonly OUService _ouService;

    public UserManagementController()
    {
        string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Resources", "ArchivoDePruebasOU.xlsx");
        _ouService = new OUService(filePath);
    }

    [HttpGet]
    public IActionResult LoginSuccess()
    {
        var ouPrincipales = _ouService.GetOUPrincipales();
        ViewBag.OUPrincipales = ouPrincipales;
        var portalEmpleado = _ouService.GetPortalEmpleado();
        ViewBag.portalEmpleado = portalEmpleado;
        return View();
    }

    [HttpPost]
    public IActionResult GetOUSecundarias([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData != null && requestData.ContainsKey("ouPrincipal"))
        {
            string ouPrincipal = requestData["ouPrincipal"];
            var ouSecundarias = _ouService.GetOUSecundarias(ouPrincipal);
            return Json(ouSecundarias);
        }

        return Json(new List<string>());
    }

    [HttpPost]
    public IActionResult GetDepartamentos([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData != null && requestData.ContainsKey("ouPrincipal"))
        {
            string ouPrincipal = requestData["ouPrincipal"];
            var departamentos = _ouService.GetDepartamentos(ouPrincipal);
            return Json(departamentos);
        }

        return Json(new List<string>());
    }

    // Nuevo método para obtener el lugar de envío
    [HttpPost]
    public IActionResult GetLugarEnvio([FromBody] Dictionary<string, string> requestData)
    {
        if (requestData != null && requestData.ContainsKey("departamento"))
        {
            string departamento = requestData["departamento"];
            var lugaresEnvio = _ouService.GetLugarEnvio(departamento);
            return Json(lugaresEnvio);
        }

        return Json(new List<string>());
    }
}
