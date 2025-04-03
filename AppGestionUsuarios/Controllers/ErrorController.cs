using Microsoft.AspNetCore.Diagnostics;
using Microsoft.AspNetCore.Mvc;

public class ErrorController : Controller
{
    [Route("/Error")]
    public IActionResult Index()
    {
        var exceptionFeature = HttpContext.Features.Get<IExceptionHandlerPathFeature>();
        if (exceptionFeature != null)
        {
            string errorMessage = exceptionFeature.Error.Message;
            ViewBag.ErrorMessage = errorMessage; // Pasar el mensaje a la vista
        }
        return View();
    }
}