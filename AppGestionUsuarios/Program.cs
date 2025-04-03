using Microsoft.AspNetCore.Server.Kestrel.Core;
using Microsoft.EntityFrameworkCore;

internal class Program
{
    private static void Main(string[] args)
    {
        var builder = WebApplication.CreateBuilder(args);

        // Configurar Kestrel para escuchar en puertos espec�ficos
        builder.WebHost.ConfigureKestrel(options =>
        {
            options.ListenAnyIP(8081, listenOptions =>
            {
                listenOptions.Protocols = HttpProtocols.Http1;
            });
            options.ListenAnyIP(8082, listenOptions =>
            {
                listenOptions.Protocols = HttpProtocols.Http1;
                listenOptions.UseHttps(); // Habilitar HTTPS en el puerto 8082
            });
        });

        // Agregar servicios de autenticaci�n por cookies
        builder.Services.AddAuthentication("CookieAuth")
            .AddCookie("CookieAuth", options =>
            {
                options.LoginPath = "/InicioSesion/Login"; // Ruta a la p�gina de inicio de sesi�n
                options.AccessDeniedPath = "/InicioSesion/Error"; // Ruta en caso de acceso denegado
            });

        builder.Services.AddAuthorization(); // Usar autenticaci�n en la aplicaci�n
        builder.Services.AddControllersWithViews();

        // Registrar OUService
        builder.Services.AddTransient<OUService>();

        var app = builder.Build();

        // Middleware
        app.UseHttpsRedirection();
        app.UseStaticFiles();

        app.UseRouting();

        // Agregar los middleware de autenticaci�n y autorizaci�n
        app.UseAuthentication();
        app.UseAuthorization();

        app.MapControllerRoute(
            name: "default",
            pattern: "{controller=InicioSesion}/{action=Login}/{id?}");

        app.Run();

        builder.Logging.AddEventLog(); // Registro del proveedor del visor de eventos
    }
}