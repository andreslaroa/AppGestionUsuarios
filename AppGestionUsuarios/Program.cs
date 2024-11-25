internal class Program
{
    private static void Main(string[] args)
    {
        var builder = WebApplication.CreateBuilder(args);

        // Agregar servicios de autenticaci�n por cookies
        builder.Services.AddAuthentication("CookieAuth")
            .AddCookie("CookieAuth", options =>
            {
                options.LoginPath = "/InicioSesion/Login"; // Ruta a la p�gina de inicio de sesi�n
                options.AccessDeniedPath = "/InicioSesion/Error"; // Ruta en caso de acceso denegado
            });

        builder.Services.AddAuthorization(); //Usar autenticaci�n en la aplicaci�n
        builder.Services.AddControllersWithViews();


        var app = builder.Build();

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

        builder.Logging.AddEventLog(); //Registro del proveedor del visor de eventos

    }
}