using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Server.Kestrel.Core;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Azure.Identity;
using AppGestionUsuarios.Controllers;
using AppGestionUsuarios.Notificaciones;
using AppGestionUsuarios.Configuration;

var builder = WebApplication.CreateBuilder(args);


// 2) Kestrel
builder.WebHost.ConfigureKestrel(opts =>
{
    opts.ListenAnyIP(8081, lo => lo.Protocols = HttpProtocols.Http1);
    opts.ListenAnyIP(8082, lo =>
    {
        lo.Protocols = HttpProtocols.Http1;
        lo.UseHttps();
    });
});

// 3) Auth
builder.Services.AddAuthentication("CookieAuth")
    .AddCookie("CookieAuth", o =>
    {
        o.LoginPath = "/InicioSesion/Login";
        o.AccessDeniedPath = "/InicioSesion/Error";
    });
builder.Services.AddAuthorization();

// 4) MVC + Controladores
builder.Services.AddControllersWithViews();
builder.Services.AddScoped<AltaUsuarioController>();
builder.Services.AddScoped<AltaMasivaController>();


var app = builder.Build();

// 6) Mostrar siempre excepción en Dev
app.UseDeveloperExceptionPage();

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();
app.UseAuthentication();
app.UseAuthorization();

// 7) Routing
app.MapControllerRoute(
    "default", "{controller=InicioSesion}/{action=Login}/{id?}"
);

app.Run();
