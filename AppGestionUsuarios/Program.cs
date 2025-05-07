using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Server.Kestrel.Core;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using AppGestionUsuarios.Controllers;

var builder = WebApplication.CreateBuilder(args);

// 1) Kestrel
builder.WebHost.ConfigureKestrel(options =>
{
    options.ListenAnyIP(8081, lo => lo.Protocols = HttpProtocols.Http1);
    options.ListenAnyIP(8082, lo =>
    {
        lo.Protocols = HttpProtocols.Http1;
        lo.UseHttps();
    });
});

// 2) AuthN / AuthZ
builder.Services.AddAuthentication("CookieAuth")
    .AddCookie("CookieAuth", opts =>
    {
        opts.LoginPath = "/InicioSesion/Login";
        opts.AccessDeniedPath = "/InicioSesion/Error";
    });
builder.Services.AddAuthorization();

// 3) MVC + controladores en DI
builder.Services.AddControllersWithViews();
builder.Services.AddScoped<AltaUsuarioController>();
builder.Services.AddScoped<AltaMasivaController>();

var app = builder.Build();

// 4) Middlewares
if (app.Environment.IsDevelopment())
    app.UseDeveloperExceptionPage();

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();
app.UseAuthentication();
app.UseAuthorization();

// 5) Routing
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=InicioSesion}/{action=Login}/{id?}"
);

app.Run();
