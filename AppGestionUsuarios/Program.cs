using Microsoft.AspNetCore.Server.Kestrel.Core;
using Microsoft.AspNetCore.DataProtection;
using System.IO;

var builder = WebApplication.CreateBuilder(args);

// 1) Data Protection: guarda las claves en disco (opcional, para que sobrevivan reinicios)
builder.Services.AddDataProtection()
    .PersistKeysToFileSystem(new DirectoryInfo(@"C:\Keys\DataProtection"))
    .SetApplicationName("MiApp");

// 2) Cache + Session
builder.Services.AddDistributedMemoryCache();
builder.Services.AddSession(options =>
{
    options.IdleTimeout = TimeSpan.FromHours(8);
    options.Cookie.HttpOnly = true;
    options.Cookie.IsEssential = true;
});

// 3) Kestrel
builder.WebHost.ConfigureKestrel(opts =>
{
    opts.ListenAnyIP(8081, lo => lo.Protocols = HttpProtocols.Http1);
    opts.ListenAnyIP(8082, lo =>
    {
        lo.Protocols = HttpProtocols.Http1;
        lo.UseHttps();
    });
});

// 4) Auth
builder.Services.AddAuthentication("CookieAuth")
    .AddCookie("CookieAuth", o =>
    {
        o.LoginPath = "/InicioSesion/Login";
        o.AccessDeniedPath = "/InicioSesion/Error";
        o.Cookie.HttpOnly = true;
        o.ExpireTimeSpan = TimeSpan.FromHours(8);
    });
builder.Services.AddAuthorization();

// 5) MVC + Controladores
builder.Services.AddControllersWithViews();
builder.Services.AddScoped<AltaUsuarioController>();
builder.Services.AddScoped<AltaMasivaController>();

var app = builder.Build();

// Pipeline
app.UseDeveloperExceptionPage();
app.UseHttpsRedirection();
app.UseStaticFiles();

// **Importante**: Session antes de Authentication/Authorization
app.UseSession();

app.UseRouting();
app.UseAuthentication();
app.UseAuthorization();

app.MapControllerRoute(
    "default", "{controller=InicioSesion}/{action=Login}/{id?}"
);

app.Run();
