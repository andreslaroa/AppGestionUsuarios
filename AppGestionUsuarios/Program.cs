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

//// 0) Bind de configuración
//builder.Services.Configure<AzureAdSettings>(
//    builder.Configuration.GetSection("AzureAd")
//);
//builder.Services.Configure<AzureAdSyncSettings>(
//    builder.Configuration.GetSection("AzureAdSync")
//);
//builder.Services.Configure<SmtpSettings>(
//    builder.Configuration.GetSection("SmtpSettings")
//);

// 1) Registrar GraphServiceClient via Options
//builder.Services.AddSingleton<GraphServiceClient>(sp =>
//{
//    var opts = sp.GetRequiredService<IOptions<AzureAdSettings>>().Value;
//    if (string.IsNullOrEmpty(opts.TenantId) ||
//        string.IsNullOrEmpty(opts.ClientId) ||
//        string.IsNullOrEmpty(opts.ClientSecret))
//        throw new InvalidOperationException("Faltan valores en AzureAd");

//    var cred = new ClientSecretCredential(
//        opts.TenantId, opts.ClientId, opts.ClientSecret
//    );
//    return new GraphServiceClient(
//        cred,
//        new[] { "https://graph.microsoft.com/.default" }
//    );
//});

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

// 5) EmailNotifier (comenta para aislar)
builder.Services.AddTransient<EmailNotifier>();

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
