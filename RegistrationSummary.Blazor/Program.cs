using RegistrationSummary.Blazor.ViewModels;
using RegistrationSummary.Common.Services;
using RegistrationSummary.Blazor.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Configuration.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

builder.Services.AddScoped<UserContextService>();

builder.Services.AddScoped<MainPageViewModel>();
builder.Services.AddTransient<EventModificationPageViewModel>();
builder.Services.AddTransient<LoginPageViewModel>();

builder.Services.AddScoped<ToastService>();

builder.Services.AddScoped<AuthenticationService>();

// 7. Blazor i routing
builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();

var app = builder.Build();

// 8. Middleware
if (!app.Environment.IsDevelopment())
{
	app.UseExceptionHandler("/Error");
	app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();

app.MapBlazorHub();
app.MapFallbackToPage("/_Host");

app.Run();
