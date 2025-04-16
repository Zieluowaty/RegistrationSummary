using Microsoft.Extensions.Options;
using RegistrationSummary.Common.Configurations;
using RegistrationSummary.Common.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System.Text.Json;
using RegistrationSummary.Blazor.ViewModels;
using RegistrationSummary.Common.Services;

// 1. Inicjalizacja buildera
var builder = WebApplication.CreateBuilder(args);

// 2. Konfiguracja JSON
builder.Configuration.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

builder.Services.AddSingleton<FileService>();

// 3. Wczytaj konfiguracje (appsettings.json → klasy konfiguracyjne)
builder.Services.Configure<MailerConfiguration>(builder.Configuration.GetSection("MailerConfiguration"));
builder.Services.Configure<ColumnsConfiguration>(builder.Configuration.GetSection("ColumnsConfiguration"));

builder.Services.AddSingleton(provider =>
{
    var fileService = provider.GetRequiredService<FileService>();

    return fileService.Load<Settings>("Settings.json");    
});

// 4. Google SheetsService singleton z użyciem Credentials.json
builder.Services.AddSingleton(provider =>
{
    var settings = provider.GetRequiredService<IOptions<Settings>>().Value;

    var credentialPath = Path.Combine(settings.ConfigFilesRoot ?? "C:/RegistrationSummary", "Credentials.json");

    using var stream = new FileStream(credentialPath, FileMode.Open, FileAccess.Read);
    var credential = GoogleCredential.FromStream(stream)
        .CreateScoped(SheetsService.Scope.Spreadsheets);

    return new SheetsService(new BaseClientService.Initializer
    {
        HttpClientInitializer = credential,
        ApplicationName = "RegistrationSummary"
    });
});

// 5. Emails.json → EmailJsonModel[]
builder.Services.AddSingleton(provider =>
{
	var settings = provider.GetRequiredService<IOptions<Settings>>().Value;
	var filePath = Path.Combine(settings.ConfigFilesRoot ?? "C:/RegistrationSummary", "Emails.json");

	if (!File.Exists(filePath))
		throw new FileNotFoundException($"Nie znaleziono {filePath}");

	var json = File.ReadAllText(filePath);
	return JsonSerializer.Deserialize<List<EmailJsonModel>>(json) ?? new();
});

// 6. Rejestracja pozostałych serwisów
builder.Services.AddSingleton<EventService>();
builder.Services.AddSingleton<ExcelService>();
builder.Services.AddSingleton<MailerService>();
builder.Services.AddSingleton<FileLoggerService>();
builder.Services.AddSingleton<MainPageViewModel>();
builder.Services.AddTransient<EventModificationPageViewModel>();

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
