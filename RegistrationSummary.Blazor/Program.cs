using Microsoft.Extensions.Options;
using RegistrationSummary.Common.Configurations;
using RegistrationSummary.Common.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using System.Text.Json;
using RegistrationSummary.Blazor.ViewModels;
using RegistrationSummary.Common.Services;
using RegistrationSummary.Blazor.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Configuration.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

builder.Services.AddSingleton<FileService>();

builder.Services.AddSingleton(provider =>
{
    var fileService = provider.GetRequiredService<FileService>();
    return fileService.Load<SettingConfiguration>("Settings.json");
});

builder.Services.AddSingleton(provider =>
{
    var settingConfiguration = provider.GetRequiredService<SettingConfiguration>();

    var credentialPath = Path.Combine(settingConfiguration.ConfigFilesRoot, "Credentials.json");

    using var stream = new FileStream(credentialPath, FileMode.Open, FileAccess.Read);
    var credential = GoogleCredential.FromStream(stream)
        .CreateScoped(SheetsService.Scope.Spreadsheets);

    return new SheetsService(new BaseClientService.Initializer
    {
        HttpClientInitializer = credential,
        ApplicationName = "RegistrationSummary"
    });
});

builder.Services.AddSingleton(provider =>
{
	var settings = provider.GetRequiredService<SettingConfiguration>();
	var filePath = Path.Combine(settings.ConfigFilesRoot, "Emails.json");

	if (!File.Exists(filePath))
		throw new FileNotFoundException($"Nie znaleziono {filePath}");

	var json = File.ReadAllText(filePath);
	return JsonSerializer.Deserialize<List<EmailJsonModel>>(json) ?? new();
});

builder.Services.AddSingleton<EventService>();
builder.Services.AddSingleton<ExcelService>();

builder.Services.AddSingleton(provider =>
{
    var settings = provider.GetRequiredService<SettingConfiguration>();
    var excelService = provider.GetRequiredService<ExcelService>();
    var emailsTemplates = provider.GetRequiredService<List<EmailJsonModel>>();

    return new MailerService(
        settings.MailerConfiguration,
        excelService,
        emailsTemplates);
});

builder.Services.AddSingleton<FileLoggerService>();

builder.Services.AddScoped<MainPageViewModel>();
builder.Services.AddTransient<EventModificationPageViewModel>();

builder.Services.AddSingleton<ToastService>();

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
