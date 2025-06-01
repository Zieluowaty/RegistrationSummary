using RegistrationSummary.Common.Configurations;
using RegistrationSummary.Common.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using RegistrationSummary.Blazor.ViewModels;
using RegistrationSummary.Common.Services;
using RegistrationSummary.Blazor.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Configuration.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

builder.Services.AddScoped<FileService>();

builder.Services.AddScoped(provider =>
{
    var fileService = provider.GetRequiredService<FileService>();
    return fileService.Load<SettingConfiguration>("Settings.json");
});

builder.Services.AddScoped(provider =>
{
    var fileService = provider.GetRequiredService<FileService>();
    var credentialsPath = Path.Combine(fileService.BasePath, "Credentials.json");

    if (!File.Exists(credentialsPath))
        throw new FileNotFoundException($"Credentials file not found: {credentialsPath}");

    var credentialsJson = File.ReadAllText(credentialsPath);

    var credential = GoogleCredential
        .FromJson(credentialsJson)
        .CreateScoped(SheetsService.Scope.Spreadsheets);

    return new SheetsService(new BaseClientService.Initializer
    {
        HttpClientInitializer = credential,
        ApplicationName = "RegistrationSummary"
    });
});

builder.Services.AddScoped(provider =>
{
    var fileService = provider.GetRequiredService<FileService>();
 
	return fileService.Load<List<EmailJsonModel>>("Emails.json") ?? new();
});

builder.Services.AddScoped<EventService>();
builder.Services.AddScoped<ExcelService>();

builder.Services.AddScoped(provider =>
{
    var settings = provider.GetRequiredService<SettingConfiguration>();
    var excelService = provider.GetRequiredService<ExcelService>();
    var emailsTemplates = provider.GetRequiredService<List<EmailJsonModel>>();

    return new MailerService(
        settings.MailerConfiguration,
        excelService,
        emailsTemplates);
});

builder.Services.AddScoped<FileLoggerService>();

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
