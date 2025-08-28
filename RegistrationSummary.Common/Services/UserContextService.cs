using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using RegistrationSummary.Common.Configurations;
using RegistrationSummary.Common.Models;

namespace RegistrationSummary.Common.Services;

public class UserContextService
{
    public FileService? FileService { get; set; }

    public SettingConfiguration SettingConfiguration { get; set; }

    public SheetsService? SheetsService { get; set; }

    public List<EmailJsonModel> EmailsTemplates { get; set; }

    public EventService? EventService { get; set; }

    public ExcelService? ExcelService { get; set; } 

    public MailerService? MailerService { get; set; }

    public FileLoggerService? FileLoggerService { get; set; }


    private string? _username;
    public string? Username 
    { 
        get => _username; 
        set
        {
            _username = value;

            if (_username is null)
                return;

            FileService = new FileService(this);
            SettingConfiguration = FileService.Load<SettingConfiguration>("Settings.json");
            EmailsTemplates = FileService.Load<List<EmailJsonModel>>("Emails.json");

            try
            { 
                InitializeSheetsService();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to initialize SheetsService. Please check your Credentials.json file.", ex);
            }

            EventService = new EventService(FileService);
            ExcelService = new ExcelService(SheetsService);
            MailerService = new MailerService(SettingConfiguration.MailerConfiguration, ExcelService, EmailsTemplates);
            FileLoggerService = new FileLoggerService(SettingConfiguration, FileService);
        }
    }

    private void InitializeSheetsService()
    {
        var fileService = FileService ?? throw new InvalidOperationException("FileService is not initialized.");
        var credentialsPath = Path.Combine(fileService.BasePath, "Credentials.json");
        
        if (!File.Exists(credentialsPath))
            throw new FileNotFoundException($"Credentials file not found: {credentialsPath}");
        var credentialsJson = File.ReadAllText(credentialsPath);
        
        var credential = GoogleCredential
            .FromJson(credentialsJson)
            .CreateScoped(SheetsService.Scope.Spreadsheets);

        SheetsService = new SheetsService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = "RegistrationSummary"
        });
    }
}