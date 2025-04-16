using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Google.Apis.Sheets.v4;
using Newtonsoft.Json;
using RegistrationSummary.Common.Configurations;
using RegistrationSummary.Common.Enums;
using RegistrationSummary.Common.Interfaces;
using RegistrationSummary.Common.Models;
using RegistrationSummary.Common.Services;
using RegistrationSummary.Maui.Services;

namespace RegistrationSummary.Maui.ViewModels;

public partial class MainPageViewModel : ObservableObject
{
    [ObservableProperty]
    public string _installationPath;

    [ObservableProperty]
    public List<Event>? _events;
	public List<EmailJsonModel>? EmailsTemplates { get; set; }

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(EditEventCommand))]
    [NotifyCanExecuteChangedFor(nameof(CloneEventCommand))]
    [NotifyCanExecuteChangedFor(nameof(ClearExcelCommand))]
    [NotifyCanExecuteChangedFor(nameof(GenerateTabsCommand))]
    [NotifyCanExecuteChangedFor(nameof(PopulateNewSingupsCommand))]
    [NotifyCanExecuteChangedFor(nameof(SendEmailsCommand))]
    private Event? _selectedEvent;

    partial void OnSelectedEventChanged(Event? oldValue, Event? newValue)
    {
        if (newValue != null)
        {
            SheetsService = GoogleSheetsService.CreateSheetService(MauiProgram.CredentialsFilePath);

            PopulateNewSingupsButtonIsVisible = newValue.CoursesAreMerged;
        }
        else
        {
            SheetsService = null;
        }
    }

    private readonly MailerConfiguration _mailerConfiguration;
    private readonly IDialogService _dialogService;

	private string _emailsTemplatesFilePath;
	private string _rawDataTabName;
	private string _preprocessedDataTabName;
	private string _summaryTabName;
    private string _groupBalanceTabName;
    private string _leaderText;
    private string _followerText;
    private string _soloText;
    private string _testMailRecepientEmailAddress;
    private int[] _prices;

    private SheetsService? SheetsService { get; set; }

	private ExcelService? ExcelService => SelectedEvent == null || SheetsService == null 
		? null
		: new ExcelService(SheetsService, SelectedEvent, _rawDataTabName, _preprocessedDataTabName, _summaryTabName, _groupBalanceTabName, _leaderText, _followerText, _soloText, _prices);

	private MailerService? MailerService 
    => SelectedEvent == null || SheetsService == null || ExcelService == null || _mailerConfiguration == null || EmailsTemplates == null
		? null
		: new MailerService(_mailerConfiguration.Mail, _mailerConfiguration.Password, _mailerConfiguration.ServerName, _mailerConfiguration.ServerPort, ExcelService, EmailsTemplates);

	public MainPageViewModel(MailerConfiguration mailerConfiguration, IDialogService dialogService,
		string emailsTemplatesFilePath, string rawDataTabName, string preprocessedDataTabName, 
        string summaryTabName, string groupBalanceTabName, string leaderText, string followerText, string soloText,
        string testMailRecepientEmailAddress, int[] prices)
	{
		_mailerConfiguration = mailerConfiguration;
		_dialogService = dialogService;
		_emailsTemplatesFilePath = emailsTemplatesFilePath;
		_rawDataTabName = rawDataTabName;
		_preprocessedDataTabName = preprocessedDataTabName;
		_summaryTabName = summaryTabName;
        _groupBalanceTabName = groupBalanceTabName;
        _leaderText = leaderText;
        _followerText = followerText;
        _soloText = soloText;
        _testMailRecepientEmailAddress = testMailRecepientEmailAddress;
        _prices = prices;

	    Events = [.. EventsService.LoadEvents().OrderByDescending(ev => ev.StartDate)];
		LoadEmailsTemplates();

        _installationPath = AppDomain.CurrentDomain.BaseDirectory;
	}

    public void RefreshData()
    {
        if (SelectedEvent == null)
            return;

        var id = SelectedEvent.Id;
        Events = EventsService.LoadEvents();
        
        SelectedEvent = null;
        SelectedEvent = Events.Single(ev => ev.Id == id);
    }

    [RelayCommand(CanExecute = nameof(IsEventSelected))]
    private async Task ClearExcel() 
	{
        var result = await _dialogService.ShowConfirmationAsync("WARNING!", "HEY DUDE, THERE ARE DATA, ARE YOU SURE YOU WANT TO DELETE THEM?");
        if (!result)
            return;

        try
        {
            ExcelService?.ClearExcel();
        }
        catch (Exception ex)
        {
            await _dialogService.ShowAlertAsync("No Tabs", ex.Message);
            return;
        }
    }

    [RelayCommand(CanExecute = nameof(IsEventSelected))]
    private async Task GenerateTabs() 
    {
        if (SheetsService?.Spreadsheets.Get(SelectedEvent.SpreadSheetId).Execute().Sheets.Count > 1)
        {
            await _dialogService.ShowAlertAsync("Too Many Generated Tabs", "The file has more than just a data tab.");
            return;
        }

        if (!(SheetsService?.Spreadsheets.Get(SelectedEvent.SpreadSheetId).Execute().Sheets.Any(sheet => sheet.Properties.Title.Equals(_rawDataTabName)) ?? false))
        {
            await _dialogService.ShowAlertAsync("No Raw Data Tab", "The file does not contain a raw data tab, check the name of the tab in configuration of the event.");
            return;
        }

        ExcelService?.SetUpRegistrationsEditableTab();
        ExcelService?.SetUpSummaryTab();
        ExcelService?.SetUpTabForAccountants();
        ExcelService?.SetUpTabForNoPayments();
        ExcelService?.SetUpGroupBalanceTab();
    }

    [RelayCommand(CanExecute = nameof(PopulateNewSignupsCanExecute))]
    private async Task PopulateNewSingups()
    {
        if (!(SheetsService?.Spreadsheets.Get(SelectedEvent.SpreadSheetId).Execute().Sheets.Any(sheet => sheet.Properties.Title.Equals(_preprocessedDataTabName)) ?? false))
        {
            await _dialogService.ShowAlertAsync("No Tab With Preprocessed Data", "The file does not contain tab with preprocessed data, check the name of the tab in configuration of the event.");
            return;
        }

        ExcelService?.PopulateRegistrationTabForAggregatedData();
    }

    private bool PopulateNewSignupsCanExecute => IsEventSelected && SelectedEvent.CoursesAreMerged;

    [ObservableProperty]
    private bool _populateNewSingupsButtonIsVisible;

    [RelayCommand(CanExecute = nameof(IsEventSelected))]
    private async Task SendEmails(string parameters)
    {
        object[] parameterList = parameters.Split(",");
        bool isTestEmails = false;

        if (parameterList.Length == 0)
            return;

        if (parameterList.Length == 2)
            isTestEmails = bool.Parse(parameterList[1].ToString());

        EmailType emailType = (EmailType)Enum.Parse(typeof(EmailType), (string)parameterList[0]);

        if (!(MailerService?.IsConnectionSuccessful() ?? false))
        {
            await _dialogService.ShowAlertAsync("Incorrect email login details", "The appsettings.json file contains incorrect login details, or the mail server is unavailable.");
            return;
        }

        var students = ExcelService?.GetStudentsFromRegularSemestersSheet();

        if (students == null)
        {
            await _dialogService.ShowAlertAsync("No Students Found", "The app was unable to retrieve the list of students to send them emails.");
            return;
        }

        if (string.IsNullOrEmpty(_testMailRecepientEmailAddress))
        {
            await _dialogService.ShowAlertAsync("Test email recepient not set", "The Setting.json file does not contain test mail recepient field and cannot sent test email." +
                Environment.NewLine + "Please provide new line in Setting.json file, as follow:" +
                Environment.NewLine + "\"TestMailRecepientEmailAddress\": \"youremail@yourdomain.com\",");
            return;
        }

        if (isTestEmails)
            students.ForEach(student => student.Email = _testMailRecepientEmailAddress);

        switch (emailType)
        {
            case EmailType.Confirmation:
                MailerService?.SendEmailsOfType(students, EmailType.Confirmation);
                break;
            case EmailType.WaitingList:
                MailerService?.SendEmailsOfType(students, EmailType.WaitingList);
                break;
            case EmailType.NotEnoughPeople:
                MailerService?.SendEmailsOfType(students, EmailType.NotEnoughPeople);
                break;
            case EmailType.FullClass:
                MailerService?.SendEmailsOfType(students, EmailType.FullClass);
                break;
            case EmailType.MissingPartner:
                MailerService?.SendEmailsOfType(students, EmailType.MissingPartner);
                break;
            case EmailType.All:
                MailerService?.PrepareAndSendEmailsForRegularSemesters(students);
                break;
        }
    }

    [RelayCommand]
    private async Task AddNewEvent()
    {
        await Shell.Current.GoToAsync("eventmodification", new Dictionary<string, object>
        {
            { "DialogService", _dialogService },
            { "SelectedEvent", null }
        });
    }

    [RelayCommand(CanExecute = nameof(IsEventSelected))]
    private async Task EditEvent()
    {
        await Shell.Current.GoToAsync("eventmodification", new Dictionary<string, object>
        {
            { "DialogService", _dialogService },
            { "SelectedEvent", SelectedEvent }
        });
    }

    [RelayCommand(CanExecute = nameof(IsEventSelected))]
    private async Task CloneEvent()
    {
        var clonedEvent = SelectedEvent.Clone();
        clonedEvent.Name = clonedEvent.Name + " CLONED";

        await Shell.Current.GoToAsync("eventmodification", new Dictionary<string, object>
        {
            { "DialogService", _dialogService },
            { "SelectedEvent", clonedEvent }
        });
    }

    private bool IsEventSelected => SelectedEvent != null;

    private void LoadEmailsTemplates()
	{
		if (string.IsNullOrEmpty(_emailsTemplatesFilePath))
		{
			throw new Exception("There is no emails templates data file path provided in appsettings.json.");
		}

		string jsonData = File.ReadAllText(_emailsTemplatesFilePath, Encoding.UTF8);
		EmailsTemplates = JsonConvert.DeserializeObject<List<EmailJsonModel>>(jsonData);

		if (EmailsTemplates == null)
		{
			throw new Exception("Cannot load emails templates data file.");
		}
	}
}