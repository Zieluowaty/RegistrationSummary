using System.Collections.ObjectModel;
using System.Reflection;
using System.Text;
using Google.Apis.Sheets.v4;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using RegistrationSummary.Blazor.Services;
using RegistrationSummary.Common.Configurations;
using RegistrationSummary.Common.Enums;
using RegistrationSummary.Common.Models;
using RegistrationSummary.Common.Services;

using LogLevel = RegistrationSummary.Common.Enums.LogLevel;

namespace RegistrationSummary.Blazor.ViewModels;

public class MainPageViewModel : ViewModelBase
{
    private readonly EventService _eventService;
    private readonly ExcelService _excelService;
    private readonly MailerService _mailerService;
    private readonly SheetsService _googleSheetsService;
    
    private readonly SettingConfiguration _settings;
    
    public ObservableCollection<Event> Events { get; private set; } = new();
    public bool CanEditSelected => SelectedEvent is not null;

    private Event? _selectedEvent;
    public Event? SelectedEvent
    {
        get => _selectedEvent;
        private set
        {
            if (_selectedEvent != value)
            {
                _selectedEvent = value;
                OnPropertyChanged();

                CanPopulateNewSignups = _selectedEvent?.CoursesAreMerged == true;
            }
        }
    }

    private bool _canPopulateNewSignups;
    public bool CanPopulateNewSignups
    {
        get => _canPopulateNewSignups;
        private set
        {
            if (_canPopulateNewSignups != value)
            {
                _canPopulateNewSignups = value;
                OnPropertyChanged();
            }
        }
    }

    public MainPageViewModel(
        EventService eventService,
        ExcelService excelService,
        SheetsService googleSheetsService,
        MailerService mailerService,
        SettingConfiguration settings,
        FileLoggerService fileLoggerService,
        ILogger<MainPageViewModel> logger,
        IJSRuntime jsRuntime,
        NavigationManager navigationManager,
        ToastService toastService)
        : base(logger, fileLoggerService, jsRuntime, navigationManager, toastService)
    {
        _eventService = eventService;
        _excelService = excelService;
        _googleSheetsService = googleSheetsService;
        _mailerService = mailerService;
        _settings = settings;

        LoadEvents();

        if (!_mailerService.IsConnectionSuccessful())
        {
            AddLog("Mail server connection failed. Please check your settings.", null, LogLevel.Error);
            ShowToast("Mail server connection failed. Please check your settings.");
        }
        else
        {
            AddLog("Mail server connection successful.", null, LogLevel.Info);
        }
    }

    private void LoadEvents()
    {
        Events.Clear();
        foreach (var ev in _eventService.GetAll())
        {
            Events.Add(ev);
        }
    }

    public async Task CloneSelectedEventAsync()
    {
        if (SelectedEvent is null)
            return;

        var confirmed = await ConfirmAsync($"Do you want to clone event \"{SelectedEvent.Name}\"?");
        if (!confirmed)
            return;

        var clone = SelectedEvent.Clone();
        clone.Id = _eventService.GenerateNextId();
        clone.Name += " (CLONE)";

        Events.Add(clone);
        SelectedEvent = clone;

        _eventService.SaveAll(Events.ToList());

        AddLog($"Event cloned: {clone.Name}", null, LogLevel.Info);

        NavigateTo($"/event/edit/{clone.Id}");
    }

    public void OnEventModified()
    {
        OnPropertyChanged(nameof(SelectedEvent));
        OnPropertyChanged(nameof(CanPopulateNewSignups));

        AddLog($"Event \"{SelectedEvent?.Name}\" has been updated.");

        _eventService.SaveAll(Events.ToList());
    }

    public void SelectEvent(string eventName)
    {
        SelectedEvent = Events.FirstOrDefault(e => e.Name == eventName);

        if (SelectedEvent == null)
            return;

        _excelService.Initialize(
            SelectedEvent,
            _settings.RawDataTabName,
            _settings.PreprocessedDataTabName,
            _settings.SummaryTabName,
            _settings.GroupBalanceTabName,
            _settings.LeaderText,
            _settings.FollowerText,
            _settings.SoloText,
            _settings.Prices
        );

        UpdateSignupEligibility();

        AddLog($"Loaded event: {SelectedEvent.Name}");
    }

    public async Task GenerateTabsAsync()
    {
        if (SelectedEvent == null || IsBusy)
            return;

        await RunWithBusyIndicator(async () =>
        {
            AddLog($"Checking spreadsheet for event: {SelectedEvent.Name}");

            var spreadsheet = _googleSheetsService
                .Spreadsheets.Get(SelectedEvent.SpreadsheetId)
                .Execute();

            if (spreadsheet.Sheets.Count > 1)
            {
                AddLog("Too many tabs in the spreadsheet.");
                return;
            }

            if (!spreadsheet.Sheets.Any(sheet =>
                    sheet.Properties.Title.Equals(_settings.RawDataTabName, StringComparison.OrdinalIgnoreCase)))
            {
                AddLog("No raw data tab found in spreadsheet.");
                return;
            }

            AddLog("Validation passed. Generating tabs...");

            await Task.Run(() => _excelService.GenerateTabs(SelectedEvent));

            UpdateSignupEligibility();

            AddLog("Tabs generated successfully.");
        });
    }

    public async Task ClearExcelAsync()
    {
        if (SelectedEvent == null || IsBusy)
            return;

        await RunWithBusyIndicator(async () =>
        {
            try
            {
                await Task.Run(() => _excelService.ClearExcel());

                UpdateSignupEligibility();

                AddLog("Excel cleared.");
            }
            catch (Exception ex)
            {
                AddLog($"Error: [ClearExcelAsync] {ex.Message}", ex, LogLevel.Error, MethodBase.GetCurrentMethod()?.Name ?? "ClearExcelAsync");
            }
        });
    }


    public async Task SendEmailsAsync(EmailType type, bool isTest = false)
    {
        if (SelectedEvent == null)
        {
            AddLog("No event selected.");
            return;
        }

        await RunWithBusyIndicator(async () =>
        {
            try
            {
                AddLog($"{(isTest ? "[TEST] " : "")}Preparing student data...");
                var students = _excelService.GetStudentsFromRegularSemestersSheet();

                if (!students.Any())
                {
                    AddLog("No students found.");
                    ShowToast("No students found.");
                    return;
                }

                AddLog($"Loaded {students.Count} students.");

                Dictionary<EmailType, int> summary;
                if (type == EmailType.All)
                {
                    summary = _mailerService.GetEmailCountsPerType(students)
                        .Where(kvp => kvp.Value > 0)
                        .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

                    if (!summary.Any())
                    {
                        AddLog("No emails to send.");
                        ShowToast("No emails to send.");
                        return;
                    }
                }
                else
                {
                    var count = _mailerService.GetPendingRecipientsOfType(students, type).Count;
                    if (count == 0)
                    {
                        ShowToast($"No '{type}' emails to send.");
                        AddLog($"No '{type}' emails to send.");
                        return;
                    }

                    summary = new Dictionary<EmailType, int> { [type] = count };
                }

                // Potwierdzenie
                var sb = new StringBuilder();
                sb.AppendLine("Do you want to send:");

                foreach (var kvp in summary)
                    sb.AppendLine($"{kvp.Value} \"{kvp.Key}\" emails");

                var confirm = await ConfirmAsync(sb.ToString());
                if (!confirm)
                {
                    ShowToast("Sending canceled by user.");
                    AddLog("Sending canceled by user.");
                    return;
                }

                if (type == EmailType.All)
                {
                    await Task.Run(() => _mailerService.PrepareAndSendEmailsForRegularSemesters(students, isTest, msg => AddLog(msg)));
                    ShowToast("Sending emails ended.");
                    AddLog("Sending emails ended.");
                }
                else
                {
                    await Task.Run(() => _mailerService.SendEmailsOfType(students, type, isTest, msg => AddLog(msg)));
                    ShowToast($"Sending {type} emails ended.");
                    AddLog($"Sending {type} emails ended.");
                }
            }
            catch (Exception ex)
            {
                AddLog($"Error during sending emails: {ex.Message}", 
                    ex, 
                    LogLevel.Error, 
                    MethodBase.GetCurrentMethod()?.Name ?? "SendEmailsAsync");
            }
        });
    }

    public async Task PopulateNewSignupsAsync()
    {
        if (SelectedEvent == null || IsBusy)
            return;

        if (!CanPopulateNewSignups)
        {
            AddLog("Cannot populate new signups – required tab does not exist.");
            return;
        }

        await RunWithBusyIndicator(async () =>
        {
            try
            {
                AddLog("Searching for new registrations...");
                await Task.Run(() => _excelService.PopulateRegistrationTabForAggregatedData());
                AddLog("New signups populated successfully.");
            }
            catch (Exception ex)
            {
                AddLog($"Error while populating new signups: {ex.Message}", ex, LogLevel.Error, MethodBase.GetCurrentMethod()?.Name ?? "PopulateNewSignupsAsync");
            }
        });
    }

    private void UpdateSignupEligibility()
    {
        try
        {
            var spreadsheet = _googleSheetsService.Spreadsheets
                .Get(SelectedEvent?.SpreadsheetId)
                .Execute();

            CanPopulateNewSignups = spreadsheet.Sheets
                .Any(sheet => sheet.Properties.Title == _settings.PreprocessedDataTabName);
        }
        catch (Exception ex)
        {
            AddLog("Could not verify preprocessed tab existence.", ex, LogLevel.Error, MethodBase.GetCurrentMethod().Name);
            CanPopulateNewSignups = false;
        }
    }

    private static readonly Dictionary<string, int> DayOrder = new()
    {
        ["Monday"] = 1,
        ["Tuesday"] = 2,
        ["Wednesday"] = 3,
        ["Thursday"] = 4,
        ["Friday"] = 5,
        ["Saturday"] = 6,
        ["Sunday"] = 7
    };

    public IEnumerable<Course> SortedCourses =>
        SelectedEvent?.Courses?
            .OrderBy(c => DayOrder.FirstOrDefault(day => day.Key == c.DayOfWeek).Value)
            .ThenBy(c => c.Time)
            ?? Enumerable.Empty<Course>();
}
