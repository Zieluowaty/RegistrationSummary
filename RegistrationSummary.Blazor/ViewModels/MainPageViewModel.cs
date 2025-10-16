using System.Collections.ObjectModel;
using System.Reflection;
using System.Text;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using RegistrationSummary.Blazor.Services;
using RegistrationSummary.Common.Enums;
using RegistrationSummary.Common.Models;
using RegistrationSummary.Common.Services;

using LogLevel = RegistrationSummary.Common.Enums.LogLevel;

namespace RegistrationSummary.Blazor.ViewModels;

public class MainPageViewModel : ViewModelBase
{
    private readonly UserContextService _userContextService;

    private bool _eventsLoaded = false;

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

    private bool _isEmailEditorVisible;
    public bool IsEmailEditorVisible
    {
        get => _isEmailEditorVisible;
        set
        {
            _isEmailEditorVisible = value;
            OnPropertyChanged();
        }
    }

    private string? _emailTemplatesContent;
    public string? EmailTemplatesContent
    {
        get => _emailTemplatesContent;
        set
        {
            _emailTemplatesContent = value;
            OnPropertyChanged();
        }
    }

    public MainPageViewModel(
        UserContextService userContextService,
        ILogger<MainPageViewModel> logger,
        IJSRuntime jsRuntime,
        NavigationManager navigationManager,
        ToastService toastService)
        : base(logger, userContextService, jsRuntime, navigationManager, toastService)
    {
        _userContextService = userContextService;
    }

    public void Initialize()
    {
        if (string.IsNullOrEmpty(_userContextService?.Username))
            return;

        LoadEvents();

        if (!_userContextService.MailerService.IsConnectionSuccessful())
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
        foreach (var ev in _userContextService.EventService.GetAll())
        {
            Events.Add(ev);
        }

        _eventsLoaded = true;
    }

    public async Task CloneSelectedEventAsync()
    {
        if (SelectedEvent is null)
            return;

        var confirmed = await ConfirmAsync($"Do you want to clone event \"{SelectedEvent.Name}\"?");
        if (!confirmed)
            return;

        var clone = SelectedEvent.Clone();
        clone.Id = _userContextService.EventService.GenerateNextId();
        clone.Name += " (CLONE)";

        Events.Add(clone);
        SelectedEvent = clone;

        _userContextService.EventService.SaveAll(Events.ToList());

        AddLog($"Event cloned: {clone.Name}", null, LogLevel.Info);

        NavigateTo($"/event/edit/{clone.Id}");
    }

    public void OnEventModified()
    {
        OnPropertyChanged(nameof(SelectedEvent));
        OnPropertyChanged(nameof(CanPopulateNewSignups));

        AddLog($"Event \"{SelectedEvent?.Name}\" has been updated.");

        _userContextService.EventService.SaveAll(Events.ToList());
    }

    public void SelectEvent(string eventName)
    {
        SelectedEvent = Events.FirstOrDefault(e => e.Name == eventName);

        if (SelectedEvent == null)
            return;

        _userContextService.ExcelService.Initialize(
            SelectedEvent,
            _userContextService.SettingConfiguration.RawDataTabName,
            _userContextService.SettingConfiguration.PreprocessedDataTabName,
            _userContextService.SettingConfiguration.SummaryTabName,
            _userContextService.SettingConfiguration.GroupBalanceTabName,
            _userContextService.SettingConfiguration.LeaderText,
            _userContextService.SettingConfiguration.FollowerText,
            _userContextService.SettingConfiguration.SoloText,
            _userContextService.SettingConfiguration.Prices
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

            var spreadsheet = _userContextService.SheetsService
                .Spreadsheets.Get(SelectedEvent.SpreadsheetId)
                .Execute();

            if (spreadsheet.Sheets.Count > 1)
            {
                AddLog("Too many tabs in the spreadsheet.");
                return;
            }

            if (!spreadsheet.Sheets.Any(sheet =>
                    sheet.Properties.Title.Equals(_userContextService.SettingConfiguration.RawDataTabName, StringComparison.OrdinalIgnoreCase)))
            {
                AddLog("No raw data tab found in spreadsheet.");
                return;
            }

            AddLog("Validation passed. Generating tabs...");

            await Task.Run(() => _userContextService.ExcelService.GenerateTabs(SelectedEvent));

            UpdateSignupEligibility();

            AddLog("Tabs generated successfully.");
        });
    }

    public async Task ClearExcelAsync()
    {
        if (SelectedEvent == null || IsBusy)
            return;

        var confirm = await ConfirmAsync("Are you sure you want to clear the spreadsheet data?\nThis action cannot be undone.");
        if (!confirm)
            return;

        await RunWithBusyIndicator(async () =>
        {
            try
            {
                await Task.Run(() => _userContextService.ExcelService.ClearExcel());

                UpdateSignupEligibility();

                AddLog("Excel cleared.");
                ShowToast("Excel cleared");
            }
            catch (Exception ex)
            {
                AddLog(
                    $"Error: [ClearExcelAsync] {ex.Message}", 
                    ex, 
                    LogLevel.Error, 
                    MethodBase.GetCurrentMethod()?.Name ?? "ClearExcelAsync");
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
                var students = _userContextService.ExcelService.GetStudentsFromRegularSemestersSheet();

                if (!students.Any())
                {
                    AddLog("No students found.");
                    ShowToast("No students found");
                    return;
                }

                AddLog($"Loaded {students.Count} students.");

                Dictionary<EmailType, int> summary;
                if (type == EmailType.All)
                {
                    summary = _userContextService.MailerService.GetEmailCountsPerType(students)
                        .Where(kvp => kvp.Value > 0)
                        .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

                    if (!summary.Any())
                    {
                        AddLog("No emails to send.");
                        ShowToast("No emails to send");
                        return;
                    }
                }
                else
                {
                    var count = _userContextService.MailerService.GetPendingRecipientsOfType(students, type).Count;
                    if (count == 0)
                    {
                        AddLog($"No '{type}' emails to send.");
                        ShowToast($"No '{type}' emails to send");
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
                    AddLog("Sending canceled by user.");
                    ShowToast("Sending canceled by user");
                    return;
                }

                if (type == EmailType.All)
                {
                    await Task.Run(() => _userContextService.MailerService.PrepareAndSendEmailsForRegularSemesters(students, isTest, msg => AddLog(msg)));
                    AddLog("Sending emails ended.");
                    ShowToast("Sending emails ended");
                }
                else
                {
                    await Task.Run(() => _userContextService.MailerService.SendEmailsOfType(students, type, isTest, msg => AddLog(msg)));
                    AddLog($"Sending {type} emails ended.");
                    ShowToast($"Sending {type} emails ended");
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
            ShowToast("Cannot populate new signups");
            return;
        }

        await RunWithBusyIndicator(async () =>
        {
            try
            {
                AddLog("Searching for new registrations...");
                await Task.Run(() => _userContextService.ExcelService.PopulateRegistrationTabForAggregatedData());
                AddLog("New signups populated successfully. It can take few seconds for Google Sheet to reload. Be patient.");
                ShowToast("New signups populated successfully.");
            }
            catch (Exception ex)
            {
                AddLog(
                    $"Error while populating new signups: {ex.Message}", 
                    ex, 
                    LogLevel.Error, 
                    MethodBase.GetCurrentMethod()?.Name ?? "PopulateNewSignupsAsync");
            }
        });
    }

    private void UpdateSignupEligibility()
    {
        try
        {
            var spreadsheet = _userContextService.SheetsService
                .Spreadsheets
                .Get(SelectedEvent?.SpreadsheetId)
                .Execute();

            CanPopulateNewSignups = spreadsheet.Sheets
                .Any(sheet => sheet.Properties.Title == _userContextService.SettingConfiguration.PreprocessedDataTabName);
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

    public void OpenEmailEditor()
    {
        var filePath = Path.Combine(_userContextService.FileService.BasePath, "Emails.json");
        if (!File.Exists(filePath))
        {
            AddLog($"Emails.json file not found.",
                    null,
                    LogLevel.Error,
                    MethodBase.GetCurrentMethod()?.Name ?? "OpenEmailEditor");
            ShowToast("Emails.json file not found.");
            return;
        }

        EmailTemplatesContent = File.ReadAllText(filePath);

        IsEmailEditorVisible = true;
    }

    public void SaveEmailTemplates()
    {
        if (EmailTemplatesContent != null)
        {
            var filePath = Path.Combine(_userContextService.FileService.BasePath, "Emails.json");
            File.WriteAllText(filePath, EmailTemplatesContent);
            ShowToast("Templates saved.");
        }

        AddLog("Emails.json was saved successfully.");
        ShowToast("Emails.json was saved successfully.");

        IsEmailEditorVisible = false;
    }

    public void CloseEmailEditor()
    {
        IsEmailEditorVisible = false;
    }
}
