using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.CompilerServices;
using Google.Apis.Sheets.v4;
using Microsoft.AspNetCore.Components;
using Microsoft.Extensions.Options;
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
    private readonly Settings _settings;

    private readonly ILogger<MainPageViewModel> _logger;
    private readonly FileLoggerService _fileLogger;

    public ObservableCollection<Event> Events { get; private set; } = new();

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

    private List<LogEntry> _logLines = new();
    public IReadOnlyList<LogEntry> MessageLogLines => _logLines;

    private string _messageLog = string.Empty;
    public string MessageLog
    {
        get => _messageLog;
        set
        {
            if (_messageLog != value)
            {
                _messageLog = value;
                OnPropertyChanged();
            }
        }
    }

    public MainPageViewModel(
        EventService eventService,
        ExcelService excelService,
        SheetsService googleSheetsService,
        MailerService mailerService,
        Settings settings,
        FileLoggerService fileLogger,
        ILogger<MainPageViewModel> logger)
    {
        _eventService = eventService;
        _excelService = excelService;
        _googleSheetsService = googleSheetsService;
        _mailerService = mailerService;
        _settings = settings;

        _fileLogger = fileLogger;
        _logger = logger;

        LoadEvents();
    }

    private void LoadEvents()
    {
        Events.Clear();
        foreach (var ev in _eventService.GetAll())
        {
            Events.Add(ev);
        }
    }
    public void OnEventModified()
    {
        // Wymuś odświeżenie powiązanych właściwości
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

        try
        {
            IsBusy = true;

            AddLog($"Checking spreadsheet for event: {SelectedEvent.Name}");

            var spreadsheet = _googleSheetsService
                .Spreadsheets.Get(SelectedEvent.SpreadSheetId)
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
        }
        catch (Exception ex)
        {
            AddLog($"Error: [GenerateTabsAsync] {ex.Message}", ex, LogLevel.Error, MethodBase.GetCurrentMethod().Name);
        }
        finally
        {
            IsBusy = false;
        }
    }

    public async Task ClearExcelAsync()
    {
        if (SelectedEvent == null || IsBusy) 
            return;

        try
        {
            IsBusy = true;
            await Task.Run(() => _excelService.ClearExcel());
        
            UpdateSignupEligibility();

            AddLog("Excel cleared.");
        }
        catch (Exception ex)
        {
            AddLog($"Error: [ClearExcelAsync] {ex.Message}", ex, LogLevel.Error, MethodBase.GetCurrentMethod().Name);
        }
        finally
        {
            IsBusy = false;
        }
    }

    public async Task SendEmailsAsync(EmailType? type = null, bool isTest = false)
    {
        if (SelectedEvent == null)
        {
            AddLog("No event selected.");
            return;
        }

        try
        {
            AddLog($"{(isTest ? "[TEST] " : "")}Preparing student data...");
            var students = _excelService.GetStudentsFromRegularSemestersSheet();

            if (!students.Any())
            {
                AddLog("No students found.");
                return;
            }

            AddLog($"Loaded {students.Count} students. Sending...");

            if (type.HasValue)
            {
                await Task.Run(() => _mailerService.SendEmailsOfType(students, type.Value, isTest));
                AddLog($"Emails of type {type} sent.");
            }
            else
            {
                await Task.Run(() => _mailerService.PrepareAndSendEmailsForRegularSemesters(students, isTest));
                AddLog("All emails sent.");
            }
        }
        catch (Exception ex)
        {
            AddLog($"Error during sending emails: {ex.Message}", ex, LogLevel.Error, MethodBase.GetCurrentMethod().Name);
        }
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

        try
        {
            IsBusy = true;

            AddLog("Searching for new registrations...");
            await Task.Run(() => _excelService.PopulateRegistrationTabForAggregatedData());
            AddLog("New signups populated successfully.");
        }
        catch (Exception ex)
        {
            AddLog($"Error while populating new signups: {ex.Message}", ex, LogLevel.Error, MethodBase.GetCurrentMethod().Name);
        }
        finally
        {
            IsBusy = false;
        }
    }

    private void UpdateSignupEligibility()
    {
        try
        {
            var spreadsheet = _googleSheetsService.Spreadsheets
                .Get(SelectedEvent?.SpreadSheetId)
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

    private void AddLog(string message, Exception? ex = null, LogLevel level = LogLevel.Info, string methodName = "")
    {
        var colorClass = level switch
        {
            LogLevel.Info => "log-info",
            LogLevel.Warning => "log-warning",
            LogLevel.Error => "log-error",
            _ => "log-info"
        };

        var formatted = $"[{DateTime.Now:HH:mm:ss}] {message}";

        _logLines.Insert(0, new LogEntry
        {
            Text = formatted,
            CssClass = colorClass,
            Level = level
        });

        if (level > LogLevel.Warning && ex != null)
        {
            _logger.LogError(ex, $"{level} during {methodName}");
            _fileLogger.LogError(ex, $"{level} during {methodName}");
        }

        OnPropertyChanged(nameof(MessageLogLines));
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
