using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using RegistrationSummary.Blazor.ViewModels.Helpers;
using RegistrationSummary.Common.Models;
using RegistrationSummary.Common.Services;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

using LogLevel = RegistrationSummary.Common.Enums.LogLevel;

namespace RegistrationSummary.Blazor.ViewModels;

public class EventModificationPageViewModel : ViewModelBase
{
    private readonly NavigationManager _nav;
    private readonly MainPageViewModel _mainVm;

    private readonly IJSRuntime _jsRuntime;

    public Event? Event { get; private set; }

    private ValidationContext _validationContext;
    private readonly List<ValidationResult> _validationResults = new();

    public List<ColumnBinding> RawColumnBindings => new()
    {
        new("Login", () => Event.RawDataColumns.Login, v => Event.RawDataColumns.Login = v),
        new("Email", () => Event.RawDataColumns.Email, v => Event.RawDataColumns.Email = v),
        new("First Name", () => Event.RawDataColumns.FirstName, v => Event.RawDataColumns.FirstName = v),
        new("Last Name", () => Event.RawDataColumns.LastName, v => Event.RawDataColumns.LastName = v),
        new("Phone", () => Event.RawDataColumns.PhoneNumber, v => Event.RawDataColumns.PhoneNumber = v),
        new("Course", () => Event.RawDataColumns.Course, v => Event.RawDataColumns.Course = v),
        new("Role", () => Event.RawDataColumns.Role, v => Event.RawDataColumns.Role = v),
        new("Partner", () => Event.RawDataColumns.Partner, v => Event.RawDataColumns.Partner = v),
        new("Installment", () => Event.RawDataColumns.Installment, v => Event.RawDataColumns.Installment = v),
        new("Accepted", () => Event.RawDataColumns.Accepted, v => Event.RawDataColumns.Accepted = v)
    };

    public List<ColumnBinding> PreprocessedColumnBindings => new()
    {
        new("Login", () => Event.PreprocessedColumns.Login, v => Event.PreprocessedColumns.Login = v),
        new("Email", () => Event.PreprocessedColumns.Email, v => Event.PreprocessedColumns.Email = v),
        new("First Name", () => Event.PreprocessedColumns.FirstName, v => Event.PreprocessedColumns.FirstName = v),
        new("Last Name", () => Event.PreprocessedColumns.LastName, v => Event.PreprocessedColumns.LastName = v),
        new("Phone", () => Event.PreprocessedColumns.PhoneNumber, v => Event.PreprocessedColumns.PhoneNumber = v),
        new("Course", () => Event.PreprocessedColumns.Course, v => Event.PreprocessedColumns.Course = v),
        new("Role", () => Event.PreprocessedColumns.Role, v => Event.PreprocessedColumns.Role = v),
        new("Partner", () => Event.PreprocessedColumns.Partner, v => Event.PreprocessedColumns.Partner = v),
        new("Installment", () => Event.PreprocessedColumns.Installment, v => Event.PreprocessedColumns.Installment = v),
        new("Accepted", () => Event.PreprocessedColumns.Accepted, v => Event.PreprocessedColumns.Accepted = v)
    };

    public EventModificationPageViewModel(
        MainPageViewModel mainVm,
        NavigationManager nav,
        IJSRuntime jsRuntime,
        FileLoggerService fileLoggerService,
        ILogger<MainPageViewModel> logger)
        : base(logger, fileLoggerService)
    {
        _mainVm = mainVm;
        _nav = nav;
        _jsRuntime = jsRuntime;

        if (_mainVm.SelectedEvent != null)
        {
            Event = _mainVm.SelectedEvent.Clone();
        }

        _validationContext = new ValidationContext(Event);
    }

    public async Task AddCourse()
    {
        var confirm = await _jsRuntime.InvokeAsync<bool>("confirm", "Do you want to add a new course?");
        if (!confirm)
            return;

        Event.Courses.Add(new Course());
        OnPropertyChanged(nameof(Event));
    }


    public void Save()
    {
        if (_mainVm.SelectedEvent == null || Event == null)
            return;

        if (!ValidateColumns())
        {            
            return;
        }

        if (!Validate())
        {
            ClearLog();

            foreach (var msg in GetValidationMessages())
                AddLog(msg, null, LogLevel.Warning);
            return;
        }

        _mainVm.SelectedEvent.Name = Event.Name;
        _mainVm.SelectedEvent.StartDate = Event.StartDate;
        _mainVm.SelectedEvent.SpreadSheetId = Event.SpreadSheetId;
        _mainVm.SelectedEvent.CoursesAreMerged = Event.CoursesAreMerged;
        _mainVm.SelectedEvent.Courses = Event.Courses;
        _mainVm.SelectedEvent.RawDataColumns = Event.RawDataColumns;
        _mainVm.SelectedEvent.PreprocessedColumns = Event.PreprocessedColumns;

        _mainVm.OnEventModified();

        _nav.NavigateTo("/");
    }

    public void Cancel()
    {
        _nav.NavigateTo("/");
    }

    private bool Validate()
    {
        _validationResults.Clear();
        return Validator.TryValidateObject(Event, _validationContext, _validationResults, true);
    }

    public IEnumerable<string> GetValidationMessages()
    => _validationResults.Select(r => r.ErrorMessage ?? "Nieznany błąd");

    public bool ValidateColumns()
    {
        var allCols = new[]
        {
            Event?.RawDataColumns?.Login, Event?.RawDataColumns?.Email, Event?.RawDataColumns?.FirstName,
            Event?.RawDataColumns?.LastName, Event?.RawDataColumns?.PhoneNumber, Event?.RawDataColumns?.Course,
            Event?.RawDataColumns?.Role, Event?.RawDataColumns?.Partner, Event?.RawDataColumns?.Installment,
            Event?.RawDataColumns?.Accepted,

            Event?.PreprocessedColumns?.Login, Event?.PreprocessedColumns?.Email, Event?.PreprocessedColumns?.FirstName,
            Event?.PreprocessedColumns?.LastName, Event?.PreprocessedColumns?.PhoneNumber, Event?.PreprocessedColumns?.Course,
            Event?.PreprocessedColumns?.Role, Event?.PreprocessedColumns?.Partner, Event?.PreprocessedColumns?.Installment,
            Event?.PreprocessedColumns?.Accepted
        };

        var invalid = allCols.Where(c => !IsValidColumn(c)).ToList();

        if (invalid.Any())
        {
            return false;
        }

        return true;
    }

    private static bool IsValidColumn(string? col)
    {
        if (string.IsNullOrWhiteSpace(col))
            return true;

        var trimmed = col.Trim().ToUpperInvariant();

        // Akceptujemy A-Z oraz AA-ZZ
        return Regex.IsMatch(trimmed, @"^[A-Z]{1,2}$") && GetColumnIndex(trimmed) <= GetColumnIndex("ZZ");
    }

    private static int GetColumnIndex(string col)
    {
        int index = 0;
        foreach (char c in col)
        {
            index *= 26;
            index += (c - 'A' + 1);
        }
        return index;
    }

    public void RemoveCourse(Course course)
    {
        Event.Courses.Remove(course);
        OnPropertyChanged(nameof(Event));
    }    
}
