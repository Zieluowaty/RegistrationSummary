﻿using Google.Apis.Sheets.v4;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using RegistrationSummary.Blazor.Services;
using RegistrationSummary.Blazor.ViewModels.Helpers;
using RegistrationSummary.Common.Configurations;
using RegistrationSummary.Common.Models;
using RegistrationSummary.Common.Services;
using System.ComponentModel.DataAnnotations;
using System.Net;
using System.Text.RegularExpressions;

using LogLevel = RegistrationSummary.Common.Enums.LogLevel;

namespace RegistrationSummary.Blazor.ViewModels;

public class EventModificationPageViewModel : ViewModelBase
{
    private readonly NavigationManager _nav;
    private readonly MainPageViewModel _mainVm; 
    private readonly SheetsService _sheetsService;

    private readonly IJSRuntime _jsRuntime;
    private readonly ToastService _toastService;

    public Event? Event { get; private set; }

    private ValidationContext _validationContext;
    private readonly List<ValidationResult> _validationResults = new();

    private Event? _initialEventSnapshot;
    
    public bool HasUnsavedChanges => !Event.Equals(_initialEventSnapshot);

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
        SheetsService sheetsService,
        IJSRuntime jsRuntime,
        ToastService toastService,
        FileLoggerService fileLoggerService,
        ILogger<MainPageViewModel> logger)
        : base(logger, fileLoggerService)
    {
        _mainVm = mainVm;
        _sheetsService = sheetsService;

        _nav = nav;
        _jsRuntime = jsRuntime;
        _toastService = toastService;

        Event = _mainVm.SelectedEvent is null
            ? new Event()
            : _mainVm.SelectedEvent.Clone();

        _initialEventSnapshot = Event.Clone();

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

    public async Task SaveAsync()
    {
        if (_mainVm.SelectedEvent == null || Event == null)
            return;

        if (!ValidateColumns())
            return;

        if (!Validate())
        {
            ClearLog();

            foreach (var msg in GetValidationMessages())
                AddLog(msg, null, LogLevel.Warning);

            return;
        }

        if (!await ValidateSpreadsheetAccessAsync())
            return;

        // Copy data from edited clone to original event
        _mainVm.SelectedEvent.Name = Event.Name;
        _mainVm.SelectedEvent.StartDate = Event.StartDate;
        _mainVm.SelectedEvent.SpreadsheetId = Event.SpreadsheetId;
        _mainVm.SelectedEvent.CoursesAreMerged = Event.CoursesAreMerged;
        _mainVm.SelectedEvent.Courses = Event.Courses;
        _mainVm.SelectedEvent.RawDataColumns = Event.RawDataColumns;
        _mainVm.SelectedEvent.PreprocessedColumns = Event.PreprocessedColumns;

        _mainVm.OnEventModified();
        _toastService.Show("Event saved successfully!");
        _nav.NavigateTo("/");
    }


    private async Task<bool> ValidateSpreadsheetAccessAsync()
    {
        try
        {
            var request = _sheetsService.Spreadsheets.Get(Event.SpreadsheetId);
            var spreadsheet = await request.ExecuteAsync();

            if (spreadsheet.SpreadsheetId != Event.SpreadsheetId)
            {
                AddLog("Warning: The spreadsheet ID returned by the API does not match the provided ID.", null, LogLevel.Warning);
                return false;
            }

            return true;
        }
        catch (Google.GoogleApiException ex) when (ex.HttpStatusCode == HttpStatusCode.NotFound)
        {
            AddLog("Error: Spreadsheet not found. Please check the Spreadsheet ID.", ex, LogLevel.Error);
            return false;
        }
        catch (Google.GoogleApiException ex) when (ex.HttpStatusCode == HttpStatusCode.Forbidden)
        {
            AddLog("Error: Access denied. Make sure the spreadsheet is shared with your service account.", ex, LogLevel.Error);
            return false;
        }
        catch (Exception ex)
        {
            AddLog($"Error: Unexpected issue while validating the Spreadsheet ID. Details: {ex.Message}", ex, LogLevel.Error);
            return false;
        }
    }

    public async Task CancelAsync()
    {
        if (HasUnsavedChanges)
        {
            var confirm = await _jsRuntime.InvokeAsync<bool>("confirm", "You have unsaved changes. Do you want to discard them?");
            if (!confirm)
                return;
        }

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

    private bool EventEquals(Event? a, Event? b)
    {
        if (a is null || b is null) return false;

        return a.Name == b.Name
            && a.SpreadsheetId == b.SpreadsheetId
            && a.StartDate == b.StartDate
            && a.CoursesAreMerged == b.CoursesAreMerged
            && CompareColumns(a.RawDataColumns, b.RawDataColumns)
            && CompareColumns(a.PreprocessedColumns, b.PreprocessedColumns)
            && CompareCourses(a.Courses, b.Courses);
    }

    private bool CompareColumns(ColumnsConfiguration a, ColumnsConfiguration b)
    {
        return a.DateTime == b.DateTime
            && a.Email == b.Email
            && a.FirstName == b.FirstName
            && a.LastName == b.LastName
            && a.Accepted == b.Accepted;
    }

    private bool CompareCourses(List<Course> a, List<Course> b)
    {
        if (a.Count != b.Count) return false;

        for (int i = 0; i < a.Count; i++)
        {
            if (!a[i].Equals(b[i])) return false;
        }

        return true;
    }
}
