using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Newtonsoft.Json.Linq;
using RegistrationSummary.Common.Configurations;
using RegistrationSummary.Common.Enums;
using RegistrationSummary.Common.Models;
using RegistrationSummary.Common.Models.Interfaces;
using RegistrationSummary.Maui.Services;
using System.Collections.ObjectModel;
using System.ComponentModel.DataAnnotations;

namespace RegistrationSummary.Maui.ViewModels;

[QueryProperty(nameof(DialogService), "DialogService")]
[QueryProperty(nameof(SelectedEvent), "SelectedEvent")]
public partial class EventModificationViewModel : ObservableValidator
{
	[ObservableProperty]
    private DialogService _dialogService;

	[ObservableProperty]
    private Event _selectedEvent;

	[Required]
    [StringLength(100, ErrorMessage = "{0} length must be between {2} and {1}.", MinimumLength = 10)]
    [ObservableProperty]
	private string _name;

    [Required]
    [ObservableProperty]
    private DateTime _startDate;

    [ObservableProperty]
	private bool _coursesAreMerged;

    [ObservableProperty]
    [StringLength(100, ErrorMessage = "{0} length must be between {2} and {1}.", MinimumLength = 10)]
    private string _spreadSheetId;

	[ObservableProperty]
	private ColumnsConfiguration _rawDataColumns;

	[ObservableProperty]
	private ColumnsConfiguration _preprocessedColumns;

	[ObservableProperty]
	private ObservableCollection<IProduct> _courses;

    private bool _newEvent;

    public void RefreshData()
    {
        var eventForEdit = SelectedEvent;

        _newEvent = eventForEdit == null;

        if (eventForEdit != null)
        {
            Name = eventForEdit.Name;
            StartDate = eventForEdit.StartDate;
            CoursesAreMerged = eventForEdit.CoursesAreMerged;
            SpreadSheetId = eventForEdit.SpreadSheetId;
            RawDataColumns = eventForEdit.RawDataColumns.Clone();
            PreprocessedColumns = eventForEdit.PreprocessedColumns.Clone();
            Courses = new ObservableCollection<IProduct>(eventForEdit.Products.ToList());
        }
        else
        {
            Name = "Provide name for your event";
            StartDate = DateTime.Today;
            CoursesAreMerged = false;
            SpreadSheetId = "Your google sheet spreadsheet ID";
            RawDataColumns = new ColumnsConfiguration
            {
                DateTime = string.Empty
            };
            PreprocessedColumns = new ColumnsConfiguration
            {
                DateTime = string.Empty,
                Login = "A",
                Email = "B",
                FirstName = "C",
                LastName = "D",
                PhoneNumber = "E",
                Course = "F",
                Role = "G",
                Partner = "H",
                Installment = "I",
                Accepted = "J"
            };
            Courses = new ObservableCollection<IProduct>();
        }
    }

    [RelayCommand]
    private async Task Save()
    {
        ValidateAllProperties();
		RawDataColumns.Validate();
		PreprocessedColumns.Validate();

        if (HasErrors || RawDataColumns.HasErrors || PreprocessedColumns.HasErrors)
		{
			var commonErrorsText = HasErrors ? string.Join(Environment.NewLine, GetErrors().Select(err => err.ErrorMessage)) : string.Empty;
            var rawDataErrorsText = RawDataColumns.HasErrors ? $"Raw data columns:{Environment.NewLine}{string.Join(Environment.NewLine, RawDataColumns.GetErrors().Select(err => err.ErrorMessage))}" : string.Empty;
            var preprocessedErrorsText = PreprocessedColumns.HasErrors ? $"Preprocessed data columns:{Environment.NewLine}{string.Join(Environment.NewLine, PreprocessedColumns.GetErrors().Select(err => err.ErrorMessage))}" : string.Empty;

            var errorText = string.Join(Environment.NewLine, new string[] {commonErrorsText, rawDataErrorsText, preprocessedErrorsText});

            await DialogService.ShowAlertAsync("Missing data", 
				"Some of the data are missing - check all fields before saving." + Environment.NewLine + errorText);
			return;
		}

		EventsService.SaveEvent(new Event(_newEvent ? 0 : SelectedEvent.Id, Name, StartDate, EventType.RegularSemester, CoursesAreMerged, SpreadSheetId, RawDataColumns, PreprocessedColumns, Courses.ToList()));

        // Go back.
        await Shell.Current.Navigation.PopAsync();
    }

	[RelayCommand]
	private async Task Cancel()
	{
		var result = await DialogService.ShowConfirmationAsync("Cancellation", "All changes (if any done) will be lost, are you sure?");

		if (result)
			await Shell.Current.Navigation.PopAsync();
	}

	[RelayCommand]
	private void AddNewCourse()
	{		
		Courses.Insert(0, 
            new Course { 
                Id = Courses.Any() ? Courses.Max(course => ((Course)course).Id) + 1 : 1, 
                Type = "Course", 
                Start = DateTime.Today, 
                End = DateTime.Today}
            );
	}

	[RelayCommand]
	private async Task DeleteCourse(string courseName)
    {
        var result = await DialogService.ShowConfirmationAsync("Removing course", $"Are you sure you want to delete '{courseName}'?");

		if (!result)
			return;

            var courseToRemove = Courses.Single(product => product.Name.Equals(courseName));
		Courses.Remove(courseToRemove);

		await DialogService.ShowAlertAsync("Course deleted", $"Course '{courseName}' deleted successfully.");		      
    }
}