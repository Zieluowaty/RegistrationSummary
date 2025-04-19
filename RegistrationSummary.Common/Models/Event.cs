using RegistrationSummary.Common.Configurations;
using RegistrationSummary.Common.Enums;
using System.ComponentModel.DataAnnotations;
using System.Text.Json.Serialization;

namespace RegistrationSummary.Common.Models;

public class Event
{
	public int Id { get; set; }

    [Required(ErrorMessage = "The event name is required.")]
    [StringLength(100, MinimumLength = 10, ErrorMessage = "The event name has to have 10 to 100 characters.")]
    public string Name { get; set; }
	public DateOnly StartDate { get; set; }
    public EventType EventType { get; set; }
	public bool CoursesAreMerged { get; set; }
	[Required(ErrorMessage = "Spreadsheet ID needs to be provided.")]
	public string SpreadsheetId { get; set; }
	public ColumnsConfiguration RawDataColumns { get; set; }
	public ColumnsConfiguration PreprocessedColumns { get; set; }

	[JsonPropertyName("Products")]
	public List<Course> Courses { get; set; }

	public Event() { }

	public Event(int id, string name, DateOnly? startDate, EventType eventType, bool coursesAreMerged, string spreadSheetId, ColumnsConfiguration rawDataColumns, ColumnsConfiguration preprocessedColumns, List<Course> courses)
	{
		Id = id;
		Name = name;
		StartDate = startDate ?? DateOnly.FromDateTime(DateTime.Now);
		EventType = eventType;
		CoursesAreMerged = coursesAreMerged;
		SpreadsheetId = spreadSheetId;
		RawDataColumns = rawDataColumns;
		PreprocessedColumns = preprocessedColumns;
		Courses = courses;

		for (int i = 0; i < Courses.Count; i++)
		{
			Courses[i].IsOddRow = i % 2 != 0;
		}
	}

	public Event Clone()
	{
		var clonedCourses = Courses?.Select(c => c.Clone()).ToList() ?? new List<Course>();
		var clonedRawDataColumns = RawDataColumns.Clone();
		var clonedPreprocessedColumns = PreprocessedColumns.Clone();

		return new Event(0, Name, StartDate, EventType, CoursesAreMerged, SpreadsheetId, clonedRawDataColumns, clonedPreprocessedColumns, clonedCourses);
	}
    public bool Equals(Event? other)
    {
        if (other is null) return false;

        return Name == other.Name &&
               SpreadsheetId == other.SpreadsheetId &&
               StartDate == other.StartDate &&
               CoursesAreMerged == other.CoursesAreMerged &&
               RawDataColumns.Equals(other.RawDataColumns) &&
               PreprocessedColumns.Equals(other.PreprocessedColumns) &&
               Courses.SequenceEqual(other.Courses);
    }

    public override bool Equals(object? obj) => Equals(obj as Event);

    public override int GetHashCode() =>
        HashCode.Combine(Name, SpreadsheetId, StartDate, CoursesAreMerged,
                         RawDataColumns, PreprocessedColumns,
                         Courses.Aggregate(0, (a, c) => HashCode.Combine(a, c.GetHashCode())));
}