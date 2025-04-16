using RegistrationSummary.Common.Configurations;
using RegistrationSummary.Common.Enums;
using System.Text.Json.Serialization;

namespace RegistrationSummary.Common.Models;

public class Event
{
	public int Id { get; set; }
	public string Name { get; set; }
	public DateTime StartDate { get; set; }
	public EventType EventType { get; set; }
	public bool CoursesAreMerged { get; set; }
	public string SpreadSheetId { get; set; }
	public ColumnsConfiguration RawDataColumns { get; set; }
	public ColumnsConfiguration PreprocessedColumns { get; set; }

	[JsonPropertyName("Products")]
	public List<Course> Courses { get; set; }

	public Event() { }

	public Event(int id, string name, DateTime? startDate, EventType eventType, bool coursesAreMerged, string spreadSheetId, ColumnsConfiguration rawDataColumns, ColumnsConfiguration preprocessedColumns, List<Course> courses)
	{
		Id = id;
		Name = name;
		StartDate = startDate ?? DateTime.Now;
		EventType = eventType;
		CoursesAreMerged = coursesAreMerged;
		SpreadSheetId = spreadSheetId;
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

		return new Event(0, Name, StartDate, EventType, CoursesAreMerged, SpreadSheetId, clonedRawDataColumns, clonedPreprocessedColumns, clonedCourses);
	}
}
