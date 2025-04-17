using RegistrationSummary.Common.Enums;
using System.Xml.Linq;


namespace RegistrationSummary.Common.Models;

[Serializable]
public class Course
{
	public int Id { get; set; }
	public string Type { get; set; }
	public string Name { get; set; }

	public string Code { get; set; }
	public DateTime Start { get; set; }
	public DateTime End { get; set; }
	public string DayOfWeek { get; set; }
	public TimeSpan Time { get; set; }
	public string FormattedTime => Time.ToString(@"hh\:mm");
    public string Location { get; set; }
	public string AdditionalComment { get; set; }
	public bool IsSolo { get; set; }
    public bool IsShorter { get; set; }
    public Role? Role { get; set; }

	public EmailType? Status = null;

	public string EmailCommentary;

	public bool IsOddRow { get; set; }

    // For deserialization.
    public Course() { }

    public Course(string name, string code, DateTime start, DateTime end, string dayOfWeek, TimeSpan time, string location, string additionalComment, bool isSolo, bool isShorter, Role role)
	{
		Type = "Course";
		Name = name;
		Code = code;
		Start = start;
		End = end;
		DayOfWeek = dayOfWeek;
		Time = time;
		Location = location;
		AdditionalComment = additionalComment;
		IsSolo = isSolo;
		IsShorter = isShorter;
		Role = role;
	}

    public Course Clone()
    {
        return new Course(Name, Code, Start, End, DayOfWeek, Time, Location, AdditionalComment, IsSolo, IsShorter, Role.HasValue ? Role.Value : default(Role));
    }

    public bool Equals(Course? other)
    {
        if (other is null) return false;

        return Id == other.Id &&
               Type == other.Type &&
               Name == other.Name &&
               Code == other.Code &&
               Start == other.Start &&
               End == other.End &&
               DayOfWeek == other.DayOfWeek &&
               Time == other.Time &&
               Location == other.Location &&
               AdditionalComment == other.AdditionalComment &&
               IsSolo == other.IsSolo &&
               IsShorter == other.IsShorter &&
               Role == other.Role &&
               Status == other.Status &&
               EmailCommentary == other.EmailCommentary &&
               IsOddRow == other.IsOddRow;
    }

    public override bool Equals(object? obj) => Equals(obj as Course);

    public override int GetHashCode() =>
        HashCode.Combine(
            HashCode.Combine(Id, Type, Name, Code, Start, End, DayOfWeek, Time),
            HashCode.Combine(Location, AdditionalComment, IsSolo, IsShorter, Role, Status, EmailCommentary, IsOddRow)
    );
}