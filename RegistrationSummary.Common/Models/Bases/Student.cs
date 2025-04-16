using RegistrationSummary.Common.Enums;

namespace RegistrationSummary.Common.Models.Bases;

public class Student
{
	public int Id { get; set; }
	public string Email { get; set; } = string.Empty;
	public string FirstName { get; set; } = string.Empty;
	public string LastName { get; set; } = string.Empty;

	public string Login => $"{Email},{FirstName},{LastName}";
	public int PaymentAmount { get; set; }
	public bool Installments { get; set; }
	public List<Course>? Courses { get; set; }

	public List<EmailType> AlreadySentEmails { get; set; } = new List<EmailType>();

	public string MergedCoursesNames(EmailType status)
		=>
		Courses == null
		? string.Empty
		:
		Courses
			.Where(course =>
				course.Status == status)
			.Aggregate("", (current, course) => $"{current} {course.Name}, ")
			.TrimEnd(new char[] { ',', ' ' })
		;

	public string MergedCoursesEmailCommentary(EmailType status)
		=>
		Courses == null
		? string.Empty
		:
		"<b><br><br>" +
        Courses
            .Where(course =>
                course.Status == status)
            .Aggregate("", (current, course) => $"{current} {course.EmailCommentary}{Environment.NewLine}")
            .TrimEnd() +
        "</b>"
        ;
}