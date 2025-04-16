using System.Text.RegularExpressions;
using RegistrationSummary.Common.Enums;
using RegistrationSummary.Common.Models;
using RegistrationSummary.Common.Models.Bases;

namespace RegistrationSummary.Common.Converters;

public class MailerConverter
{
	public static string ConvertConfirmationEmail(string rawEmailHeader, string rawEmailPaymentInfo, string rawEmailCourseInfo, string rawEmailFooter,
		Student student)
	{
		var output =
			rawEmailHeader
				.Replace("<STUDENT_FIRST_NAME>", student.FirstName)
				.Replace("<COURSE_LIST>", student.MergedCoursesNames(EmailType.Confirmation))
			+ rawEmailPaymentInfo
				.Replace("<STUDENT_FIRST_NAME>", student.FirstName)
				.Replace("<STUDENT_LAST_NAME>", student.LastName)
				.Replace("<COURSE_LIST>", student.MergedCoursesNames(EmailType.Confirmation))
				.Replace("<PAYMENT_AMOUNT>", student.Installments ? (student.PaymentAmount / 2).ToString() : student.PaymentAmount.ToString())
				.Replace("<INSTALLMENT_AMOUNT>", (student.PaymentAmount / 2).ToString())
			+ MultipleCourses(student.Courses, rawEmailCourseInfo, EmailType.Confirmation)
			+ rawEmailFooter;

        output = ManageInstallmentTag(student, output);
        output = ManageCoupleTag(student, output);
        output = ManageCommentaryTag(student, output, EmailType.Confirmation);

        return output;
	}

	public static string ConvertWaitingListEmail(string rawEmailHeader, string rawEmailPaymentInfo, string rawEmailCourseInfo, string rawEmailFooter,
		Student student)
	{
		var output =
			rawEmailHeader
				.Replace("<STUDENT_FIRST_NAME>", student.FirstName)
				.Replace("<COURSE_LIST>", student.MergedCoursesNames(EmailType.WaitingList))
			+ MultipleCourses(student.Courses, rawEmailCourseInfo, EmailType.WaitingList)
			+ rawEmailFooter;

        output = ManageCommentaryTag(student, output, EmailType.WaitingList);

        return output;
	}

    public static string ConvertNotEnoughPeopleEmail(string rawEmailHeader, string rawEmailPaymentInfo, string rawEmailCourseInfo, string rawEmailFooter,
        Student student)
    {
        var output =
            rawEmailHeader
                .Replace("<STUDENT_FIRST_NAME>", student.FirstName)
                .Replace("<COURSE_LIST>", student.MergedCoursesNames(EmailType.NotEnoughPeople))
            + MultipleCourses(student.Courses, rawEmailCourseInfo, EmailType.NotEnoughPeople)
            + rawEmailFooter;

        output = ManageCommentaryTag(student, output, EmailType.NotEnoughPeople);

        return output;
    }

    public static string ConvertFullClassEmail(string rawEmailHeader, string rawEmailCourseInfo, string rawEmailFooter,
        Student student)
    {
        var output =
            rawEmailHeader
                .Replace("<STUDENT_FIRST_NAME>", student.FirstName)
                .Replace("<COURSE_LIST>", student.MergedCoursesNames(EmailType.FullClass))
            + MultipleCourses(student.Courses, rawEmailCourseInfo, EmailType.FullClass)
            + rawEmailFooter;

        output = ManageCommentaryTag(student, output, EmailType.FullClass);

        return output;
    }

    public static string ConvertMissingPartnerEmail(string rawEmailHeader, string rawEmailFooter,
        Student student)
    {
        var output =
            rawEmailHeader
                .Replace("<STUDENT_FIRST_NAME>", student.FirstName)
                .Replace("<COURSE_LIST>", student.MergedCoursesNames(EmailType.MissingPartner))
            + rawEmailFooter;

        output = ManageCommentaryTag(student, output, EmailType.MissingPartner);

        return output;
    }

	private static string ManageInstallmentTag(Student student, string output)
	{
        if (student.Installments)
        {
            // Remowe only <INSTALLMENT> marks if student want to pay in installments.
            output =
                output
                .Replace("<INSTALLMENT>", string.Empty)
                .Replace("</INSTALLMENT>", string.Empty);
        }
        else
        {
            // Remove all texts inside of <INSTALLMENT> marks if student wants to pay normal.
            string pattern = @"<INSTALLMENT>.*?</INSTALLMENT>";
            output = Regex.Replace(output, pattern, string.Empty);
        }

        return output;
    }

    private static string ManageCoupleTag(Student student, string output)
    {
        if (student.Courses?.Any(course => !course.IsSolo) ?? false)
        {
            // Remowe only <INSTALLMENT> marks if student want to pay in installments.
            output =
                output
                .Replace("<COUPLE>", string.Empty)
                .Replace("</COUPLE>", string.Empty);
        }
        else
        {
            // Remove all texts inside of <INSTALLMENT> marks if student wants to pay normal.
            string pattern = @"<COUPLE>.*?</COUPLE>";
            output = Regex.Replace(output, pattern, string.Empty);
        }

        return output;
    }

    private static string ManageCommentaryTag(Student student, string output, EmailType emailType)
    {
        if (student.Courses?.Any(course => string.IsNullOrEmpty(course.EmailCommentary)) ?? false)
        {
            // Remowe only <INSTALLMENT> marks if student want to pay in installments.
            output =
                output
                .Replace("<COMMENTARY/>", string.Empty);
        }
        else
        {
            // Remove all texts inside of <INSTALLMENT> marks if student wants to pay normal.
            string pattern = @"<COMMENTARY/>";
            output = Regex.Replace(output, pattern, student.MergedCoursesEmailCommentary(emailType));
        }

        return output;
    }

    private static string MultipleCourses (List<Course>? courses, string rawEmailCourseInfo, EmailType status)
		=> courses?
			.Where(course => course.Status == status)
			.Aggregate(
			"",
			(current, course) => current +
			$"{rawEmailCourseInfo
				.Replace("<COURSE_NAME>", course.Name)
				.Replace("<COURSE_START>", course.Start.ToString("dd/MM"))
                .Replace("<COURSE_END>", course.End.ToString("dd/MM"))
                .Replace("<ROLE>", course.Role.ToString())
                .Replace("<COURSE_LOCATION>", course.Location)
				.Replace("<COURSE_DAY_OF_WEEK>", ConvertDayOfWeekToPolishCulture(course.DayOfWeek))
				.Replace("<COURSE_TIME>", course.Time.ToString(@"hh\:mm"))
                .Replace("<ADDITIONAL_COMMENT>", (!string.IsNullOrEmpty(course.AdditionalComment) ? $"{course.AdditionalComment}</br>" : string.Empty))}"
		) 
        ?? string.Empty;

	private static string ConvertDayOfWeekToPolishCulture(string dayOfWeek)
	{

        switch (dayOfWeek)
        {
            case "Monday":
                return "Poniedziałek";
            case "Tuesday":
                return "Wtorek";
            case "Wendesday":
                return "Środa";
            case "Thursday":
                return "Czwartek";
            case "Friday":
                return "Piątek";
            case "Saturday":
                return "Sobota";
            case "Sunday":
                return "Niedziela";
            default:
                return "Błąd";
        }
	}
}
