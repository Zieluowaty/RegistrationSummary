using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using RegistrationSummary.Common.Converters;
using RegistrationSummary.Common.Enums;
using RegistrationSummary.Common.Models;
using RegistrationSummary.Common.Models.Bases;

namespace RegistrationSummary.Common.Services;

public class MailerService
{
	private readonly string _mail;
	private readonly string _password;
	private readonly string _serverName;
	private readonly int _serverPort;

	private readonly ExcelService _excelService;
	private readonly List<EmailJsonModel> _emailsTemplates;

	public bool IsConnectionSuccessful()
	{
        using (var client = new SmtpClient())
        {
            try
            {
                client.Connect(_serverName, _serverPort, SecureSocketOptions.Auto);
                client.Authenticate(_mail, _password);
                client.Disconnect(true);
            }
            catch
            {
				return false;
            }

			return true;
        }
	}

	public MailerService(string? mail, string? password, string? serverName, int serverPort, ExcelService excelService, List<EmailJsonModel> emailsTemplates)
	{
		_mail = mail ?? string.Empty;
		_password = password ?? string.Empty;
		_serverName = serverName ?? string.Empty;
		_serverPort = serverPort;
		_excelService = excelService;
		_emailsTemplates = emailsTemplates;
	}

	public void PrepareAndSendEmailsForRegularSemesters(List<Student> students)
	{
		SendEmailsOfType(students, EmailType.Confirmation);
        SendEmailsOfType(students, EmailType.WaitingList);
        SendEmailsOfType(students, EmailType.NotEnoughPeople);
        SendEmailsOfType(students, EmailType.FullClass);
        SendEmailsOfType(students, EmailType.MissingPartner);
    }

	public void SendEmailsOfType(List<Student> students, EmailType status)
	{
        var onlyChosenStudents = 
			students
				.Where(stu => 
					(stu.Courses?.Any(course => course.Status == status) ?? false) 
					&& !stu.AlreadySentEmails.Any(email => email == status)
				).ToList();

        foreach (var student in onlyChosenStudents)
        {
            var email = PrepareEmail(student, status);
            if (email == null)
                continue;

            MarkMailAsSent(student, status);
            SendEmail(email);
        }
    }

	private MimeMessage? PrepareEmail(Student student, EmailType emailType)
	{
		var emailTemplate = _emailsTemplates.FirstOrDefault(email => email.Name.Equals(emailType.ToString()));
		if (emailTemplate == null)
			return null;

		var message = new MimeMessage();
		message.From.Add(new MailboxAddress(_mail, _mail));
		message.To.Add(
			new MailboxAddress($"{student.FirstName} {student.LastName}", student.Email)
		);
		message.Cc.Add(new MailboxAddress(_mail, _mail));
		message.Subject = emailTemplate.Title;

		string mailText = string.Empty;

		switch (emailType)
		{
			case EmailType.Confirmation:
				mailText = MailerConverter.ConvertConfirmationEmail(
				emailTemplate.HeaderMerged,
				emailTemplate.PaymentInfoMerged,
				emailTemplate.CourseInfoMerged,
				emailTemplate.FooterMerged,
				student);
				break;

			case EmailType.WaitingList:
				mailText = MailerConverter.ConvertWaitingListEmail(
				emailTemplate.HeaderMerged,
				emailTemplate.PaymentInfoMerged,
				emailTemplate.CourseInfoMerged,
				emailTemplate.FooterMerged,
				student);
				break;

            case EmailType.NotEnoughPeople:
                mailText = MailerConverter.ConvertNotEnoughPeopleEmail(
                emailTemplate.HeaderMerged,
                emailTemplate.PaymentInfoMerged,
                emailTemplate.CourseInfoMerged,
                emailTemplate.FooterMerged,
                student);
                break;

            case EmailType.FullClass:
                mailText = MailerConverter.ConvertFullClassEmail(
                emailTemplate.HeaderMerged,
                emailTemplate.CourseInfoMerged,
                emailTemplate.FooterMerged,
                student);
                break;

            case EmailType.MissingPartner:
                mailText = MailerConverter.ConvertMissingPartnerEmail(
                emailTemplate.HeaderMerged,
                emailTemplate.FooterMerged,
                student);
                break;

            default:
				return null;
		}

		message.Body = new TextPart("html")
		{
			Text = mailText
		};

		return message;
	}

	private void MarkMailAsSent(Student student, EmailType type)
	{
		var spreadsheet = _excelService.SheetService.Spreadsheets.Get(_excelService.SpreadSheetId).Execute();
		var summaryTab = spreadsheet.Sheets.FirstOrDefault(sheet => sheet.Properties.Title.Equals(_excelService.SummaryTabName));

		if (summaryTab == null)
			return;

		int confirmationColumnOffset = 0;
		string confirmationText = string.Empty;

		switch (type)
		{
			case EmailType.Confirmation:
				confirmationColumnOffset = 12;
				confirmationText = "CONFIRMATION";
				break;
			case EmailType.WaitingList:
				confirmationColumnOffset = 13;
				confirmationText = "WAITING LIST";
				break;
            case EmailType.NotEnoughPeople:
                confirmationColumnOffset = 14;
                confirmationText = "NOT ENOUGH PEOPLE";
                break;
			case EmailType.FullClass:
				confirmationColumnOffset = 15;
				confirmationText = "CLASS IS FULL";
				break;
            case EmailType.MissingPartner:
                confirmationColumnOffset = 16;
                confirmationText = "MISSING PARTNER";
                break;
            default:
				return;
		}

		var confirmationColumnIndex = _excelService.GetColumnIndex(_excelService.ColumnNameForInstallmentsSum) + confirmationColumnOffset;

		_excelService.AddFormula(summaryTab.Properties.SheetId,
			$"{_excelService.SummaryTabName}!{_excelService.GetColumnName(confirmationColumnIndex)}{_excelService.ROW_STARTING_INDEX + student.Id}",
			$"=\"{DateTime.Today.ToString("yyyy-MM-dd")} {confirmationText}\"");

		_excelService.BatchUpdate();
	}

	protected virtual void SendEmail(MimeMessage message, int retries = 0)
	{
		// Send the email
		using (var client = new SmtpClient())
		{
			try
			{
                client.Connect(_serverName, _serverPort, SecureSocketOptions.Auto);
                client.Authenticate(_mail, _password);
                client.Send(message);
            }
			catch
			{
				if (retries < 5)
					SendEmail(message, ++retries);
			}

            client.Disconnect(true);
		}
	}
}