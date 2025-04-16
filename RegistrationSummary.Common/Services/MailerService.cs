// Services/MailerService.cs
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.Extensions.Options;
using MimeKit;
using RegistrationSummary.Common.Configurations;
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

	private readonly string _testRecipientEmail;

    public MailerService(
		IOptions<MailerConfiguration> config,
		IOptions<Settings> settings,
		ExcelService excelService,
		List<EmailJsonModel> emailsTemplates)
    {
        _mail = config.Value.Mail ?? string.Empty;
        _password = config.Value.Password ?? string.Empty;
        _serverName = config.Value.ServerName ?? string.Empty;
        _serverPort = config.Value.ServerPort;

        _excelService = excelService;
        _emailsTemplates = emailsTemplates;
        _testRecipientEmail = settings.Value.TestMailRecepientEmailAddress ?? "test@example.com";
    }

    public bool IsConnectionSuccessful()
	{
		using var client = new SmtpClient();
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

	public void PrepareAndSendEmailsForRegularSemesters(List<Student> students, bool isTest)
	{
		SendEmailsOfType(students, EmailType.Confirmation, isTest);
		SendEmailsOfType(students, EmailType.WaitingList, isTest);
		SendEmailsOfType(students, EmailType.NotEnoughPeople, isTest);
		SendEmailsOfType(students, EmailType.FullClass, isTest);
		SendEmailsOfType(students, EmailType.MissingPartner, isTest);
	}

	public void SendEmailsOfType(List<Student> students, EmailType status, bool isTest)
	{
		var onlyChosenStudents = students
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
			SendEmail(email, isTest);
		}
	}

	private MimeMessage? PrepareEmail(Student student, EmailType emailType)
	{
		var emailTemplate = _emailsTemplates.FirstOrDefault(email => email.Name.Equals(emailType.ToString()));
		if (emailTemplate == null)
			return null;

		var message = new MimeMessage();
		message.From.Add(new MailboxAddress(_mail, _mail));
		message.To.Add(new MailboxAddress($"{student.FirstName} {student.LastName}", student.Email));
		message.Cc.Add(new MailboxAddress(_mail, _mail));
		message.Subject = emailTemplate.Title;

		string mailText = emailType switch
		{
			EmailType.Confirmation => MailerConverter.ConvertConfirmationEmail(emailTemplate.HeaderMerged, emailTemplate.PaymentInfoMerged, emailTemplate.CourseInfoMerged, emailTemplate.FooterMerged, student),
			EmailType.WaitingList => MailerConverter.ConvertWaitingListEmail(emailTemplate.HeaderMerged, emailTemplate.PaymentInfoMerged, emailTemplate.CourseInfoMerged, emailTemplate.FooterMerged, student),
			EmailType.NotEnoughPeople => MailerConverter.ConvertNotEnoughPeopleEmail(emailTemplate.HeaderMerged, emailTemplate.PaymentInfoMerged, emailTemplate.CourseInfoMerged, emailTemplate.FooterMerged, student),
			EmailType.FullClass => MailerConverter.ConvertFullClassEmail(emailTemplate.HeaderMerged, emailTemplate.CourseInfoMerged, emailTemplate.FooterMerged, student),
			EmailType.MissingPartner => MailerConverter.ConvertMissingPartnerEmail(emailTemplate.HeaderMerged, emailTemplate.FooterMerged, student),
			_ => null
		} ?? string.Empty;

		message.Body = new TextPart("html") { Text = mailText };
		return message;
	}

	private void MarkMailAsSent(Student student, EmailType type)
	{
		var spreadsheet = _excelService.SheetService.Spreadsheets.Get(_excelService.SpreadSheetId).Execute();
		var summaryTab = spreadsheet.Sheets.FirstOrDefault(sheet => sheet.Properties.Title.Equals(_excelService.SummaryTabName));
		if (summaryTab == null) return;

		int offset = type switch
		{
			EmailType.Confirmation => 12,
			EmailType.WaitingList => 13,
			EmailType.NotEnoughPeople => 14,
			EmailType.FullClass => 15,
			EmailType.MissingPartner => 16,
			_ => 0
		};

		string text = type switch
		{
			EmailType.Confirmation => "CONFIRMATION",
			EmailType.WaitingList => "WAITING LIST",
			EmailType.NotEnoughPeople => "NOT ENOUGH PEOPLE",
			EmailType.FullClass => "CLASS IS FULL",
			EmailType.MissingPartner => "MISSING PARTNER",
			_ => string.Empty
		};

		var colIndex = _excelService.GetColumnIndex(_excelService.ColumnNameForInstallmentsSum) + offset;
		_excelService.AddFormula(summaryTab.Properties.SheetId,
			$"{_excelService.SummaryTabName}!{_excelService.GetColumnName(colIndex)}{_excelService.ROW_STARTING_INDEX + student.Id}",
			$"=\"{DateTime.Today:yyyy-MM-dd} {text}\"");

		_excelService.BatchUpdate();
	}

	protected virtual void SendEmail(MimeMessage message, bool isTest = false, int retries = 0)
{
    using (var client = new SmtpClient())
    {
        try
        {
            client.Connect(_serverName, _serverPort, SecureSocketOptions.Auto);
            client.Authenticate(_mail, _password);

            if (isTest)
            {
                message.To.Clear();
                message.Cc.Clear();
                message.Bcc.Clear();
                message.To.Add(MailboxAddress.Parse(_testRecipientEmail));
                message.Subject = $"[TEST] {message.Subject}";
            }

            client.Send(message);
        }
        catch
        {
            if (retries < 5)
                SendEmail(message, isTest, ++retries);
        }

        client.Disconnect(true);
    }
}

}
