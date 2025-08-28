using MailKit.Net.Smtp;
using MailKit.Security;
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
    private string _password;
    private readonly string _serverName;
    private readonly int _serverPort;

    private readonly ExcelService _excelService;
    private readonly List<EmailJsonModel> _emailsTemplates;

    private readonly string _testRecipientEmail;

    public MailerService(
        MailerConfiguration mailerConfiguration,
        ExcelService excelService,
        List<EmailJsonModel> emailsTemplates)
    {
        _mail = mailerConfiguration.Mail ?? string.Empty;
        _password = mailerConfiguration.Password ?? string.Empty;
        _serverName = mailerConfiguration.ServerName ?? string.Empty;
        _serverPort = mailerConfiguration.ServerPort;

        _excelService = excelService;
        _emailsTemplates = emailsTemplates;
        _testRecipientEmail = mailerConfiguration.TestMailRecepientEmailAddress ?? "test@example.com";
    }

    public bool IsConnectionSuccessful()
    {
        using var client = new SmtpClient();
        try
        {
            client.CheckCertificateRevocation = false;

            client.Connect(_serverName, _serverPort, SecureSocketOptions.SslOnConnect);
            client.Authenticate(_mail, _password);
            client.Disconnect(true);
        }
        catch (Exception ex)
        {
            return false;
        }

        return true;
    }

    public void PrepareAndSendEmailsForRegularSemesters(
        List<Student> students, 
        bool isTest, 
        Action<string>? log = null)
    {
        SendEmailsOfType(students, EmailType.Confirmation, isTest, log);
        SendEmailsOfType(students, EmailType.WaitingList, isTest, log);
        SendEmailsOfType(students, EmailType.NotEnoughPeople, isTest, log);
        SendEmailsOfType(students, EmailType.FullClass, isTest, log);
        SendEmailsOfType(students, EmailType.MissingPartner, isTest, log);
    }

    public void SendEmailsOfType(
        List<Student> students, 
        EmailType status, 
        bool isTest, 
        Action<string>? log = null)

    {
        foreach (var student in GetPendingRecipientsOfType(students, status))
        {
            var email = PrepareEmail(student, status);
            if (email == null)
                continue;

            try
            {
                SendEmail(email, isTest);
                MarkMailAsSent(student, status);
            }
            catch (Exception ex)
            {
                log?.Invoke(
                    $"Failed to send '{status}' email to {(isTest ? _testRecipientEmail : student.Email)}: {ex.Message}");
            }
        }
    }

    public List<Student> GetPendingRecipientsOfType(List<Student> students, EmailType status)
    {
        return students
            .Where(stu =>
                (stu.Courses?.Any(course => course.Status == status) ?? false) &&
                !stu.AlreadySentEmails.Any(email => email == status)
            )
            .ToList();
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
        if (string.IsNullOrEmpty(message.Subject) || string.IsNullOrEmpty(message.HtmlBody))
        {
            // TODO: DO SOMETHING.
        }

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
                    message.Cc.Add(MailboxAddress.Parse(_testRecipientEmail));

                    message.Subject = $"[TEST] {message.Subject}";
                }

                client.Send(message);
            }
            catch (Exception)
            {
                if (retries < 5)
                    SendEmail(message, isTest, ++retries);
                else
                    throw;
            }

            client.Disconnect(true);
        }
    }

    public Dictionary<EmailType, int> GetEmailCountsPerType(List<Student> students)
    {
        var summary = new Dictionary<EmailType, int>();

        foreach (var type in Enum.GetValues<EmailType>())
        {
            if (type == EmailType.All)
                continue;

            var recipients = GetPendingRecipientsOfType(students, type);
            summary[type] = recipients.Count;
        }

        return summary;
    }
}