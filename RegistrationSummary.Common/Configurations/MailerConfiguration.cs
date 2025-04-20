namespace RegistrationSummary.Common.Configurations;

public class MailerConfiguration
{
    public string? Mail { get; set; }
    public string? Password { get; set; }
    public string? ServerName { get; set; }
    public int ServerPort { get; set; }

    public string TestMailRecepientEmailAddress { get; set; }

    public MailerConfiguration() { }
}