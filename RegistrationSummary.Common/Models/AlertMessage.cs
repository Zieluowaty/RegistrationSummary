namespace RegistrationSummary.Common.Models;

public class AlertMessage
{
    public string Title { get; }
    public string Message { get; }
    public string Cancel { get; }

    public AlertMessage(string title, string message, string cancel)
    {
        Title = title;
        Message = message;
        Cancel = cancel;
    }
}