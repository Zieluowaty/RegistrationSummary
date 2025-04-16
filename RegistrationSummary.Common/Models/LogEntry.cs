using RegistrationSummary.Common.Enums;

namespace RegistrationSummary.Common.Models;

public class LogEntry
{
    public string Text { get; set; } = string.Empty;
    public string CssClass { get; set; } = string.Empty;
    public LogLevel Level { get; set; }
}