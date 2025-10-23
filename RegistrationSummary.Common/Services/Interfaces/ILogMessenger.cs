using RegistrationSummary.Common.Enums;

namespace RegistrationSummary.Common.Services.Interfaces;
public interface ILogMessenger
{
    public void LogMessage(string message, Exception? ex = null, LogLevel level = LogLevel.Info, string methodName = "");
}
