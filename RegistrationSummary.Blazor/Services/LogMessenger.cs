using RegistrationSummary.Blazor.ViewModels;
using RegistrationSummary.Common.Services.Interfaces;

using LogLevel = RegistrationSummary.Common.Enums.LogLevel;

namespace RegistrationSummary.Blazor.Services;

public class LogMessenger : ILogMessenger
{
    private readonly MainPageViewModel _mainPageViewModel;

    // === Constructor ===
    public LogMessenger(MainPageViewModel mainPageViewModel)
    {
        _mainPageViewModel = mainPageViewModel;
    }

    // === Logging & Messaging ===
    public void LogMessage(string message, Exception? ex = null, LogLevel level = LogLevel.Info, string methodName = "")
    {
        _mainPageViewModel.AddLog(message, ex, level, methodName);
    }
}
