using RegistrationSummary.Common.Models;
using RegistrationSummary.Common.Services;
using System.ComponentModel;
using System.Runtime.CompilerServices;

using LogLevel = RegistrationSummary.Common.Enums.LogLevel;

namespace RegistrationSummary.Blazor.ViewModels;

public class ViewModelBase : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    private bool _isBusy;
    public bool IsBusy
    {
        get => _isBusy;
        set
        {
            if (_isBusy != value)
            {
                _isBusy = value;
                OnPropertyChanged();
            }
        }
    }

    private readonly ILogger<MainPageViewModel> _logger;
    private readonly FileLoggerService _fileLogger;
    public ViewModelBase(ILogger<MainPageViewModel> logger, FileLoggerService fileLoggerService)
    {
        _logger = logger;
        _fileLogger = fileLoggerService;
    }

    private List<LogEntry> _logLines = new();
    public IReadOnlyList<LogEntry> MessageLogLines => _logLines;

    private string _messageLog = string.Empty;
    public string MessageLog
    {
        get => _messageLog;
        set
        {
            if (_messageLog != value)
            {
                _messageLog = value;
                OnPropertyChanged();
            }
        }
    }

    protected void AddLog(string message, Exception? ex = null, LogLevel level = LogLevel.Info, string methodName = "")
    {
        var colorClass = level switch
        {
            LogLevel.Info => "log-info",
            LogLevel.Warning => "log-warning",
            LogLevel.Error => "log-error",
            _ => "log-info"
        };

        var formatted = $"[{DateTime.Now:HH:mm:ss}] {message}";

        _logLines.Insert(0, new LogEntry
        {
            Text = formatted,
            CssClass = colorClass,
            Level = level
        });

        if (level > LogLevel.Warning && ex != null)
        {
            _logger.LogError(ex, $"{level} during {methodName}");
            _fileLogger.LogError(ex, $"{level} during {methodName}");
        }

        OnPropertyChanged(nameof(MessageLogLines));
    }

    protected void ClearLog()
    {
        _logLines.Clear();
    }
}