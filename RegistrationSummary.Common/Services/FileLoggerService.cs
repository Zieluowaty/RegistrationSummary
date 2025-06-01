using Microsoft.Extensions.Options;
using RegistrationSummary.Common.Configurations;
using System.Text;

namespace RegistrationSummary.Common.Services;

public class FileLoggerService
{
    private readonly string _logDirectory;
    private readonly int _retentionDays;

    public FileLoggerService(IOptions<SettingConfiguration> settingsOptions, FileService fileService)
    {
        var settings = settingsOptions.Value;
        _logDirectory = Path.Combine(fileService.BasePath, "Logs");
        _retentionDays = settings.LogRetentionDays;

        if (!Directory.Exists(_logDirectory))
            Directory.CreateDirectory(_logDirectory);

        CleanupOldLogs();
    }

    public void Log(string message)
    {
        var logFilePath = GetLogFilePath();
        var fullMessage = $"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}";

        File.AppendAllText(logFilePath, fullMessage, Encoding.UTF8);
    }

    public void LogError(Exception ex, string context = "")
    {
        var logFilePath = GetLogFilePath();
        var fullMessage = new StringBuilder();
        fullMessage.AppendLine($"[{DateTime.Now:HH:mm:ss}] ERROR: {context}");
        fullMessage.AppendLine($"Exception: {ex.Message}");
        fullMessage.AppendLine(ex.StackTrace ?? "");
        fullMessage.AppendLine();

        File.AppendAllText(logFilePath, fullMessage.ToString(), Encoding.UTF8);
    }

    private string GetLogFilePath()
    {
        var fileName = $"log-{DateTime.Today:yyyy-MM-dd}.log";
        return Path.Combine(_logDirectory, fileName);
    }

    private void CleanupOldLogs()
    {
        try
        {
            var files = Directory.GetFiles(_logDirectory, "log-*.log");
            foreach (var file in files)
            {
                var fileInfo = new FileInfo(file);
                if (fileInfo.CreationTime < DateTime.Now.AddDays(-_retentionDays))
                {
                    fileInfo.Delete();
                }
            }
        }
        catch (Exception ex)
        {
            LogError(ex, "The system could not remove old log files.");
        }
    }
}
