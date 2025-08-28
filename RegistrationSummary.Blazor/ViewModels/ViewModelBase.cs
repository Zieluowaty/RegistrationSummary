using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using RegistrationSummary.Blazor.Services;
using RegistrationSummary.Common.Models;
using RegistrationSummary.Common.Services;
using System.ComponentModel;
using System.Runtime.CompilerServices;

using LogLevel = RegistrationSummary.Common.Enums.LogLevel;

namespace RegistrationSummary.Blazor.ViewModels;

public class ViewModelBase : INotifyPropertyChanged
{
    // === Events ===
    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    // === Dependencies ===
    private readonly ILogger<MainPageViewModel> _logger;
    private readonly UserContextService _userContextService;
    private readonly IJSRuntime _jsRuntime;
    private readonly NavigationManager _navigationManager;
    private readonly ToastService _toastService;

    // === Constructor ===
    public ViewModelBase(
        ILogger<MainPageViewModel> logger,
        UserContextService userContextService,
        IJSRuntime jsRuntime,
        NavigationManager navigationManager,
        ToastService toastService)
    {
        _logger = logger;
        _userContextService = userContextService;
        _jsRuntime = jsRuntime;
        _navigationManager = navigationManager;
        _toastService = toastService;
    }

    // === Properties ===
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

    private List<LogEntry> _logLines = new();
    public IReadOnlyList<LogEntry> MessageLogLines => _logLines;

    public Action? OnLogUpdated { get; set; }

    // === Logging & Messaging ===
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
            _userContextService.FileLoggerService.LogError(ex, $"{level} during {methodName}");
            ShowToast($"Error occured, check the logs");
        }

        OnPropertyChanged(nameof(MessageLogLines));
        OnLogUpdated?.Invoke();
    }

    protected void ClearLog()
    {
        _logLines.Clear();
        OnPropertyChanged(nameof(MessageLogLines));
    }

    protected void ShowToast(string message)
    {
        _toastService.Show(message);
    }

    public async Task<bool> ConfirmAsync(string message)
    {
        return await _jsRuntime.InvokeAsync<bool>("confirm", message);
    }

    public void NavigateTo(string url)
    {
        _navigationManager.NavigateTo(url);
    }

    protected async Task RunWithBusyIndicator(Func<Task> action)
    {
        IsBusy = true;
        OnPropertyChanged(nameof(IsBusy)); // just in case

        await Task.Yield();

        try
        {
            await action();
        }
        finally
        {
            IsBusy = false;
            OnPropertyChanged(nameof(IsBusy));
        }
    }
}
