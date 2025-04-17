namespace RegistrationSummary.Blazor.Services;

/// <summary>
/// Service for displaying toast notifications.
/// </summary>

public class ToastService
{
    private readonly ILogger<ToastService> _logger;

    public string? CurrentMessage { get; private set; }
    public bool Visible { get; private set; }

    public event Action? OnChange;

    public ToastService(ILogger<ToastService> logger)
    {
        _logger = logger;
    }

    public void Show(string message, int duration = 3000)
    {
        _logger.LogInformation($"Toast: {message}");
        _ = ShowInternalAsync(message, duration);
    }

    private async Task ShowInternalAsync(string message, int duration)
    {
        CurrentMessage = message;
        Visible = true;
        OnChange?.Invoke();

        await Task.Delay(duration);

        Visible = false;
        CurrentMessage = null;
        OnChange?.Invoke();
    }
}
