using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using RegistrationSummary.Blazor.Services;
using RegistrationSummary.Blazor.ViewModels;
using RegistrationSummary.Common.Services;
using System.Text.RegularExpressions;

public class LoginPageViewModel : ViewModelBase
{
    private readonly AuthenticationService _authService;
    private readonly NavigationManager _navigation;
    private readonly FileService _fileService;

    public string Username { get; set; } = "";
    public string Password { get; set; } = "";
    public string Email { get; set; } = "";
    public string NewPassword { get; set; } = "";
    public string RepeatPassword { get; set; } = "";

    public bool RequiresInitialization { get; private set; }

    public LoginPageViewModel(
        ILogger<MainPageViewModel> logger,
        FileLoggerService fileLoggerService,
        IJSRuntime jsRuntime, 
        AuthenticationService authService, 
        NavigationManager navigation, 
        ToastService toast,
        FileService fileService)
        : base(logger, fileLoggerService, jsRuntime, navigation, toast)
    {
        _authService = authService;
        _navigation = navigation;
        _fileService = fileService;
    }

    public async Task TryLoginAsync()
    {
        await RunWithBusyIndicator(async () =>
        {
            RequiresInitialization = _authService.RequiresInitialization(Username);
            if (!RequiresInitialization)
            {
                if (await _authService.LoginAsync(Username, Password))
                {
                    _fileService.UserFolderName = Username;
                    NavigateTo("/");
                }
                else
                {
                    ShowToast("Invalid login or password.");
                }
            }
        });
    }

    public async Task InitializeAccountAsync()
    {
        await RunWithBusyIndicator(async () =>
        {
            if (await _authService.InitializeAccountAsync(Username, NewPassword, RepeatPassword, Email))
            {
                ShowToast("Account initialized. Please log in.");
                ResetFields();
            }
            else
            {
                ShowToast("Failed to initialize account.");
            }
        });
    }

    private void ResetFields()
    {
        Username = "";
        Password = "";
        Email = "";
        NewPassword = "";
        RepeatPassword = "";
        RequiresInitialization = false;

        OnPropertyChanged(nameof(Username));
        OnPropertyChanged(nameof(Password));
        OnPropertyChanged(nameof(Email));
        OnPropertyChanged(nameof(NewPassword));
        OnPropertyChanged(nameof(RepeatPassword));
        OnPropertyChanged(nameof(RequiresInitialization));
    }
}