namespace RegistrationSummary.Maui;

public partial class App : Application
{
    public App()
    {
        InitializeComponent();
        MainPage = new AppShell();
    }

    protected override async void OnStart()
    {
        base.OnStart();
        await Shell.Current.GoToAsync("//mainpage");
    }
}
