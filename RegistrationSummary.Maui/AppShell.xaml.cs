using RegistrationSummary.Maui.Views;

namespace RegistrationSummary.Maui;

public partial class AppShell : Shell
{
    public AppShell()
    {
        InitializeComponent();

        Routing.RegisterRoute("mainpage/eventmodification", typeof(EventModificationView));
    }
}