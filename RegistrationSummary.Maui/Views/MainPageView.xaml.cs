using RegistrationSummary.Maui.ViewModels;

namespace RegistrationSummary.Maui.Views;

public partial class MainPageView : ContentPage
{
    public MainPageView()
    {
        InitializeComponent();
    }

    protected override void OnAppearing()
    {
        base.OnAppearing();

        if (BindingContext is MainPageViewModel viewModel)
        {
            viewModel.RefreshData();
        }
    }
}