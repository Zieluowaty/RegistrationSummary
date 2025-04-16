using RegistrationSummary.Maui.ViewModels;

namespace RegistrationSummary.Maui.Views;

public partial class EventModificationView : ContentPage
{
    public EventModificationView()
    {
        InitializeComponent();
    }

    protected override void OnAppearing()
    {
        base.OnAppearing();

        if (BindingContext is EventModificationViewModel viewModel)
        {
            viewModel.RefreshData();
        }
    }
}
