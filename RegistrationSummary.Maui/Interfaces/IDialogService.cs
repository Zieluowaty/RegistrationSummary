namespace RegistrationSummary.Maui.Interfaces;

public interface IDialogService
{
    Task ShowMessage(string title, string message);
    Task<bool> ShowConfirmation(string title, string message);
}
