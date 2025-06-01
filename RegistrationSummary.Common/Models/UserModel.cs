namespace RegistrationSummary.Common.Models;

public class UserModel
{
    public string Username { get; set; }
    public string PasswordHash { get; set; }
    public string Email { get; set; }
}