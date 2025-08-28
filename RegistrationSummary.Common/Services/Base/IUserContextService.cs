namespace RegistrationSummary.Common.Services.Base;

public class UserContextServiceOwner
{
    protected readonly UserContextService _userContext;

    public string BasePath =>
        string.IsNullOrEmpty(_userContext.Username)
            ? throw new InvalidOperationException("User not initialized.")
            : Path.Combine("C:/RegistrationSummary", _userContext.Username);

    public UserContextServiceOwner(UserContextService userContext)
    {
        _userContext = userContext;
    }
}
