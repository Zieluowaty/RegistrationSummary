using RegistrationSummary.Common.Models;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace RegistrationSummary.Blazor.Services;

public class AuthenticationService
{
    private readonly string _usersPath = Path.Combine("C:/RegistrationSummary", "Users.json");
    private List <UserModel> _users = new();

    public string? LoggedInUser { get; private set; }
    public bool IsAuthenticated => !string.IsNullOrWhiteSpace(LoggedInUser);

    public AuthenticationService()
    {
        LoadUsers();
    }

    private void LoadUsers()
    {
        if (File.Exists(_usersPath))
        {
            var json = File.ReadAllText(_usersPath);
            _users = JsonSerializer.Deserialize<List<UserModel>>(json) ?? new();
        }
    }

    private void SaveUsers()
    {
        var json = JsonSerializer.Serialize(_users);
        File.WriteAllText(_usersPath, json);
    }

    public async Task<bool> LoginAsync(string username, string password)
    {
        var user = _users.SingleOrDefault(user => user.Username == username);

        if (user == null)
            return false;

        string hashedPassword = HashPassword(password);

        if (user.PasswordHash == hashedPassword)
        {
            LoggedInUser = username;
            return true;
        }

        return false;
    }

    private string HashPassword(string password)
    {
        using var sha256 = SHA256.Create();
        byte[] bytes = sha256.ComputeHash(System.Text.Encoding.UTF8.GetBytes(password));
        return Convert.ToBase64String(bytes);
    }

    public bool RequiresInitialization(string username)
        => string.IsNullOrWhiteSpace(_users.SingleOrDefault(user => user.Username == username)?.PasswordHash);

    public async Task<bool> InitializeAccountAsync(string username, string password, string confirmPassword, string email)
    {
        var user = _users.SingleOrDefault(user => user.Username == username);

        if (user == null) 
            return false;
        
        if (string.IsNullOrWhiteSpace(password) || password != confirmPassword) 
            return false;

        if (!IsEmailValid(email))
            return false;

        user.PasswordHash = HashPassword(password);
        user.Email = email;

        await Task.Run(SaveUsers);

        return true;
    }

    private bool IsEmailValid(string email)
        => Regex.IsMatch(email,
            @"^[^@\s]+@[^@\s]+\.[^@\s]+$",
            RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250));
}