using CommunityToolkit.Mvvm.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace RegistrationSummary.Common.Configurations;

public class ColumnsConfiguration : ObservableValidator
{
    public string DateTime { get; set; }
    [Required]
    [StringLength(2, ErrorMessage = "{0} length must be between {2} and {1}.", MinimumLength = 1)]
    public string Email { get; set; }
    [Required]
    [StringLength(2, ErrorMessage = "{0} length must be between {2} and {1}.", MinimumLength = 1)]
    public string FirstName { get; set; }
    [Required]
    [StringLength(2, ErrorMessage = "{0} length must be between {2} and {1}.", MinimumLength = 1)]
    public string LastName { get; set; }
    [Required]
    [StringLength(2, ErrorMessage = "{0} length must be between {2} and {1}.", MinimumLength = 1)]
    public string PhoneNumber { get; set; }
    [Required]
    [StringLength(2, ErrorMessage = "{0} length must be between {2} and {1}.", MinimumLength = 1)]
    public string Course { get; set; }
    [Required]
    [StringLength(2, ErrorMessage = "{0} length must be between {2} and {1}.", MinimumLength = 1)]
    public string Role { get; set; }
    [Required]
    [StringLength(2, ErrorMessage = "{0} length must be between {2} and {1}.", MinimumLength = 1)]
    public string Partner { get; set; }
    [Required]
    [StringLength(2, ErrorMessage = "{0} length must be between {2} and {1}.", MinimumLength = 1)]
    public string Installment { get; set; }

	public string Login { get; set; }
	public string Accepted { get; set; }

    public ColumnsConfiguration Clone()
    {
        return new ColumnsConfiguration
        {
            DateTime = DateTime,
            Email = Email,
            FirstName = FirstName,
            LastName = LastName,
            PhoneNumber = PhoneNumber,
            Course = Course,
            Role = Role,
            Partner = Partner,
            Installment = Installment,
            Login = Login,
            Accepted = Accepted
        };
    }

    public void Validate() => ValidateAllProperties();

    public bool Equals(ColumnsConfiguration? other)
    {
        if (other is null) return false;

        return DateTime == other.DateTime &&
                FirstName == other.FirstName &&
                LastName == other.LastName &&
                Email == other.Email &&
                Accepted == other.Accepted;
    }

    public override bool Equals(object? obj) => Equals(obj as ColumnsConfiguration);

    public override int GetHashCode() =>
        HashCode.Combine(DateTime, FirstName, LastName, Email, Accepted);
}