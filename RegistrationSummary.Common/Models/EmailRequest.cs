namespace RegistrationSummary.Common.Models;

public class EmailRequest
{
	public string Type { get; set; } = string.Empty;
	public bool IsTest { get; set; }
}