namespace RegistrationSummary.Common.Models;

public class ProductModel
{
	public string Code { get; set; } = string.Empty;
	public string Name { get; set; } = string.Empty;
	public DateTime Start { get; set; }
	public string DayOfWeek { get; set; } = string.Empty;
	public TimeSpan Time { get; set; }
	public bool IsSolo { get; set; }
	public string Location { get; set; } = string.Empty;
}