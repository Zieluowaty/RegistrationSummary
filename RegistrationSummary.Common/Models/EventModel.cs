namespace RegistrationSummary.Common.Models;

public class EventModel
{
	public string Id { get; set; } = Guid.NewGuid().ToString();
	public string Name { get; set; } = string.Empty;
	public DateTime StartDate { get; set; }
	public DateTime EndDate { get; set; }
	public string EventType { get; set; } = string.Empty;
	public string Description { get; set; } = string.Empty;
	public List<ProductModel> Products { get; set; } = new();
}