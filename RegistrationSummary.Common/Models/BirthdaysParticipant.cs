using RegistrationSummary.Common.Models.Bases;

namespace RegistrationSummary.Common.Models;

public class BirthdaysParticipant : Student
{
	public bool English { get; set; }
	public bool EarlyPass { get; set; }
	public string ChosenLevel { get; set; }

	public List<string> Parties { get; set; } = new List<string>();
}