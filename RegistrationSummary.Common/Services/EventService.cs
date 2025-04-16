using RegistrationSummary.Common.Models;

public class EventService
{
	private readonly FileService _fileService;
	private const string FileName = "Events.json";

	public EventService(FileService fileService)
	{
		_fileService = fileService;
	}

	public List<Event> GetAll()
		=> _fileService.Load<List<Event>>(FileName);

	public Event? GetByName(string name)
		=> GetAll().FirstOrDefault(e => e.Name == name);

	public void SaveAll(List<Event> events)
		=> _fileService.Save(FileName, events);
}