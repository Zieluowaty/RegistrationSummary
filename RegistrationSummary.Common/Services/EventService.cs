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

    public int GenerateNextId()
    {
        var events = GetAll();
        return events.Any() ? events.Max(e => e.Id) + 1 : 1;
    }

    public Event? GetByName(string name)
		=> GetAll().FirstOrDefault(e => e.Name == name);

    public Event? GetById(int id)
    {
        return GetAll().FirstOrDefault(e => e.Id == id);
    }

    public void SaveAll(List<Event> events)
		=> _fileService.Save(FileName, events);
}