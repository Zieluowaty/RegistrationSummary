using Newtonsoft.Json;
using RegistrationSummary.Common.Models;
using RegistrationSummary.Maui.Converters;
using System.Text;

namespace RegistrationSummary.Maui.Services;

public static class EventsService
{
    public static List<Event> LoadEvents()
    {
        if (string.IsNullOrEmpty(MauiProgram.EventsDataFilePath))
        {
            throw new Exception("There is no events data file path provided in appsettings.json.");
        }

        string jsonData = File.ReadAllText(MauiProgram.EventsDataFilePath, Encoding.UTF8);

        var events = JsonConvert.DeserializeObject<List<Event>>(jsonData, new ProductConverter());

        if (events == null)
        {
            throw new Exception("Cannot load events data file.");
        }

        return events;
    }

    public static void SaveEvents(List<Event> events)
    {
        if (string.IsNullOrEmpty(MauiProgram.EventsDataFilePath))
        {
            throw new Exception("There is no events data file path provided in appsettings.json.");
        }

        var jsonString = JsonConvert.SerializeObject(events);

        File.WriteAllText(MauiProgram.EventsDataFilePath, jsonString);
    }

    public static void SaveEvent(Event eventToModify)
    {
        if (string.IsNullOrEmpty(MauiProgram.EventsDataFilePath))
        {
            throw new Exception("There is no events data file path provided in appsettings.json.");
        }

        var events = LoadEvents();

        if (events.Any(ev => ev.Id == eventToModify.Id))
        {
            events.Remove(events.Single(ev => ev.Id == eventToModify.Id));
        }

        if (eventToModify.Id == 0)
        {
            eventToModify.Id = events.Max(ev => ev.Id) + 1;
        }

        events.Add(eventToModify);

        SaveEvents(events);
    }
}