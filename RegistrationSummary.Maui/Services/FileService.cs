using Newtonsoft.Json;
using RegistrationSummary.Common.Models;
using System.Reflection;
using System.Text;

namespace RegistrationSummary.Maui.Services;

public class FileService
{
    public static string AppFolderPath = "C:\\RegistrationSummary\\";

    private static string ReadEmbeddedResource(string resourceName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        string resourceContent;
        using (var stream = assembly.GetManifestResourceStream(resourceName))
        {
            if (stream == null) throw new InvalidOperationException("Could not find embedded resource");
            using (var reader = new StreamReader(stream))
            {
                resourceContent = reader.ReadToEnd();
            }
        }
        return resourceContent;
    }

    public static void CopyTemplateFilesIfDontExist()
    {
        List<string> templateFiles = new List<string>
        {
            "Credentials.json",
            "Emails.json",
            "Events.json",
            "Settings.json"
        };

        foreach(var fileName in templateFiles) 
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = $"{assembly.GetName().Name}.Templates.{fileName}";
            var templateContent = ReadEmbeddedResource(resourceName);

            var fullPath = Path.Combine(AppFolderPath, fileName);

            Directory.CreateDirectory(AppFolderPath);

            if (!File.Exists(fullPath))
            {
                File.WriteAllText(fullPath, templateContent);
            }
        }
    }

    public static Settings GetSettings()
    {
        string jsonData = File.ReadAllText(AppFolderPath + "Settings.json", Encoding.UTF8);
        return JsonConvert.DeserializeObject<Settings>(jsonData);
    }
}