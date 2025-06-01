using Microsoft.Extensions.Configuration;
using System.Reflection;
using System.Text.Json;

public class FileService
{
	public string UserFolderName { get; set; }
	public string BasePath =>
		!string.IsNullOrEmpty(UserFolderName)
			? Path.Combine("C:/RegistrationSummary", UserFolderName)
			: " "; 

	public FileService(IConfiguration configuration)
	{ }

	public T Load<T>(string filename)
	{
		var fullPath = Path.Combine(BasePath, filename);

		if (!File.Exists(fullPath))
			throw new FileNotFoundException($"Plik '{filename}' nie istnieje w ścieżce {BasePath}");

		var json = File.ReadAllText(fullPath);
		return JsonSerializer.Deserialize<T>(json)!;
	}

	public void Save<T>(string filename, T data)
	{
		var fullPath = Path.Combine(BasePath, filename);
		var json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
		File.WriteAllText(fullPath, json);
	}

	private string ReadEmbeddedResource(string resourceName)
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

	public void CopyTemplateFilesIfDontExist()
	{
		// TODO: Remove obsolete function.

		List<string> templateFiles = new List<string>
		{
			"Credentials.json",
			"Emails.json",
			"Events.json",
			"Settings.json"
		};

		foreach (var fileName in templateFiles)
		{
			var assembly = Assembly.GetExecutingAssembly();
			var resourceName = $"{assembly.GetName().Name}.Templates.{fileName}";
			var templateContent = ReadEmbeddedResource(resourceName);

			var fullPath = Path.Combine(BasePath, fileName);

			Directory.CreateDirectory(BasePath);

			if (!File.Exists(fullPath))
			{
				File.WriteAllText(fullPath, templateContent);
			}
		}
	}
}