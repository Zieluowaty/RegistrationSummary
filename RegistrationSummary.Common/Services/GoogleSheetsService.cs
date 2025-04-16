using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;

namespace RegistrationSummary.Common.Services;

	public class GoogleSheetsService
	{
		public static SheetsService CreateSheetService(string credentialsFilePath)
		{
			if (string.IsNullOrEmpty(credentialsFilePath))
				throw new Exception("Cannot create new sheet service due to lack of the proper name in the appsettings.json file.");

			GoogleCredential credential;
			using (
				var stream = new FileStream(
					credentialsFilePath,
					FileMode.Open,
					FileAccess.Read
				)
			)
			{
				credential = GoogleCredential
					.FromStream(stream)
					.CreateScoped(SheetsService.Scope.Spreadsheets);
			}

			return new SheetsService(
				new BaseClientService.Initializer()
				{
					HttpClientInitializer = credential,
					ApplicationName = "Google Sheets API Example"
				}
			);
		}
	}
