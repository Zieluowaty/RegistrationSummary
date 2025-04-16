namespace RegistrationSummary.Common.Models;

public class Settings
{
	public string? ConfigFilesRoot { get; set; }
	public string? SpreadsheetId { get; set; }
	public string Mail { get; set; }
    public string Password { get; set; }
    public string ServerName { get; set; }
    public int ServerPort { get; set; }
    public string EventsDataFilePath { get; set; }
    public string EmailsTemplatesFilePath { get; set; }
    public string CredentialsFilePath { get; set; }
    public string RawDataTabName { get; set; }
    public string PreprocessedDataTabName { get; set; }
    public string SummaryTabName { get; set; }
    public string GroupBalanceTabName { get; set; }
    public string LeaderText { get; set; }
    public string FollowerText { get; set; }
    public string SoloText { get; set; }
    public string TestMailRecepientEmailAddress { get; set; }
    public int[] Prices { get; set; }
    public int LogRetentionDays { get; set; } = 30;
}