namespace RegistrationSummary.Common.Configurations;

public class SettingConfiguration
{	
    public MailerConfiguration MailerConfiguration { get; set; } = new();

    // Tabs' titles.
    public string RawDataTabName { get; set; }
    public string PreprocessedDataTabName { get; set; }
    public string SummaryTabName { get; set; }
    public string GroupBalanceTabName { get; set; }

    // Role headers' names.
    public string LeaderText { get; set; }
    public string FollowerText { get; set; }
    public string SoloText { get; set; }

    public int[] Prices { get; set; }
    
    public int LogRetentionDays { get; set; } = 30;
}