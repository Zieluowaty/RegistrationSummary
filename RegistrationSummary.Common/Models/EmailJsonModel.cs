namespace RegistrationSummary.Common.Models;

public class EmailJsonModel
{
    public string Name { get; set; }
    public string Title { get; set; }

    public List<string> Header { get; set; } = new();
    public List<string> PaymentInfo { get; set; } = new();
    public List<string> CourseInfo { get; set; } = new();
    public List<string> Footer { get; set; } = new();

    public string HeaderMerged => string.Join("", Header);
    public string PaymentInfoMerged => string.Join("", PaymentInfo);
    public string CourseInfoMerged => string.Join("", CourseInfo);
    public string FooterMerged => string.Join("", Footer);
}
