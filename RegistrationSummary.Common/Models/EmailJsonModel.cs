namespace RegistrationSummary.Common.Models;

public class EmailJsonModel
{
    public string Name { get; set; }
    public string Title { get; set; }

    private List<string> _headers = new List<string>();
    public List<string> Header { set => _headers = value; }
    public string HeaderMerged { get => _headers.Aggregate("", (current, header) => $"{current}{header}"); }

    private List<string> _paymentInfos = new List<string>();
    public List<string> PaymentInfo { set => _paymentInfos = value; }
    public string PaymentInfoMerged { get => _paymentInfos.Aggregate("", (current, paymentInfo) => $"{current}{paymentInfo}"); }

    private List<string> _courseInfos = new List<string>();
    public List<string> CourseInfo { set => _courseInfos = value; }
    public string CourseInfoMerged { get => _courseInfos.Aggregate("", (current, courseInfo) => $"{current}{courseInfo}"); }

    private List<string> _footers = new List<string>();
    public List<string> Footer { set => _footers = value; }
    public string FooterMerged { get => _footers.Aggregate("", (current, footer) => $"{current}{footer}"); }
}