using Org.BouncyCastle.Utilities;
using RegistrationSummary.Common.Configurations;
using RegistrationSummary.Common.Enums;
using RegistrationSummary.Common.Models.Interfaces;

namespace RegistrationSummary.Common.Models;

public class Event
{
    public int Id { get; set; }
    public string Name { get; set; }
    public DateTime StartDate { get; set; }
    public EventType EventType { get; }
    public bool CoursesAreMerged { get; }
    public string SpreadSheetId { get; }
    public ColumnsConfiguration RawDataColumns { get; }
    public ColumnsConfiguration PreprocessedColumns { get; }
    public List<IProduct> Products { get; }

    public Event(int id, string name, DateTime? startDate, EventType eventType, bool coursesAreMerged, string spreadSheetId, ColumnsConfiguration rawDataColumns, ColumnsConfiguration preprocessedColumns, List<IProduct> products)
    {
        Id = id;
        Name = name;
        StartDate = startDate ?? DateTime.Now;
        EventType = eventType;
        CoursesAreMerged = coursesAreMerged;
        SpreadSheetId = spreadSheetId;
        RawDataColumns = rawDataColumns;
        PreprocessedColumns = preprocessedColumns;
        Products = products;

        for (int i = 0; i < Products.Count; i++)
        {
            Products[i].IsOddRow = i % 2 != 0;
        }
    }

    public Event Clone()
    {
        var clonedProducts = new List<IProduct>();
        foreach (var product in Products)
        {
            clonedProducts.Add(product.Clone());
        }

        var clonedRawDataColumns = RawDataColumns.Clone();
        var clonedPreprocessedColumns = PreprocessedColumns.Clone();

        return new Event(0, Name, StartDate, EventType, CoursesAreMerged, SpreadSheetId, clonedRawDataColumns, clonedPreprocessedColumns, clonedProducts);
    }
}
