using Google.Apis.Sheets.v4.Data;

namespace RegistrationSummary.Common.Styles;

public static class SharedFormats
{
    // === Cell Formats ===

    public static readonly (CellFormat Format, string Fields) BoldCenterWrapHeaderFormat = (
        new CellFormat
        {
            TextFormat = new TextFormat { Bold = true },
            WrapStrategy = "WRAP",
            HorizontalAlignment = "CENTER"
        },
        "userEnteredFormat.textFormat.bold,userEnteredFormat.wrapStrategy,userEnteredFormat.horizontalAlignment"
    );

    public static readonly (CellFormat Format, string Fields) CurrencyPlnCellFormat = (
        new CellFormat
        {
            NumberFormat = new NumberFormat { Type = "CURRENCY", Pattern = "#,##0 \"zł\"" }
        },
        "userEnteredFormat.numberFormat"
    );

    // === Borders ===
    public static readonly Border SolidBlackBorder = new()
    {
        Style = "SOLID",
        Width = 1,
        Color = new Color { Red = 0, Green = 0, Blue = 0 }
    };

    // === Dimension Properties ===
    public static readonly (DimensionProperties Props, string Fields) NarrowColumnWidth = (
        new DimensionProperties
        {
            PixelSize = 55
        },
        "pixelSize"
    );
}