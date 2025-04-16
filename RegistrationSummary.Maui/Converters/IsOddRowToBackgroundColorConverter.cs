using System.Globalization;

namespace RegistrationSummary.Maui.Converters;

public class IsOddRowToBackgroundColorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        // Check if the value is a boolean and true
        if (value is bool isOddRow)
        {
            // Return LightGray if true, otherwise Transparent
            return isOddRow ? Colors.LightGray : Colors.Transparent;
        }
        // In case the value is not a boolean, return a default color or throw an exception
        return Colors.Transparent;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
