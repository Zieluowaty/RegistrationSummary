using System.Globalization;

namespace RegistrationSummary.Maui.Converters;

public class TimeSpanToTimeStringConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is TimeSpan timeSpan)
        {
            return timeSpan.ToString(@"hh\:mm");
        }

        return string.Empty;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        return new NotImplementedException();
    }
}
