namespace RegistrationSummary.Blazor.ViewModels.Helpers;

public class ColumnBinding
{
    public string Label { get; }
    public Func<string?> Getter { get; }
    public Action<string?> Setter { get; }

    public ColumnBinding(string label, Func<string?> getter, Action<string?> setter)
    {
        Label = label;
        Getter = getter;
        Setter = setter;
    }

    public string? Value
    {
        get => Getter();
        set => Setter(value);
    }
}