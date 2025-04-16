namespace RegistrationSummary.Common.Models.Interfaces;

public interface IProduct
{
	string Type { get; }
	string Name { get; }

    public bool IsOddRow { get; set; }

    abstract IProduct Clone();
}
