using RegistrationSummary.Common.Models.Interfaces;

namespace RegistrationSummary.Common.Models;

public class Tshirt : IProduct
{
	public string Type { get; }
	public string Name { get; }
	public string Size { get; set; }
	public int PricePln { get; set; }
	public int PriceEuro { get; set; }
    public bool IsOddRow { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

    public Tshirt(string name, int pricePln, int priceEuro, string type = "", string size = "")
	{
		Name = name;

		PricePln = pricePln;
		PriceEuro = priceEuro;

		Type = type;
		Size = size;
	}

    public IProduct Clone()
    {
        throw new NotImplementedException();
    }
}
