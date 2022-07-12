using DotLiquid;

namespace Unidigital.Cobros.Interfaces;
class PosCharge : ILiquidizable
{
    public decimal ValueVES { get; set; }
    public decimal ExchangeRate { get; set; }
    public decimal ValueUSD { get; set; }
    public string Description { get; set; } = null!;
    public string Date { get; set; } = null!;

    public object ToLiquid()
    {
        return new
        {
            ValueVES = Util.formatNumber(ValueVES),
            ExchangeRate = Util.formatNumber(ExchangeRate),
            ValueUSD = Util.formatNumber(ValueUSD),
            Description,
            Date
        };
    }
}

class PosClient
{
    public string Id { get; set; } = null!;
    public string AfiliationCode { get; set; } = null!;
    public string Terminal { get; set; } = null!;
    public string Serial { get; set; } = null!;
    public string Model { get; set; } = null!;
    public string Name { get; set; } = null!;
    public string RIF { get; set; } = null!;
    public string State { get; set; } = null!;
    public string Email { get; set; } = null!;
    public string Bank { get; set; } = null!;
    public string Phone { get; set; } = null!;
    public IList<PosCharge> Charges { get; set; } = null!;
    public string ChargedVES { get; set; } = null!;
    public string ChargedUSD { get; set; } = null!;
    public string Debt { get; set; } = null!;
    public string Title { get; set; } = null!;
}




class CentralizedCharge
{
    public int Day { get; set; }
    public int Month { get; set; }
    public decimal ValueVES { get; set; }
    public decimal ExchangeRate { get; set; }
    public decimal ValueUSD { get; set; }
    public string Description { get; set; } = null!;
    public string Date { get; set; } = null!;
}

class CentralizedMonth
{
    public int Month { get; set; }
    public decimal MonthlyFee { get; set; }
    public decimal Charged { get; set; }
    public decimal Pending { get; set; }
}

class CentralizedUser
{
    public string Id { get; set; } = null!;
    public string Model { get; set; } = null!;
    public string Name { get; set; } = null!;
    public string RIF { get; set; } = null!;
    public string AffiliationCode { get; set; } = null!;
    public string Terminal { get; set; } = null!;
    public string State { get; set; } = null!;
    public string Bank { get; set; } = null!;
    public IList<CentralizedCharge> Charges { get; set; } = null!;
    public IList<CentralizedMonth> Months { get; set; } = null!;
    public string Fee { get; set; } = null!;
    public string ChargedTotal { get; set; } = null!;
    public string Debt { get; set; } = null!;
}

