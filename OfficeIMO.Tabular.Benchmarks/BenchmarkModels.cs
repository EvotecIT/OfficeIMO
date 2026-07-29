using System.Runtime.Serialization;

namespace OfficeIMO.Tabular.Benchmarks;

public readonly record struct Observation(int Rows, int Cells, long Checksum);

[DataContract]
public sealed class SalesRecord {
    [DataMember(Name = "Region")]
    public string Region { get; set; } = string.Empty;

    [DataMember(Name = "Country")]
    public string Country { get; set; } = string.Empty;

    [DataMember(Name = "Item Type")]
    public string ItemType { get; set; } = string.Empty;

    [DataMember(Name = "Sales Channel")]
    public string SalesChannel { get; set; } = string.Empty;

    [DataMember(Name = "Order Priority")]
    public string OrderPriority { get; set; } = string.Empty;

    [DataMember(Name = "Order Date")]
    public DateTime OrderDate { get; set; }

    [DataMember(Name = "Order ID")]
    public int OrderId { get; set; }

    [DataMember(Name = "Ship Date")]
    public DateTime ShipDate { get; set; }

    [DataMember(Name = "Units Sold")]
    public int UnitsSold { get; set; }

    [DataMember(Name = "Unit Price")]
    public decimal UnitPrice { get; set; }

    [DataMember(Name = "Unit Cost")]
    public decimal UnitCost { get; set; }

    [DataMember(Name = "Total Revenue")]
    public decimal TotalRevenue { get; set; }

    [DataMember(Name = "Total Cost")]
    public decimal TotalCost { get; set; }

    [DataMember(Name = "Total Profit")]
    public decimal TotalProfit { get; set; }
}
