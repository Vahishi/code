public class PurchaseOrderViewModel
{
    public string PONumber { get; set; }
    public DateTime PODate { get; set; }
    public string VATNo { get; set; }
    public string Project { get; set; }
    public string Attention { get; set; }
    public string Reference { get; set; }
    public string Subject { get; set; }

    // Purchase Order Details
    public List<PurchaseOrderItem> Items { get; set; }
    public decimal TotalAmount { get; set; }
    public decimal Discount { get; set; }
    public decimal VATAmount { get; set; }
    public decimal NetAmount { get; set; }

    // Company details
    public string CompanyName { get; set; }
    public string Address { get; set; }
    public string ContactPerson { get; set; }

    public string AmountInWords { get; set; }
}

public class PurchaseOrderItem
{
    public string StockNumber { get; set; }
    public string Description { get; set; }
    public string Unit { get; set; }
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
    public decimal AmountBeforeTax { get; set; }
    public decimal VAT { get; set; }
    public decimal Discount { get; set; } // Add this line
    public decimal VATAmount { get; set; }
}
