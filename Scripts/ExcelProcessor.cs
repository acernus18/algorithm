using System.Text.Json;

namespace Algorithm.Scripts;

internal class PurchaseRecord
{
    public int ID { get; set; }
    public string CreatedAt { get; set; }
    public string UpdatedAt { get; set; }
    public string? DeletedAt { get; set; }
    public string EntryID { get; set; }
    public string VoucherID { get; set; }
    public string Date { get; set; }
    public string Type { get; set; }
    public string Code { get; set; }
    public string State { get; set; }
    public int Quantity { get; set; }
    public double Weight { get; set; }
    public string Dimensions { get; set; }
    public string Shape { get; set; }
    public string Currency { get; set; }
    public double UnitPrice { get; set; }
    public double TotalPrice { get; set; }
    public int Discount { get; set; }
    public string Cert { get; set; }
    public string CertID { get; set; }
    public int Expenses { get; set; }
    public string ParcelID { get; set; }
    public string InventoryID { get; set; }
    public object Comments { get; set; }
    public string Remarks { get; set; }
}

public static class ExcelProcessor
{
    public static void Process()
    {
        using var workbook = new ClosedXML.Excel.XLWorkbook("/Users/maples/Downloads/Book1.xlsx");
        var sheet = workbook.Worksheet(1);
        var records = new List<PurchaseRecord>();
        for (var i = 0; i < 195; i++)
        {
            var temp = new PurchaseRecord
            {
                ID = 0,
                CreatedAt = "0001-01-01T00:00:00Z",
                UpdatedAt = "0001-01-01T00:00:00Z",
                DeletedAt = null,
                EntryID = "",
                VoucherID = "",
                Date = "0001-01-01T00:00:00Z",
                Type = "",
                Code = "DCK",
                State = "",
                Quantity = 1,
                Weight = 0.0,
                Dimensions = "",
                Shape = "",
                Currency = "",
                UnitPrice = 0,
                TotalPrice = 0,
                Discount = 0,
                Cert = "",
                CertID = "",
                Expenses = 115,
                ParcelID = "",
                InventoryID = "",
                Remarks = "",
            };
            const int remark = 2;
            const int type = 3;
            const int weight = 4;
            const int dimensions = 13;
            const int shape = 12;
            const int totalPrice = 14;
            const int cert = 9;
            const int certId = 10;
            var row = sheet.Row(i + 1);
            temp.Remarks = row.Cell(remark).GetString();
            temp.Type = row.Cell(type).GetString();
            temp.Weight = row.Cell(weight).GetDouble();
            temp.Dimensions = row.Cell(dimensions).GetString();
            temp.Shape = row.Cell(shape).GetString();
            temp.TotalPrice = row.Cell(totalPrice).GetDouble();
            temp.Cert = row.Cell(cert).GetString();
            temp.CertID = row.Cell(certId).GetString();
            temp.UnitPrice = Math.Round(temp.TotalPrice / temp.Weight, 3);
            records.Add(temp);
        }

        Console.WriteLine(JsonSerializer.Serialize(records));
    }
}