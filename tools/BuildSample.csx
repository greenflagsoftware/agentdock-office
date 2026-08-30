#r "nuget: ClosedXML, 0.104.2"
using ClosedXML.Excel;

var filePath = "Team_Inventory.xlsx";
using var wb = new XLWorkbook();

var ws1 = wb.Worksheets.Add("Team Inventory");
ws1.Cell("A1").Value = "Item";
ws1.Cell("B1").Value = "Category";
ws1.Cell("C1").Value = "Assigned To";
ws1.Cell("D1").Value = "Status";
ws1.Cell("E1").Value = "Purchase Date";
ws1.Cell("F1").Value = "Value";
ws1.Range("A1:F1").Style.Font.Bold = true;

var items = new (string, string, string, string, string, string)[]
{
    ("MacBook Pro 16\"", "Laptop", "Alice Chen", "Active", "2026-01-15", "$3,499"),
    ("Dell UltraSharp 27\"", "Monitor", "Alice Chen", "Active", "2025-11-20", "$799"),
    ("Standing Desk", "Furniture", "Alice Chen", "Active", "2025-08-01", "$1,299"),
    ("ThinkPad X1 Carbon", "Laptop", "Bob Martinez", "Active", "2026-03-10", "$2,899"),
    ("Logitech MX Keys", "Keyboard", "Bob Martinez", "Active", "2026-03-10", "$199"),
    ("LG 34\" Ultrawide", "Monitor", "Bob Martinez", "Active", "2025-06-15", "$1,099"),
    ("Mac Mini M4", "Desktop", "Carol Nguyen", "Active", "2026-02-01", "$1,599"),
    ("AirPods Pro", "Accessories", "Carol Nguyen", "Active", "2026-02-01", "$249"),
    ("Magic Keyboard", "Keyboard", "Carol Nguyen", "Active", "2025-09-12", "$349"),
    ("Dell Latitude 5540", "Laptop", "David Park", "Active", "2025-07-22", "$1,899"),
    ("Jabra Evolve2 65", "Headset", "David Park", "Active", "2025-07-22", "$299"),
    ("ThinkPad Dock", "Dock", "Eve Johnson", "Active", "2026-04-05", "$329"),
    ("Surface Pro 9", "Tablet", "Eve Johnson", "Active", "2026-04-05", "$1,699"),
    ("Sony WH-1000XM5", "Headset", "Eve Johnson", "Active", "2025-10-01", "$329"),
};
for (int i = 0; i < items.Length; i++)
{
    var r = i + 2;
    ws1.Cell(r, 1).Value = items[i].Item1;
    ws1.Cell(r, 2).Value = items[i].Item2;
    ws1.Cell(r, 3).Value = items[i].Item3;
    ws1.Cell(r, 4).Value = items[i].Item4;
    ws1.Cell(r, 5).Value = items[i].Item5;
    ws1.Cell(r, 6).Value = items[i].Item6;
}
ws1.Columns().AdjustToContents();

var ws2 = wb.Worksheets.Add("Budget Overview");
ws2.Cell("A1").Value = "Department";
ws2.Cell("B1").Value = "Q1 Budget";
ws2.Cell("C1").Value = "Q2 Budget";
ws2.Cell("D1").Value = "Q3 Budget";
ws2.Cell("E1").Value = "Q4 Budget";
ws2.Range("A1:E1").Style.Font.Bold = true;
var budgets = new (string, string, string, string, string)[]
{
    ("Engineering", "$45,000", "$48,000", "$50,000", "$52,000"),
    ("Design", "$22,000", "$24,000", "$24,000", "$26,000"),
    ("Marketing", "$35,000", "$38,000", "$42,000", "$45,000"),
    ("Operations", "$18,000", "$18,000", "$20,000", "$20,000"),
};
for (int i = 0; i < budgets.Length; i++)
{
    var r = i + 2;
    ws2.Cell(r, 1).Value = budgets[i].Item1;
    ws2.Cell(r, 2).Value = budgets[i].Item2;
    ws2.Cell(r, 3).Value = budgets[i].Item3;
    ws2.Cell(r, 4).Value = budgets[i].Item4;
    ws2.Cell(r, 5).Value = budgets[i].Item5;
}
ws2.Columns().AdjustToContents();

wb.SaveAs(filePath);
Console.WriteLine($"Created {filePath}");