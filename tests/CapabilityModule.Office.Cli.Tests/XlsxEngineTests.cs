namespace CapabilityModule.Office.Cli.Tests;

public class XlsxEngineTests : IDisposable
{
    private readonly string _dir;

    public XlsxEngineTests()
    {
        _dir = Path.Combine(Path.GetTempPath(), "office-cli-xlsx-tests-" + Guid.NewGuid());
        Directory.CreateDirectory(_dir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_dir))
        {
            try { Directory.Delete(_dir, recursive: true); }
            catch { /* best-effort cleanup */ }
        }
    }

    private string PathFor(string name) => Path.Combine(_dir, name);

    [Fact]
    public void Create_ThenReadText_RoundTripsRows()
    {
        var file = PathFor("roundtrip.xlsx");
        var rows = """[["Name","Role"],["Alice","Engineer"],["Bob","Designer"]]""";

        XlsxEngine.Create(file, "Sheet1", rows);
        var text = XlsxEngine.ReadText(file);

        Assert.Contains("Alice", text);
        Assert.Contains("Engineer", text);
        Assert.Contains("Bob", text);
        Assert.Contains("Designer", text);
        Assert.Contains("Name", text);
        Assert.Contains("Role", text);
    }

    [Fact]
    public void Create_ThenReadRows_ReturnsStructuredData()
    {
        var file = PathFor("structured.xlsx");
        var rows = """[["A","B"],["1","2"]]""";

        XlsxEngine.Create(file, "Data", rows);
        var data = XlsxEngine.ReadRows(file);

        Assert.Single(data);
        Assert.Equal("Data", data[0]["sheetName"]);

        var rowsList = (List<List<string?>>)data[0]["rows"]!;
        Assert.Equal(2, rowsList.Count);
        Assert.Equal("A", rowsList[0][0]);
        Assert.Equal("B", rowsList[0][1]);
        Assert.Equal("1", rowsList[1][0]);
        Assert.Equal("2", rowsList[1][1]);
    }

    [Fact]
    public void Create_ThenReadRows_WithSheetNameFilter()
    {
        var file = PathFor("filtered.xlsx");
        var rows = """[["Col1","Col2"]]""";

        XlsxEngine.Create(file, "Sheet1", rows);
        var data = XlsxEngine.ReadRows(file, "Sheet1");

        Assert.Single(data);
        Assert.Equal("Sheet1", data[0]["sheetName"]);

        // Non-matching sheet name returns empty
        var noMatch = XlsxEngine.ReadRows(file, "Nonexistent");
        Assert.Empty(noMatch);
    }

    [Fact]
    public void Create_EmptyContent_ProducesValidWorkbook()
    {
        var file = PathFor("empty.xlsx");
        XlsxEngine.Create(file, "Sheet1", "[[]]");

        var info = XlsxEngine.GetInfo(file);
        Assert.Equal(1, info["sheetCount"]);
    }

    [Fact]
    public void GetInfo_ReturnsCountsMatchingContent()
    {
        var file = PathFor("counts.xlsx");
        var rows = """[["H1","H2"],["A","B"],["C","D"]]""";

        XlsxEngine.Create(file, "Data", rows);
        var info = XlsxEngine.GetInfo(file);

        Assert.Equal(1, info["sheetCount"]);
        Assert.False(info.ContainsKey("error"));

        var sheets = (List<Dictionary<string, object?>>)info["sheets"]!;
        Assert.Single(sheets);
        Assert.Equal("Data", sheets[0]["name"]);
        Assert.Equal(3, sheets[0]["rowCount"]); // 3 rows used (header + 2 data)
        Assert.Equal(2, sheets[0]["columnCount"]);
    }

    [Fact]
    public void GetInfo_EmptyFile_Throws()
    {
        var file = PathFor("empty.xlsx");
        // Create an empty file (not a valid xlsx)
        File.WriteAllBytes(file, new byte[0]);

        Assert.ThrowsAny<Exception>(() => XlsxEngine.GetInfo(file));
    }

    [Fact]
    public void SetCell_UpdatesSingleCell()
    {
        var file = PathFor("setcell.xlsx");
        XlsxEngine.Create(file, "Sheet1", """[["Old"]]""");

        XlsxEngine.SetCell(file, "Sheet1", "A1", "NewValue");
        var text = XlsxEngine.ReadText(file);

        Assert.Contains("NewValue", text);
        Assert.DoesNotContain("Old", text);
    }

    [Fact]
    public void SetCell_MultipleSheets_UpdatesCorrectSheet()
    {
        var file = PathFor("multi-sheet-setcell.xlsx");
        using var wb = new ClosedXML.Excel.XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        ws1.Cell("A1").Value = "Original1";
        var ws2 = wb.Worksheets.Add("Sheet2");
        ws2.Cell("A1").Value = "Original2";
        wb.SaveAs(file);

        XlsxEngine.SetCell(file, "Sheet2", "A1", "Updated2");

        var text = XlsxEngine.ReadText(file);
        Assert.Contains("Original1", text);
        Assert.Contains("Updated2", text);
        Assert.DoesNotContain("Original2", text);
    }

    [Fact]
    public void GetInfo_ReportsFormulasCorrectly()
    {
        var file = PathFor("formulas.xlsx");
        using var wb = new ClosedXML.Excel.XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").Value = 10;
        ws.Cell("A2").Value = 20;
        ws.Cell("A3").FormulaA1 = "=SUM(A1:A2)";
        wb.SaveAs(file);

        var info = XlsxEngine.GetInfo(file);
        var sheets = (List<Dictionary<string, object?>>)info["sheets"]!;
        Assert.True((bool)info["hasFormulas"]!);
        Assert.True((bool)sheets[0]["hasFormulas"]!);
    }

    [Fact]
    public void ReadText_FormulaCell_DoesNotThrow()
    {
        var file = PathFor("formula-read.xlsx");
        using var wb = new ClosedXML.Excel.XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").Value = 10;
        ws.Cell("A2").Value = 20;
        ws.Cell("A3").FormulaA1 = "=SUM(A1:A2)";
        wb.SaveAs(file);

        // ReadText should not throw when encountering formula cells
        var text = XlsxEngine.ReadText(file);
        Assert.Contains("Sheet1", text);
    }

    [Fact]
    public void ReadRows_FormulaCell_DoesNotThrow()
    {
        var file = PathFor("formula-rows.xlsx");
        using var wb = new ClosedXML.Excel.XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("A1").Value = 5;
        ws.Cell("A2").FormulaA1 = "=A1*2";
        wb.SaveAs(file);

        var data = XlsxEngine.ReadRows(file);
        var rows = (List<List<string?>>)data[0]["rows"]!;
        Assert.Equal(2, rows.Count);
        Assert.True((bool)data[0]["hasFormulas"]!);
    }
}