using CapabilityModule.Office.WebApi.Cli;
using System.Text.Json;

namespace CapabilityModule.Office.WebApi.Tests;

public class XlsxWebApiTests : IDisposable
{
    private readonly string _dir;

    public XlsxWebApiTests()
    {
        _dir = Path.Combine(Path.GetTempPath(), "office-webapi-xlsx-tests-" + Guid.NewGuid());
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

    [Fact]
    public async Task XlsxRead_WithRowsFlag_ReturnsStructuredData()
    {
        // Create a test .xlsx file via the CLI
        var filePath = "test.xlsx";
        var rows = """[["Col1","Col2"],["Val1","Val2"]]""";
        await CliRunner.RunAsync(new List<string>
        {
            "xlsx", "create", filePath, "--sheet", "Sheet1", "--content", rows, "--root", _dir
        });

        var args = new List<string>
        {
            "xlsx", "read", filePath, "--rows", "--root", _dir
        };

        var json = await CliRunner.RunAsync(args);
        Assert.Contains("Sheet1", json);
        Assert.Contains("Col1", json);
        Assert.Contains("Val1", json);
    }

    [Fact]
    public async Task XlsxRead_PlainText_ReturnsFlattenedText()
    {
        var filePath = "test2.xlsx";
        var rows = """[["Hello","World"]]""";
        await CliRunner.RunAsync(new List<string>
        {
            "xlsx", "create", filePath, "--sheet", "Data", "--content", rows, "--root", _dir
        });

        var args = new List<string>
        {
            "xlsx", "read", filePath, "--root", _dir
        };

        var json = await CliRunner.RunAsync(args);
        Assert.Contains("Data", json);
        Assert.Contains("Hello", json);
        Assert.Contains("World", json);
    }

    [Fact]
    public async Task XlsxInfo_ReturnsMetadata()
    {
        var filePath = "info.xlsx";
        var rows = """[["Test"]]""";
        await CliRunner.RunAsync(new List<string>
        {
            "xlsx", "create", filePath, "--sheet", "Sheet1", "--content", rows, "--root", _dir
        });

        var args = new List<string>
        {
            "xlsx", "info", filePath, "--root", _dir
        };

        var json = await CliRunner.RunAsync(args);
        Assert.Contains("sheetCount", json);
        Assert.Contains("Sheet1", json);
    }

    [Fact]
    public async Task XlsxCreate_ThenRead_RoundTrips()
    {
        var filePath = "roundtrip.xlsx";
        var rows = """[["A","B"],["1","2"]]""";

        var createArgs = new List<string>
        {
            "xlsx", "create", filePath, "--sheet", "TestSheet", "--content", rows, "--root", _dir
        };

        var createJson = await CliRunner.RunAsync(createArgs);
        Assert.Contains("TestSheet", createJson);

        // Now read it back
        var readArgs = new List<string>
        {
            "xlsx", "read", filePath, "--root", _dir
        };

        var readJson = await CliRunner.RunAsync(readArgs);
        Assert.Contains("TestSheet", readJson);
        Assert.Contains("1", readJson);
        Assert.Contains("2", readJson);
    }

    [Fact]
    public async Task XlsxSetCell_UpdatesCell()
    {
        var filePath = "setcell.xlsx";
        var rows = """[["Original"]]""";
        await CliRunner.RunAsync(new List<string>
        {
            "xlsx", "create", filePath, "--sheet", "Sheet1", "--content", rows, "--root", _dir
        });

        var setCellArgs = new List<string>
        {
            "xlsx", "set-cell", filePath, "--sheet", "Sheet1", "--cell", "A1", "--value", "Updated", "--root", _dir
        };

        var setCellJson = await CliRunner.RunAsync(setCellArgs);
        Assert.Contains("version", setCellJson);
        Assert.Contains("versionPath", setCellJson);
        Assert.Contains("lastModifiedUtc", setCellJson);

        // Read back to verify
        var readArgs = new List<string>
        {
            "xlsx", "read", filePath, "--root", _dir
        };
        var readJson = await CliRunner.RunAsync(readArgs);
        Assert.Contains("Updated", readJson);
        Assert.DoesNotContain("Original", readJson);
    }
}