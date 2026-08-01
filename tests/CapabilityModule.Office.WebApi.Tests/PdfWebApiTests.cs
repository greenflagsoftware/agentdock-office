using CapabilityModule.Office.WebApi.Cli;
using System.Text.Json;

namespace CapabilityModule.Office.WebApi.Tests;

public class PdfWebApiTests : IDisposable
{
    private readonly string _dir;

    public PdfWebApiTests()
    {
        _dir = Path.Combine(Path.GetTempPath(), "office-webapi-pdf-tests-" + Guid.NewGuid());
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

    // /view's isPdf branch resolves the path via the CLI's "read" command (to validate
    // it exists within the restricted root) and hands the frontend a downloadUrl for
    // pdf.js to fetch the raw bytes from — it does not attempt to parse or extract text
    // from the PDF itself, unlike the docx/xlsx branches.
    [Fact]
    public async Task PdfPath_ResolvesWithinRestrictedRoot()
    {
        var filePath = "sample.pdf";
        File.WriteAllBytes(Path.Combine(_dir, filePath), MinimalPdfBytes());

        var args = new List<string> { "read", filePath, "--root", _dir };
        var json = await CliRunner.RunAsync(args);

        using var doc = JsonDocument.Parse(json);
        var resolved = doc.RootElement.GetProperty("resolved").GetString();

        Assert.NotNull(resolved);
        Assert.EndsWith("sample.pdf", resolved);
    }

    [Fact]
    public async Task PdfPath_OutsideRestrictedRoot_Throws()
    {
        var args = new List<string> { "read", "../outside.pdf", "--root", _dir };

        await Assert.ThrowsAsync<CliToolException>(() => CliRunner.RunAsync(args));
    }

    private static byte[] MinimalPdfBytes() =>
        System.Text.Encoding.ASCII.GetBytes("%PDF-1.4\n%%EOF\n");
}
