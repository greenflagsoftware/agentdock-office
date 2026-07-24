namespace CapabilityModule.Office.Tests;

public class XlsxToolsValidationTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public async Task XlsxRead_NullOrEmptyPath_ThrowsArgumentException(string? invalidPath)
    {
        var ex = await Assert.ThrowsAsync<ArgumentException>(() =>
            Tools.XlsxTools.XlsxRead(invalidPath!));

        Assert.Contains("path", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public async Task XlsxCreate_NullOrEmptyPath_ThrowsArgumentException(string? invalidPath)
    {
        var ex = await Assert.ThrowsAsync<ArgumentException>(() =>
            Tools.XlsxTools.XlsxCreate(invalidPath!, "Sheet1", "[]"));

        Assert.Contains("path", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public async Task XlsxInfo_NullOrEmptyPath_ThrowsArgumentException(string? invalidPath)
    {
        var ex = await Assert.ThrowsAsync<ArgumentException>(() =>
            Tools.XlsxTools.XlsxInfo(invalidPath!));

        Assert.Contains("path", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public async Task XlsxSetCell_NullOrEmptyPath_ThrowsArgumentException(string? invalidPath)
    {
        var ex = await Assert.ThrowsAsync<ArgumentException>(() =>
            Tools.XlsxTools.XlsxSetCell(invalidPath!, "Sheet1", "A1", "value"));

        Assert.Contains("path", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public async Task XlsxSetCell_NullOrEmptyCellRef_ThrowsArgumentException(string? invalidCell)
    {
        var ex = await Assert.ThrowsAsync<ArgumentException>(() =>
            Tools.XlsxTools.XlsxSetCell("test.xlsx", "Sheet1", invalidCell!, "value"));

        Assert.Contains("cell", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task XlsxRead_ValidArgs_ReachesCliLayer_NotValidationError()
    {
        // Should not throw ArgumentException — validation passes.
        // The CLI call may succeed or fail with some other error (file system, etc.)
        var ex = await Record.ExceptionAsync(() =>
            Tools.XlsxTools.XlsxRead("test.xlsx"));

        // If it failed, it must not be an ArgumentException (validation error)
        if (ex != null)
        {
            Assert.IsNotType<ArgumentException>(ex);
        }
    }

    [Fact]
    public async Task XlsxCreate_ValidArgs_ReachesCliLayer_NotValidationError()
    {
        var ex = await Record.ExceptionAsync(() =>
            Tools.XlsxTools.XlsxCreate("test.xlsx", "Sheet1", "[[]]"));

        if (ex != null)
        {
            Assert.IsNotType<ArgumentException>(ex);
        }
    }

    [Fact]
    public async Task XlsxInfo_ValidArgs_ReachesCliLayer_NotValidationError()
    {
        var ex = await Record.ExceptionAsync(() =>
            Tools.XlsxTools.XlsxInfo("test.xlsx"));

        if (ex != null)
        {
            Assert.IsNotType<ArgumentException>(ex);
        }
    }

    [Fact]
    public async Task XlsxSetCell_ValidArgs_ReachesCliLayer_NotValidationError()
    {
        var ex = await Record.ExceptionAsync(() =>
            Tools.XlsxTools.XlsxSetCell("test.xlsx", "Sheet1", "A1", "value"));

        if (ex != null)
        {
            Assert.IsNotType<ArgumentException>(ex);
        }
    }
}