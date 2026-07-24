using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol;
using ModelContextProtocol.Server;

namespace CapabilityModule.Office.Tools;

/// <summary>
/// MCP tools that shell out to the CapabilityModule.Office.CLI for xlsx operations.
/// Mirrors <see cref="DocxTools"/>'s pattern — each method maps to a CLI subcommand
/// invocation.
/// </summary>
[McpServerToolType]
public static class XlsxTools
{
    private static string ResolveRoot()
    {
        var env = Environment.GetEnvironmentVariable("OFFICE_CLI_ROOT");
        if (!string.IsNullOrWhiteSpace(env))
            return env;

        var dataDir = Path.Combine(AppContext.BaseDirectory, "data");
        Directory.CreateDirectory(dataDir);
        return dataDir;
    }

    private static List<string> BuildArgs(params string[] baseArgs)
    {
        var args = new List<string>(baseArgs);

        var root = ResolveRoot();
        var cwd = Directory.GetCurrentDirectory();
        if (!string.Equals(root, cwd, StringComparison.OrdinalIgnoreCase))
        {
            args.Add("--root");
            args.Add(root);
        }

        return args;
    }

    private static void ValidatePath(string path, string paramName)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path must not be null or empty.", paramName);
        }
    }

    private static async Task<JsonDocument> CallCliAsync(IReadOnlyList<string> arguments, TimeSpan? timeout = null)
    {
        string json;
        try
        {
            json = await CliRunner.RunAsync(arguments, timeout);
        }
        catch (CliToolException ex)
        {
            throw new McpException($"CLI tool call failed. {ex.Message}");
        }
        catch (CliTimeoutException ex)
        {
            throw new McpException(ex.Message);
        }
        catch (FileNotFoundException ex)
        {
            throw new McpException(
                $"CLI binary not found: {ex.FileName}. The module may not be deployed correctly.");
        }

        if (string.IsNullOrWhiteSpace(json))
        {
            throw new McpException(
                "CLI tool produced empty output. This may indicate an internal error.");
        }

        try
        {
            return JsonDocument.Parse(json);
        }
        catch (JsonException ex)
        {
            throw new McpException(
                $"CLI tool produced malformed JSON output: {ex.Message}. Raw output (first 200 chars): {json[..Math.Min(json.Length, 200)]}");
        }
    }

    [McpServerTool, Description("Read the plain text content of an .xlsx spreadsheet. Returns sheet names and cell values in a flattened text format.")]
    public static async Task<string> XlsxRead(
        [Description("Path to the .xlsx file, relative to the restricted root.")] string path)
    {
        ValidatePath(path, nameof(path));

        using var doc = await CallCliAsync(BuildArgs("xlsx", "read", path));
        var content = doc.RootElement.TryGetProperty("content", out var c) ? c.GetString() : "";
        return content ?? "";
    }

    [McpServerTool, Description("Create a new .xlsx workbook from row data. Rows are provided as a JSON array of arrays of strings.")]
    public static async Task<string> XlsxCreate(
        [Description("Path for the new .xlsx file, relative to the restricted root.")] string path,
        [Description("Sheet name (default 'Sheet1').")] string? sheet,
        [Description("JSON array of arrays of strings for cell data.")] string? content)
    {
        ValidatePath(path, nameof(path));

        var args = new List<string> { "xlsx", "create", path };
        args.Add("--sheet");
        args.Add(sheet ?? "Sheet1");
        args.Add("--content");
        args.Add(content ?? "[[]]");

        using var doc = await CallCliAsync(BuildArgs(args.ToArray()));

        var resolved = doc.RootElement.TryGetProperty("resolved", out var r) ? r.GetString() : path;
        return $"Created .xlsx at {resolved}";
    }

    [McpServerTool, Description("Get metadata about an .xlsx workbook: sheet count, sheet names, row/column counts per sheet, and whether formulas are present.")]
    public static async Task<string> XlsxInfo(
        [Description("Path to the .xlsx file, relative to the restricted root.")] string path)
    {
        ValidatePath(path, nameof(path));

        using var doc = await CallCliAsync(BuildArgs("xlsx", "info", path));

        var root = doc.RootElement;
        var sheetCount = root.TryGetProperty("sheetCount", out var sc) ? sc.GetInt32() : 0;
        var hasFormulas = root.TryGetProperty("hasFormulas", out var hf) && hf.GetBoolean();

        return $"Sheets: {sheetCount}\nContains formulas: {hasFormulas}";
    }

    [McpServerTool, Description("Set the value of a single cell in an .xlsx workbook. The original file is versioned before overwriting, so the previous content is recoverable.")]
    public static async Task<string> XlsxSetCell(
        [Description("Path to the .xlsx file, relative to the restricted root.")] string path,
        [Description("Sheet name containing the cell.")] string? sheet,
        [Description("Cell reference, e.g. 'A1', 'B3'.")] string cell,
        [Description("New value for the cell.")] string? value)
    {
        ValidatePath(path, nameof(path));

        if (string.IsNullOrWhiteSpace(cell))
        {
            throw new ArgumentException("Cell reference must not be null or empty.", nameof(cell));
        }

        using var doc = await CallCliAsync(BuildArgs(
            "xlsx", "set-cell", path,
            "--sheet", sheet ?? "Sheet1",
            "--cell", cell,
            "--value", value ?? ""));

        var root = doc.RootElement;
        var resolved = root.TryGetProperty("resolved", out var r) ? r.GetString() : path;
        var version = root.TryGetProperty("version", out var v) ? v.GetInt32() : 0;
        var versionPath = root.TryGetProperty("versionPath", out var vp) ? vp.GetString() : "";
        var lastModifiedUtc = root.TryGetProperty("lastModifiedUtc", out var lm) ? lm.GetString() : "";

        return $"Set cell {cell} in {resolved}. Version {version} saved to {versionPath}. Last modified: {lastModifiedUtc}";
    }
}