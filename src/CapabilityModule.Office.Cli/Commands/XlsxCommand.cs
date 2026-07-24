using System.CommandLine;
using System.Text.Json;

namespace CapabilityModule.Office.Cli.Commands;

internal sealed class XlsxCommand
{
    public Command Command()
    {
        var cmd = new Command("xlsx", "Work with Excel (.xlsx) spreadsheets — read text, read rows, create, inspect metadata, and set cell values.");

        cmd.AddCommand(ReadSubCommand());
        cmd.AddCommand(CreateSubCommand());
        cmd.AddCommand(InfoSubCommand());
        cmd.AddCommand(SetCellSubCommand());

        return cmd;
    }

    private static Command ReadSubCommand()
    {
        var pathArg = new Argument<string>("path", "Path to the .xlsx file (relative to the restricted root).");
        var rowsOpt = new Option<bool>("--rows", "Return structured row/column data instead of flattened text.");
        var sheetOpt = new Option<string>("--sheet", () => string.Empty, "Sheet name to read (default: all sheets).");
        var rootOpt = SharedOptions.RootOption();

        var cmd = new Command("read", "Extract text content from an .xlsx spreadsheet.")
        {
            pathArg, rowsOpt, sheetOpt, rootOpt,
        };

        cmd.SetHandler((string path, bool rows, string sheet, string rootOverride) =>
        {
            try
            {
                var root = PathSecurity.EffectiveRoot(rootOverride);
                var fullPath = PathSecurity.ResolveWithinRoot(root, path);

                if (!File.Exists(fullPath))
                {
                    Console.Error.WriteLine($"error: file not found: {path}");
                    Environment.Exit(1);
                }

                if (rows)
                {
                    var sheetName = string.IsNullOrEmpty(sheet) ? null : sheet;
                    var data = XlsxEngine.ReadRows(fullPath, sheetName);
                    var result = new Dictionary<string, object?>
                    {
                        ["path"] = path,
                        ["resolved"] = fullPath,
                        ["sheets"] = data,
                    };
                    Console.WriteLine(JsonSerializer.Serialize(result));
                }
                else
                {
                    var text = XlsxEngine.ReadText(fullPath);
                    var result = new Dictionary<string, object?>
                    {
                        ["path"] = path,
                        ["resolved"] = fullPath,
                        ["content"] = text,
                    };
                    Console.WriteLine(JsonSerializer.Serialize(result));
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.Error.WriteLine($"error: {ex.Message}");
                Environment.Exit(2);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"error: {ex.Message}");
                Environment.Exit(3);
            }
        }, pathArg, rowsOpt, sheetOpt, rootOpt);

        return cmd;
    }

    private static Command CreateSubCommand()
    {
        var pathArg = new Argument<string>("path", "Path for the new .xlsx file (relative to the restricted root).");
        var sheetOpt = new Option<string>("--sheet", () => "Sheet1", "Sheet name for the new workbook.");
        var contentOpt = new Option<string>("--content", () => string.Empty,
            "JSON array of arrays of strings for cell data. If omitted, reads from stdin (content piping).");
        var rootOpt = SharedOptions.RootOption();

        var cmd = new Command("create", "Create a new .xlsx workbook from row data.")
        {
            pathArg, sheetOpt, contentOpt, rootOpt,
        };

        cmd.SetHandler((string path, string sheet, string content, string rootOverride) =>
        {
            try
            {
                var root = PathSecurity.EffectiveRoot(rootOverride);
                var fullPath = PathSecurity.ResolveWithinRoot(root, path);

                if (File.Exists(fullPath))
                {
                    Console.Error.WriteLine($"error: file already exists: {path} (delete it first or use a different path)");
                    Environment.Exit(5);
                }

                // If no content provided via --content, try stdin (content piping)
                if (string.IsNullOrEmpty(content) && Console.IsInputRedirected)
                {
                    content = Console.In.ReadToEnd();
                }

                if (string.IsNullOrEmpty(content))
                {
                    // Default to an empty workbook with one row
                    content = "[[]]";
                }

                var dir = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                XlsxEngine.Create(fullPath, sheet, content);

                var result = new Dictionary<string, object?>
                {
                    ["path"] = path,
                    ["resolved"] = fullPath,
                    ["sheet"] = sheet,
                };
                Console.WriteLine(JsonSerializer.Serialize(result));
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.Error.WriteLine($"error: {ex.Message}");
                Environment.Exit(2);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"error: {ex.Message}");
                Environment.Exit(3);
            }
        }, pathArg, sheetOpt, contentOpt, rootOpt);

        return cmd;
    }

    private static Command InfoSubCommand()
    {
        var pathArg = new Argument<string>("path", "Path to the .xlsx file (relative to the restricted root).");
        var rootOpt = SharedOptions.RootOption();

        var cmd = new Command("info", "Show metadata about an .xlsx workbook (sheet names, row/column counts, formula presence).")
        {
            pathArg, rootOpt,
        };

        cmd.SetHandler((string path, string rootOverride) =>
        {
            try
            {
                var root = PathSecurity.EffectiveRoot(rootOverride);
                var fullPath = PathSecurity.ResolveWithinRoot(root, path);

                if (!File.Exists(fullPath))
                {
                    Console.Error.WriteLine($"error: file not found: {path}");
                    Environment.Exit(1);
                }

                var info = XlsxEngine.GetInfo(fullPath);
                info["path"] = path;
                info["resolved"] = fullPath;
                Console.WriteLine(JsonSerializer.Serialize(info));
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.Error.WriteLine($"error: {ex.Message}");
                Environment.Exit(2);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"error: {ex.Message}");
                Environment.Exit(3);
            }
        }, pathArg, rootOpt);

        return cmd;
    }

    private static Command SetCellSubCommand()
    {
        var pathArg = new Argument<string>("path", "Path to the .xlsx file (relative to the restricted root).");
        var sheetOpt = new Option<string>("--sheet", () => "Sheet1", "Sheet name containing the cell.");
        var cellOpt = new Option<string>("--cell", "Cell reference, e.g. A1, B3.");
        var valueOpt = new Option<string>("--value", "New value for the cell.");
        var rootOpt = SharedOptions.RootOption();

        var cmd = new Command("set-cell", "Set the value of a single cell in an .xlsx workbook. The original file is versioned before overwriting.")
        {
            pathArg, sheetOpt, cellOpt, valueOpt, rootOpt,
        };

        cmd.SetHandler((string path, string sheet, string cell, string value, string rootOverride) =>
        {
            try
            {
                if (string.IsNullOrEmpty(cell))
                {
                    Console.Error.WriteLine("error: --cell must not be empty.");
                    Environment.Exit(1);
                }

                var root = PathSecurity.EffectiveRoot(rootOverride);
                var fullPath = PathSecurity.ResolveWithinRoot(root, path);

                if (!File.Exists(fullPath))
                {
                    Console.Error.WriteLine($"error: file not found: {path}");
                    Environment.Exit(1);
                }

                // Snapshot the pre-edit content to the version store
                var (version, versionPath) = VersionStore.Snapshot(fullPath, root, path);

                // Perform the cell edit
                XlsxEngine.SetCell(fullPath, sheet, cell, value);

                // Read the last-write timestamp after the overwrite
                var lastWriteUtc = File.GetLastWriteTimeUtc(fullPath);

                var result = new Dictionary<string, object?>
                {
                    ["path"] = path,
                    ["resolved"] = fullPath,
                    ["sheet"] = sheet,
                    ["cell"] = cell,
                    ["value"] = value,
                    ["version"] = version,
                    ["versionPath"] = versionPath,
                    ["lastModifiedUtc"] = lastWriteUtc.ToString("O"),
                };
                Console.WriteLine(JsonSerializer.Serialize(result));
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.Error.WriteLine($"error: {ex.Message}");
                Environment.Exit(2);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"error: {ex.Message}");
                Environment.Exit(3);
            }
        }, pathArg, sheetOpt, cellOpt, valueOpt, rootOpt);

        return cmd;
    }
}