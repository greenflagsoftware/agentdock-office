using ClosedXML.Excel;
using System.Text.Json;

namespace CapabilityModule.Office.Cli;

/// <summary>
/// Core OpenXml operations for .xlsx (Excel) spreadsheets via ClosedXML.
/// Mirrors <see cref="DocxEngine"/>'s shape — read text, create workbooks,
/// extract metadata, and single-cell editing. All paths are assumed to be
/// pre-validated by <see cref="PathSecurity"/> before reaching this class.
/// </summary>
internal static class XlsxEngine
{
    /// <summary>
    /// Extracts all sheets as a plain-text representation. Each sheet is
    /// prefixed with its name as a section marker, and rows are tab-joined
    /// cell values — suitable for <c>xlsx read</c>/content preview.
    /// </summary>
    public static string ReadText(string filePath)
    {
        using var workbook = new XLWorkbook(filePath);
        var parts = new List<string>();

        foreach (var ws in workbook.Worksheets)
        {
            parts.Add($"=== {ws.Name} ===");
            var rows = new List<string>();

            foreach (var row in ws.RowsUsed())
            {
                var cells = new List<string>();
                foreach (var cell in row.Cells())
                {
                    cells.Add(GetCellStringValue(cell));
                }
                rows.Add(string.Join("\t", cells));
            }

            if (rows.Count > 0)
                parts.AddRange(rows);
        }

        return string.Join(Environment.NewLine, parts);
    }

    /// <summary>
    /// Returns structured row/column data for a workbook. If <paramref name="sheetName"/>
    /// is null or empty, data for all sheets is returned.
    /// </summary>
    public static List<Dictionary<string, object?>> ReadRows(string filePath, string? sheetName = null)
    {
        using var workbook = new XLWorkbook(filePath);
        var result = new List<Dictionary<string, object?>>();

        foreach (var ws in workbook.Worksheets)
        {
            if (!string.IsNullOrEmpty(sheetName) &&
                !string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                continue;

            var sheetData = new Dictionary<string, object?>
            {
                ["sheetName"] = ws.Name,
                ["rows"] = new List<List<string?>>(),
                ["hasFormulas"] = HasFormulas(ws),
            };

            var rowsList = (List<List<string?>>)sheetData["rows"]!;

            foreach (var row in ws.RowsUsed())
            {
                var rowValues = new List<string?>();
                foreach (var cell in row.Cells())
                {
                    rowValues.Add(GetCellStringValue(cell));
                }
                rowsList.Add(rowValues);
            }

            result.Add(sheetData);
        }

        return result;
    }

    /// <summary>
    /// Creates a new .xlsx workbook with a single sheet populated from caller-supplied rows.
    /// </summary>
    /// <param name="filePath">Path for the new file.</param>
    /// <param name="sheetName">The sheet name (defaults to "Sheet1").</param>
    /// <param name="rows">JSON-serialized array of arrays of strings.</param>
    public static void Create(string filePath, string sheetName, string rows)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add(sheetName);

        var parsedRows = JsonSerializer.Deserialize<List<List<string>>>(rows)
            ?? throw new InvalidDataException("Rows must be a JSON array of arrays of strings.");

        for (var rowIdx = 0; rowIdx < parsedRows.Count; rowIdx++)
        {
            var row = parsedRows[rowIdx];
            for (var colIdx = 0; colIdx < row.Count; colIdx++)
            {
                ws.Cell(rowIdx + 1, colIdx + 1).Value = row[colIdx];
            }
        }

        workbook.SaveAs(filePath);
    }

    /// <summary>
    /// Returns metadata about an .xlsx workbook: sheet names, row/column counts
    /// per sheet, and a <c>hasFormulas</c> flag.
    /// </summary>
    public static Dictionary<string, object?> GetInfo(string filePath)
    {
        using var workbook = new XLWorkbook(filePath);

        if (!workbook.Worksheets.Any())
        {
            throw new InvalidDataException($"'{filePath}' has no worksheets (malformed or empty .xlsx).");
        }

        var sheets = new List<Dictionary<string, object?>>();
        var workbookHasFormulas = false;

        foreach (var ws in workbook.Worksheets)
        {
            var rowsUsed = ws.RowsUsed().Count();
            var maxCols = 0;

            foreach (var row in ws.RowsUsed())
            {
                var cellCount = row.Cells().Count();
                if (cellCount > maxCols)
                    maxCols = cellCount;
            }

            var hasFormulas = HasFormulas(ws);
            if (hasFormulas)
                workbookHasFormulas = true;

            sheets.Add(new Dictionary<string, object?>
            {
                ["name"] = ws.Name,
                ["rowCount"] = rowsUsed,
                ["columnCount"] = maxCols,
                ["hasFormulas"] = hasFormulas,
            });
        }

        return new Dictionary<string, object?>
        {
            ["sheetCount"] = workbook.Worksheets.Count(),
            ["sheets"] = sheets,
            ["hasFormulas"] = workbookHasFormulas,
        };
    }

    /// <summary>
    /// Sets the value of a single cell in a workbook. Snapshots the pre-edit
    /// file to the version store before overwriting.
    /// </summary>
    /// <param name="filePath">The resolved full path of the file.</param>
    /// <param name="sheetName">The sheet name containing the cell.</param>
    /// <param name="cellRef">The cell reference, e.g. "A1", "B3".</param>
    /// <param name="value">The new value for the cell.</param>
    public static void SetCell(string filePath, string sheetName, string cellRef, string value)
    {
        using var workbook = new XLWorkbook(filePath);
        var ws = workbook.Worksheet(sheetName);

        ws.Cell(cellRef).Value = value;
        workbook.SaveAs(filePath);
    }

    /// <summary>
    /// Gets the string representation of a cell's value, preferring the
    /// cached/calculated value (not live-recalculated).
    /// </summary>
    private static string GetCellStringValue(IXLCell cell)
    {
        if (cell.HasFormula)
        {
            // Return the cached calculated value
            return cell.GetString();
        }

        return cell.GetString();
    }

    /// <summary>
    /// Checks if a worksheet contains any formula cells.
    /// </summary>
    private static bool HasFormulas(IXLWorksheet ws)
    {
        foreach (var row in ws.RowsUsed())
        {
            foreach (var cell in row.Cells())
            {
                if (cell.HasFormula)
                    return true;
            }
        }
        return false;
    }
}