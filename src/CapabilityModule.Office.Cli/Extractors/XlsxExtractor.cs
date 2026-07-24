using ClosedXML.Excel;

namespace CapabilityModule.Office.Cli.Extractors;

/// <summary>
/// .xlsx content extractor using ClosedXML. Each sheet becomes a chunk with
/// the sheet name as the heading path, preserving row/column structure.
/// </summary>
internal sealed class XlsxExtractor : IContentExtractor
{
    public NormalizedDocument Extract(string filePath)
    {
        using var workbook = new XLWorkbook(filePath);

        if (!workbook.Worksheets.Any())
            return new NormalizedDocument();

        var paragraphTexts = new List<string>();
        var chapters = new List<ContentChunk>();
        var sheetIndex = 0;

        foreach (var ws in workbook.Worksheets)
        {
            var sheetRows = new List<string>();
            foreach (var row in ws.RowsUsed())
            {
                var cells = row.Cells().Select(c => c.GetString()).ToList();
                sheetRows.Add(string.Join("\t", cells));
            }

            var sheetText = string.Join(Environment.NewLine, sheetRows);
            if (string.IsNullOrWhiteSpace(sheetText))
                continue;

            paragraphTexts.Add($"=== {ws.Name} ===");
            paragraphTexts.AddRange(sheetRows);

            chapters.Add(new ContentChunk
            {
                Text = sheetText,
                HeadingPath = new[] { ws.Name },
                PageNumber = sheetIndex,
            });

            sheetIndex++;
        }

        var allText = string.Join(Environment.NewLine, paragraphTexts);

        return new NormalizedDocument
        {
            Text = allText,
            Paragraphs = paragraphTexts,
            Chapters = chapters,
        };
    }
}