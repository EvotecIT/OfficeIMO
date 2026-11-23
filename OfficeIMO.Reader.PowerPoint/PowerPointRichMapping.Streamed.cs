using OfficeIMO.PowerPoint;
using System.Globalization;

namespace OfficeIMO.Reader.PowerPoint;

internal static partial class PowerPointRichMapping {
    /// <summary>
    /// Adds the blocks, tables and page of one streamed slide: the items <see cref="Apply"/> maps from the loaded
    /// model, with the same identifiers, kinds, text and locations. Geometry, links and chart snapshots are left out
    /// because the text projection, the only caller, drops them. Returns false for content only the model maps.
    /// </summary>
    internal static bool AddStreamedSlide(PowerPointSlideContent slide, string? path, ReaderOptions readerOptions,
        ReaderPowerPointOptions options, StreamedSlideProjection projection) {
        int slideNumber = slide.Number;
        var slideBlocks = new List<OfficeDocumentBlock>();
        var slideTables = new List<ReaderTable>();
        int linkIndex = 0;
        int shapeIndex = 0;
        string? title = null;
        foreach (PowerPointShapeContent shape in slide.Shapes) {
            if (!options.IncludeHiddenShapes && !shape.IsVisible) continue;
            projection.ShapeCount++;
            string kind = shape.Kind.ToString().ToLowerInvariant();
            string shapeAnchor = "powerpoint-slide-" + slideNumber.ToString("D4", CultureInfo.InvariantCulture)
                + "-shape-" + (shape.Id?.ToString(CultureInfo.InvariantCulture) ?? shapeIndex.ToString(CultureInfo.InvariantCulture));
            ReaderLocation location = BuildPowerPointLocation(path, slideNumber, shapeIndex, kind, shapeAnchor);

            if (shape.Kind == PowerPointShapeContentType.TextBox) {
                if (title == null && shape.IsTitle) title = shape.Text.Trim();
                ProjectPowerPointParagraphs(shape.Paragraphs, shape.IsTitle, location, null, slideBlocks, null, ref linkIndex);
            } else if (shape.Table is PowerPointTableContent table) {
                // The model cannot size a table without a column grid; leave that table to it.
                if (!table.HasGrid) return false;
                IReadOnlyList<IReadOnlyList<string>> cellTexts = table.Rows
                    .Select(row => (IReadOnlyList<string>)row.Cells.Select(static cell => cell.Text).ToArray())
                    .ToArray();
                ReaderTable mapped = MapPowerPointTable(table.Name, table.ColumnCount, table.HeaderRow, cellTexts,
                    location, projection.TableIndex++, readerOptions.MaxTableRows);
                projection.Tables.Add(mapped);
                slideTables.Add(mapped);
                slideBlocks.Add(new OfficeDocumentBlock {
                    Id = shapeAnchor,
                    Kind = "table",
                    Text = DocumentReaderEngine.BuildRichTableText(mapped),
                    Location = location
                });
            } else if (shape.Kind == PowerPointShapeContentType.Chart) {
                projection.ChartCount++;
                slideBlocks.Add(new OfficeDocumentBlock { Id = shapeAnchor, Kind = "chart", Text = shape.Name ?? "Chart", Location = location });
            } else {
                slideBlocks.Add(new OfficeDocumentBlock {
                    Id = shapeAnchor,
                    Kind = kind,
                    Text = shape.AltText ?? shape.Name ?? shape.Kind.ToString(),
                    Location = location
                });
            }
            shapeIndex++;
        }

        if (options.IncludeNotes && slide.Notes != null) {
            string notesAnchor = "powerpoint-slide-" + slideNumber.ToString("D4", CultureInfo.InvariantCulture) + "-notes";
            slideBlocks.Add(new OfficeDocumentBlock {
                Id = notesAnchor,
                Kind = "speaker-notes",
                Text = slide.Notes,
                Location = BuildPowerPointLocation(path, slideNumber, null, "speaker-notes", notesAnchor)
            });
        }

        // Slide pages would hold these same blocks and tables, whose locations already name the slide, and the text
        // projection drops pages when it reduces the result, so the streamed read does not build them.
        projection.Blocks.AddRange(slideBlocks);
        return true;
    }

    /// <summary>Completes a streamed result the way <see cref="Apply"/> completes a model result.</summary>
    internal static OfficeDocumentReadResult CompleteStreamed(OfficeDocumentReadResult result,
        (string? Title, string? Creator, string? Subject, string? Keywords) properties, StreamedSlideProjection projection, IReadOnlyList<string> capabilities) {
        result.Source.Title = properties.Title;
        result.Source.Author = properties.Creator;
        result.Source.Subject = properties.Subject;
        result.Source.Keywords = properties.Keywords;
        // The model maps locations under the result's source path, which an empty deck leaves unset.
        string? path = result.Source.Path;
        foreach (OfficeDocumentBlock block in projection.Blocks) block.Location!.Path = path;
        foreach (ReaderTable table in projection.Tables) table.Location!.Path = path;
        var metadata = new[] {
            DocumentReaderEngine.BuildCountMetadataEntry("powerpoint-shape-count", "powerpoint.structure", "ShapeCount", projection.ShapeCount),
            DocumentReaderEngine.BuildCountMetadataEntry("powerpoint-chart-count", "powerpoint.structure", "ChartCount", projection.ChartCount)
        };
        return DocumentReaderEngine.EnrichDocumentResult(result, capabilities, projection.Blocks, projection.Tables,
            Array.Empty<OfficeDocumentLink>(), Array.Empty<ReaderVisual>(), Array.Empty<OfficeDocumentPage>(), metadata);
    }
}

/// <summary>Blocks, tables and pages accumulated while slides are streamed.</summary>
internal sealed class StreamedSlideProjection {
    internal List<OfficeDocumentBlock> Blocks { get; } = new();
    internal List<ReaderTable> Tables { get; } = new();
    internal int TableIndex { get; set; }
    internal int ShapeCount { get; set; }
    internal int ChartCount { get; set; }
}
