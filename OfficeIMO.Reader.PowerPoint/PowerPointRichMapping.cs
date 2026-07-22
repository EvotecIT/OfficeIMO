using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace OfficeIMO.Reader.PowerPoint;

internal static class PowerPointRichMapping {
    internal static OfficeDocumentReadResult Apply(
        PowerPointPresentation presentation,
        ReaderOptions readerOptions,
        ReaderPowerPointOptions options,
        OfficeDocumentReadResult result,
        CancellationToken cancellationToken) {
        PowerPointBuiltinDocumentProperties properties = presentation.BuiltinDocumentProperties;
        result.Source.Title = properties.Title;
        result.Source.Author = properties.Creator;
        result.Source.Subject = properties.Subject;
        result.Source.Keywords = properties.Keywords;

        var blocks = new List<OfficeDocumentBlock>();
        var tables = new List<ReaderTable>();
        var links = new List<OfficeDocumentLink>();
        var visuals = new List<ReaderVisual>();
        var pages = new List<OfficeDocumentPage>(presentation.Slides.Count);
        int tableIndex = 0;
        int linkIndex = 0;
        int shapeCount = 0;
        for (int slideIndex = 0; slideIndex < presentation.Slides.Count; slideIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PowerPointSlide slide = presentation.Slides[slideIndex];
            int slideNumber = slideIndex + 1;
            var slideBlocks = new List<OfficeDocumentBlock>();
            var slideTables = new List<ReaderTable>();
            var slideLinks = new List<OfficeDocumentLink>();
            PowerPointShape[] slideShapes = slide.EnumerateShapesDeep(
                slide.Shapes, includeHidden: options.IncludeHiddenShapes).ToArray();
            for (int shapeIndex = 0; shapeIndex < slideShapes.Length; shapeIndex++) {
                PowerPointShape shape = slideShapes[shapeIndex];
                shapeCount++;
                string shapeAnchor = "powerpoint-slide-" + slideNumber.ToString("D4", CultureInfo.InvariantCulture)
                    + "-shape-" + (shape.Id?.ToString(CultureInfo.InvariantCulture) ?? shapeIndex.ToString(CultureInfo.InvariantCulture));
                ReaderLocation location = BuildPowerPointLocation(result.Source.Path, slideNumber, shapeIndex, ResolvePowerPointShapeKind(shape), shapeAnchor);
                OfficeDocumentRegion region = BuildPowerPointRegion(shape);

                if (shape is PowerPointTextBox textBox) {
                    ProjectPowerPointTextBox(textBox, location, region, slideBlocks, slideLinks, ref linkIndex);
                } else if (shape is PowerPointTable table) {
                    ReaderTable mapped = MapPowerPointTable(table, location, tableIndex++, readerOptions.MaxTableRows);
                    tables.Add(mapped);
                    slideTables.Add(mapped);
                    slideBlocks.Add(new OfficeDocumentBlock {
                        Id = shapeAnchor,
                        Kind = "table",
                        Text = DocumentReaderEngine.BuildRichTableText(mapped),
                        Location = location,
                        Region = region
                    });
                    AddPowerPointTableLinks(table, location, slideLinks, ref linkIndex);
                } else if (shape is PowerPointChart chart) {
                    string chartText = shape.Name ?? "Chart";
                    if (chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot)) {
                        chartText = snapshot.Title ?? snapshot.Name;
                        string payload = BuildPowerPointChartPayload(snapshot);
                        visuals.Add(new ReaderVisual {
                            Kind = "chart",
                            Language = "officeimo-powerpoint-chart",
                            Content = payload,
                            PayloadHash = OfficeDocumentAssetHash.ComputeSha256Hex(Encoding.UTF8.GetBytes(payload)),
                            SourceName = snapshot.Name,
                            Width = snapshot.WidthPoints,
                            Height = snapshot.HeightPoints,
                            X = shape.LeftPoints,
                            Y = shape.TopPoints,
                            PlacedWidth = shape.WidthPoints,
                            PlacedHeight = shape.HeightPoints,
                            PlacementCount = 1,
                            HasGeometry = true,
                            IsAxisAligned = true,
                            Location = BuildPowerPointLocation(result.Source.Path, slideNumber, shapeIndex, "chart", shapeAnchor)
                        });
                    }
                    slideBlocks.Add(new OfficeDocumentBlock { Id = shapeAnchor, Kind = "chart", Text = chartText, Location = location, Region = region });
                } else {
                    slideBlocks.Add(new OfficeDocumentBlock {
                        Id = shapeAnchor,
                        Kind = ResolvePowerPointShapeKind(shape),
                        Text = shape.AltText ?? shape.Name ?? shape.ShapeContentType.ToString(),
                        Location = location,
                        Region = region
                    });
                }

                if (shape.Hyperlink != null) {
                    slideLinks.Add(BuildPowerPointLink(
                        "powerpoint-link-" + linkIndex.ToString("D4", CultureInfo.InvariantCulture),
                        shape.Hyperlink.ToString(),
                        shape.AltText ?? shape.Name,
                        location,
                        region));
                    linkIndex++;
                }
            }

            if (options.IncludeNotes && slide.Notes.TryGetExistingText(out string notesText)) {
                string notesAnchor = "powerpoint-slide-" + slideNumber.ToString("D4", CultureInfo.InvariantCulture) + "-notes";
                slideBlocks.Add(new OfficeDocumentBlock {
                    Id = notesAnchor,
                    Kind = "speaker-notes",
                    Text = notesText,
                    Location = BuildPowerPointLocation(result.Source.Path, slideNumber, null, "speaker-notes", notesAnchor)
                });
            }

            blocks.AddRange(slideBlocks);
            links.AddRange(slideLinks);
            pages.Add(new OfficeDocumentPage {
                Number = slideNumber,
                Name = ResolvePowerPointSlideName(slideShapes, slideNumber),
                Width = presentation.SlideSize.WidthPoints,
                Height = presentation.SlideSize.HeightPoints,
                Location = BuildPowerPointLocation(result.Source.Path, slideNumber, null, "slide", "powerpoint-slide-" + slideNumber.ToString("D4", CultureInfo.InvariantCulture)),
                Blocks = slideBlocks,
                Tables = slideTables,
                Assets = result.Assets.Where(asset => asset.Location.Slide == slideNumber).ToArray(),
                Links = slideLinks,
                Forms = Array.Empty<OfficeDocumentFormField>()
            });
        }

        var metadata = new[] {
            DocumentReaderEngine.BuildCountMetadataEntry("powerpoint-shape-count", "powerpoint.structure", "ShapeCount", shapeCount),
            DocumentReaderEngine.BuildCountMetadataEntry("powerpoint-chart-count", "powerpoint.structure", "ChartCount", visuals.Count)
        };
        return DocumentReaderEngine.EnrichDocumentResult(
            result,
            new[] { "officeimo.powerpoint.shape-model", "officeimo.powerpoint.chart-snapshot", "officeimo.reader.powerpoint.rich-v5" },
            blocks,
            tables,
            links,
            visuals,
            pages,
            metadata);
    }

    private static void ProjectPowerPointTextBox(
        PowerPointTextBox textBox,
        ReaderLocation shapeLocation,
        OfficeDocumentRegion region,
        List<OfficeDocumentBlock> blocks,
        List<OfficeDocumentLink> links,
        ref int linkIndex) {
        IReadOnlyList<PowerPointParagraph> paragraphs = textBox.Paragraphs;
        if (paragraphs.Count == 0) {
            blocks.Add(new OfficeDocumentBlock { Id = shapeLocation.BlockAnchor!, Kind = "text-box", Text = textBox.Text, Location = shapeLocation, Region = region });
            return;
        }

        bool isTitle = textBox.ShapePlaceholderType == PlaceholderValues.Title || textBox.ShapePlaceholderType == PlaceholderValues.CenteredTitle;
        var numberingState = new Dictionary<int, int>();
        for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++) {
            PowerPointParagraph paragraph = paragraphs[paragraphIndex];
            bool isList = paragraph.BulletCharacter != null || paragraph.IsNumbered;
            int level = paragraph.Level ?? 0;
            string? marker = paragraph.BulletCharacter;
            if (isTitle) {
                marker = null;
                numberingState.Clear();
            } else if (!isList) {
                numberingState.Clear();
            } else if (paragraph.IsNumbered) {
                foreach (int nestedLevel in numberingState.Keys
                             .Where(candidate => candidate > level)
                             .ToArray()) {
                    numberingState.Remove(nestedLevel);
                }
                int number = paragraph.NumberingStartAt
                    ?? (numberingState.TryGetValue(level, out int previous) ? previous + 1 : 1);
                numberingState[level] = number;
                marker = PowerPointNumberingFormatter.FormatMarker(number,
                    paragraph.NumberingScheme);
            } else {
                foreach (int resetLevel in numberingState.Keys
                             .Where(candidate => candidate >= level)
                             .ToArray()) {
                    numberingState.Remove(resetLevel);
                }
            }
            string kind = isTitle ? "heading" : isList ? "list-item" : "paragraph";
            ReaderLocation location = BuildPowerPointLocation(
                shapeLocation.Path,
                shapeLocation.Slide!.Value,
                shapeLocation.SourceBlockIndex,
                kind,
                shapeLocation.BlockAnchor + "-paragraph-" + paragraphIndex.ToString("D4", CultureInfo.InvariantCulture));
            blocks.Add(new OfficeDocumentBlock {
                Id = location.BlockAnchor!,
                Kind = kind,
                Text = paragraph.Text,
                Level = isTitle ? 1 : paragraph.Level,
                Marker = marker,
                Location = location,
                Region = region
            });
            AddPowerPointRunLinks(paragraph.Runs, location, region, links, ref linkIndex);
        }
    }

    private static void AddPowerPointTableLinks(
        PowerPointTable table,
        ReaderLocation location,
        List<OfficeDocumentLink> links,
        ref int linkIndex) {
        foreach (PowerPointTableRow row in table.RowItems) {
            foreach (PowerPointTableCell cell in row.Cells) {
                foreach (PowerPointParagraph paragraph in cell.Paragraphs) {
                    AddPowerPointRunLinks(paragraph.Runs, location, BuildPowerPointRegion(table), links, ref linkIndex);
                }
            }
        }
    }

    private static void AddPowerPointRunLinks(
        IReadOnlyList<PowerPointTextRun> runs,
        ReaderLocation location,
        OfficeDocumentRegion region,
        List<OfficeDocumentLink> links,
        ref int linkIndex) {
        for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
            PowerPointTextRun run = runs[runIndex];
            if (run.Hyperlink == null) continue;
            links.Add(BuildPowerPointLink(
                "powerpoint-link-" + linkIndex.ToString("D4", CultureInfo.InvariantCulture),
                run.Hyperlink.ToString(),
                run.Text,
                BuildPowerPointLocation(location.Path, location.Slide!.Value, location.SourceBlockIndex, "hyperlink", location.BlockAnchor + "-link-" + runIndex.ToString("D4", CultureInfo.InvariantCulture)),
                region));
            linkIndex++;
        }
    }

    private static OfficeDocumentLink BuildPowerPointLink(string id, string uri, string? text, ReaderLocation location, OfficeDocumentRegion region) {
        return new OfficeDocumentLink { Id = id, Kind = "uri", Uri = uri, Text = text, Location = location, Region = region };
    }

    private static ReaderTable MapPowerPointTable(PowerPointTable table, ReaderLocation location, int tableIndex, int maxRows) {
        IReadOnlyList<PowerPointTableRow> sourceRows = table.RowItems;
        int columnCount = Math.Max(table.Columns, sourceRows.Count == 0 ? 0 : sourceRows.Max(static row => row.Cells.Count));
        bool hasHeaderRow = table.HeaderRow && sourceRows.Count > 0;
        IReadOnlyList<string> columns = hasHeaderRow
            ? Enumerable.Range(0, columnCount).Select(index => GetPowerPointCellText(sourceRows[0], index, "Column " + (index + 1).ToString(CultureInfo.InvariantCulture))).ToArray()
            : DocumentReaderEngine.BuildFallbackColumns(columnCount);
        int dataStart = hasHeaderRow ? 1 : 0;
        int totalRows = Math.Max(0, sourceRows.Count - dataStart);
        int emittedRows = maxRows > 0 ? Math.Min(totalRows, maxRows) : totalRows;
        IReadOnlyList<IReadOnlyList<string>> rows = sourceRows.Skip(dataStart).Take(emittedRows)
            .Select(row => (IReadOnlyList<string>)Enumerable.Range(0, columnCount).Select(index => GetPowerPointCellText(row, index, string.Empty)).ToArray())
            .ToArray();
        ReaderLocation tableLocation = BuildPowerPointLocation(location.Path, location.Slide!.Value, location.SourceBlockIndex, "table", location.BlockAnchor ?? "powerpoint-table-" + tableIndex.ToString("D4", CultureInfo.InvariantCulture));
        tableLocation.TableIndex = tableIndex;
        return new ReaderTable {
            Title = table.Name ?? "PowerPoint table " + (tableIndex + 1).ToString(CultureInfo.InvariantCulture),
            Kind = "powerpoint-table",
            Location = tableLocation,
            Columns = columns,
            ColumnProfiles = ReaderTableProfiler.CreateProfiles(columns, rows),
            Rows = rows,
            TotalRowCount = totalRows,
            Truncated = emittedRows < totalRows
        };
    }

    private static string GetPowerPointCellText(PowerPointTableRow row, int index, string fallback) {
        if (index >= row.Cells.Count || string.IsNullOrWhiteSpace(row.Cells[index].Text)) return fallback;
        return row.Cells[index].Text;
    }

    private static string BuildPowerPointChartPayload(OfficeChartSnapshot snapshot) {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream)) {
            writer.WriteStartObject();
            writer.WriteString("name", snapshot.Name);
            writer.WriteString("title", snapshot.Title);
            writer.WriteString("kind", snapshot.ChartKind.ToString());
            writer.WritePropertyName("categories");
            writer.WriteStartArray();
            foreach (string category in snapshot.Data.Categories) writer.WriteStringValue(category);
            writer.WriteEndArray();
            writer.WritePropertyName("series");
            writer.WriteStartArray();
            foreach (OfficeChartSeries series in snapshot.Data.Series) {
                writer.WriteStartObject();
                writer.WriteString("name", series.Name);
                WritePowerPointChartValues(writer, "values", series.Values);
                if (series.XValues == null) {
                    writer.WriteNull("xValues");
                } else {
                    WritePowerPointChartValues(writer, "xValues", series.XValues);
                }
                writer.WriteString("kind", series.RenderKind?.ToString());
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
        return Encoding.UTF8.GetString(stream.ToArray());
    }

    private static void WritePowerPointChartValues(Utf8JsonWriter writer, string propertyName, IReadOnlyList<double> values) {
        writer.WritePropertyName(propertyName);
        writer.WriteStartArray();
        foreach (double value in values) writer.WriteNumberValue(value);
        writer.WriteEndArray();
    }

    private static string ResolvePowerPointSlideName(IEnumerable<PowerPointShape> slideShapes, int slideNumber) {
        PowerPointTextBox? title = slideShapes.OfType<PowerPointTextBox>().FirstOrDefault(textBox =>
            textBox.ShapePlaceholderType == PlaceholderValues.Title || textBox.ShapePlaceholderType == PlaceholderValues.CenteredTitle);
        string? text = title?.Text?.Trim();
        return string.IsNullOrWhiteSpace(text) ? "Slide " + slideNumber.ToString(CultureInfo.InvariantCulture) : text!;
    }

    private static string ResolvePowerPointShapeKind(PowerPointShape shape) {
        return shape.ShapeContentType.ToString().ToLowerInvariant();
    }

    private static OfficeDocumentRegion BuildPowerPointRegion(PowerPointShape shape) {
        return new OfficeDocumentRegion { X = shape.LeftPoints, Y = shape.TopPoints, Width = shape.WidthPoints, Height = shape.HeightPoints };
    }

    private static ReaderLocation BuildPowerPointLocation(string? path, int slide, int? sourceBlockIndex, string kind, string anchor) {
        return new ReaderLocation {
            Path = path,
            Slide = slide,
            SourceBlockIndex = sourceBlockIndex,
            SourceBlockKind = kind,
            BlockAnchor = anchor
        };
    }
}
