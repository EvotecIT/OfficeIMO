#if NET8_0_OR_GREATER
using System;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.Json;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Excel;

namespace OfficeIMO.TestAssets;

/// <summary>Shared independent validator for Node and browser-produced TypeScript corpus files.</summary>
internal static class JavaScriptWorkbookContract {
    internal static void Verify(string path, JsonElement? spec = null, long? maxCharactersInPart = null) {
        using SpreadsheetDocument sdk = SpreadsheetDocument.Open(path, false);
        string[] errors = new OpenXmlValidator().Validate(sdk).Select(e => e.Description).ToArray();
        Require(errors.Length == 0, Path.GetFileName(path) + ": " + string.Join("; ", errors));
        using var model = ExcelDocument.Load(path, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        var readerOptions = new ReaderOptions();
        if (maxCharactersInPart.HasValue) readerOptions.OpenXmlMaxCharactersInPart = maxCharactersInPart.Value;
        var generic = new OfficeDocumentReaderBuilder().AddExcelHandler().Build().ReadDocument(path, readerOptions);
        Require(generic.Kind == ReaderInputKind.Excel && generic.CapabilitiesUsed.Contains("officeimo.reader.excel.rich-v5"), "The generic Excel adapter did not run.");
        if (spec is not JsonElement fixture) return;
        Sheet[] wireSheets = sdk.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToArray();
        string[] decodedNames = wireSheets.Select(s => XmlConvert.DecodeName(s.Name!.Value!)).ToArray();
        Require(decodedNames.Distinct(StringComparer.OrdinalIgnoreCase).Count() == decodedNames.Length, "Decoded sheet names are not unique.");
        Require(decodedNames.All(n => n.Length is > 0 and <= 31 && n.IndexOfAny(new[] { '[', ']', ':', '*', '?', '/', '\\' }) < 0), "Decoded sheet name contains forbidden characters.");
        int sheetIndex = 0;
        foreach (JsonElement sheet in fixture.GetProperty("sheets").EnumerateArray()) {
            string expectedName = sheet.TryGetProperty("expectedName", out JsonElement renamed) ? renamed.GetString()! : sheet.GetProperty("name").GetString()!;
            Require(decodedNames[sheetIndex] == expectedName, "Decoded sheet name differs: " + decodedNames[sheetIndex]);
            string wireName = wireSheets[sheetIndex++].Name!.Value!;
            foreach (JsonProperty expected in sheet.GetProperty("expected").EnumerateObject()) {
                using var data = ExcelDocument.OpenDataReader(path, new ExcelReadOptions {
                    SheetName = wireName, A1Range = expected.Name + ":" + expected.Name, HasHeaderRow = false
                });
                Require(data.Read(), "Expected cell is absent: " + expected.Name);
                object value = data.GetValue(0);
                bool equal = expected.Value.ValueKind switch {
                    JsonValueKind.String => Equals(expected.Value.GetString(), value),
                    JsonValueKind.Number => Math.Abs(expected.Value.GetDouble() - Convert.ToDouble(value, CultureInfo.InvariantCulture)) < 1e-9,
                    JsonValueKind.True => Equals(true, value),
                    JsonValueKind.False => Equals(false, value),
                    JsonValueKind.Object => Equals(DateTime.SpecifyKind(DateTime.Parse(expected.Value.GetProperty("value").GetString()!, CultureInfo.InvariantCulture,
                        DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal), DateTimeKind.Unspecified), value),
                    _ => throw new InvalidDataException("Unsupported expected cell value.")
                };
                Require(equal, Path.GetFileName(path) + ": cell " + expected.Name + " differs: " + value);
            }
        }
        if (fixture.GetProperty("name").GetString() == "ooxml-attributes") {
            Stylesheet styles = sdk.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            Require(styles.Fonts!.Elements<Font>().Any(f => XmlConvert.DecodeName(f.FontName?.Val?.Value ?? "") == "Font_x0041_"), "Literal font name differs.");
            Require(styles.NumberingFormats!.Elements<NumberingFormat>().Any(f => XmlConvert.DecodeName(f.FormatCode?.Value ?? "") == "\"_x003A_\"0"), "Literal number-format code differs.");
        }
        if (fixture.GetProperty("name").GetString() == "report-table") VerifyReport(sdk.WorkbookPart!);
        if (fixture.TryGetProperty("parts", out JsonElement parts)) {
            using Package package = Package.Open(path, FileMode.Open, FileAccess.Read);
            foreach (JsonElement part in parts.EnumerateArray()) {
                JsonElement relationship = part.GetProperty("relationship");
                var source = new Uri(relationship.TryGetProperty("source", out JsonElement suppliedSource) ? suppliedSource.GetString()! : "/xl/workbook.xml", UriKind.Relative);
                string id = relationship.GetProperty("id").GetString()!;
                PackageRelationship link = source.OriginalString == "/" ? package.GetRelationship(id) : package.GetPart(source).GetRelationship(id);
                Require(link.TargetMode == TargetMode.Internal && !link.TargetUri.IsAbsoluteUri, "Relationship became an external URI.");
                Uri resolved = PackUriHelper.ResolvePartUri(source, link.TargetUri);
                Require(PackUriHelper.ComparePartUri(resolved, new Uri(part.GetProperty("uri").GetString()!, UriKind.Relative)) == 0 && package.PartExists(resolved),
                    "Relationship did not resolve to its internal part.");
                using var text = new StreamReader(package.GetPart(resolved).GetStream(FileMode.Open, FileAccess.Read));
                Require(text.ReadToEnd() == part.GetProperty("data").GetString(), "Relationship target contents differ.");
            }
        }
        if (fixture.GetProperty("name").GetString() == "styles-custom") {
            WorkbookPart workbook = sdk.WorkbookPart!;
            Stylesheet styles = workbook.WorkbookStylesPart!.Stylesheet!;
            Require(styles.Fonts!.Elements<Font>().Any(f => f.Bold is not null && f.Italic is not null && f.Color?.Rgb?.Value == "FF123456"), "Registered font differs.");
            Require(styles.Borders!.Elements<Border>().Any(b => b.BottomBorder?.Style?.Value == BorderStyleValues.Thin), "Registered border differs.");
            Require(styles.Fills!.Descendants<ForegroundColor>().Any(f => f.Rgb?.Value == "FFABCDEF"), "Registered fill differs.");
            Cell cell = workbook.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Single(c => c.CellReference?.Value == "A2");
            CellFormat applied = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
            Font font = styles.Fonts.Elements<Font>().ElementAt((int)applied.FontId!.Value);
            Fill fill = styles.Fills.Elements<Fill>().ElementAt((int)applied.FillId!.Value);
            Border border = styles.Borders.Elements<Border>().ElementAt((int)applied.BorderId!.Value);
            Require(font.Bold is not null && font.Italic is not null && font.Color?.Rgb?.Value == "FF123456" &&
                fill.Descendants<ForegroundColor>().Any(f => f.Rgb?.Value == "FFABCDEF") && border.BottomBorder?.Style?.Value == BorderStyleValues.Thin,
                "Column style composition lost its font, fill or border.");
            Require(styles.NumberingFormats!.Elements<NumberingFormat>().Single(f => f.NumberFormatId?.Value == applied.NumberFormatId!.Value).FormatCode?.Value == "0.000",
                "Column number-format override was not applied.");
            Require(workbook.CustomXmlParts.Count() == 1, "Custom XML part relationship differs.");
        }
    }
    private static void VerifyReport(WorkbookPart workbook) {
        WorksheetPart sheet = workbook.WorksheetParts.Single();
        Worksheet worksheet = sheet.Worksheet ?? throw new InvalidDataException("Report worksheet is absent.");
        Table table = sheet.TableDefinitionParts.Single().Table!;
        Require(table.Name?.Value == "ControllerReport" && table.Reference?.Value == "A1:D3", "Report table name or range differs.");
        Require(table.TableStyleInfo?.Name?.Value == "TableStyleMedium9" && table.TableStyleInfo?.ShowRowStripes?.Value == true, "Report table banding differs.");
        Pane pane = worksheet.Descendants<Pane>().Single();
        Require(pane.HorizontalSplit?.Value == 1 && pane.VerticalSplit?.Value == 1 && pane.TopLeftCell?.Value == "B2", "Report frozen panes differ.");
        Stylesheet styles = workbook.WorkbookStylesPart!.Stylesheet!;
        CellFormat Applied(string address) {
            Cell cell = worksheet.Descendants<Cell>().Single(c => c.CellReference?.Value == address);
            return styles.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
        }
        CellFormat number = Applied("B3"), date = Applied("C3");
        Require(number.NumberFormatId?.Value == Applied("B2").NumberFormatId?.Value && date.NumberFormatId?.Value == Applied("C2").NumberFormatId?.Value && date.NumberFormatId?.Value > 0,
            "Highlighting lost number or date formatting.");
        Font font = styles.Fonts!.Elements<Font>().ElementAt((int)number.FontId!.Value);
        Fill fill = styles.Fills!.Elements<Fill>().ElementAt((int)number.FillId!.Value);
        Require(font.Bold is not null && font.Color?.Rgb?.Value == "FFC00000" && fill.Descendants<ForegroundColor>().Any(f => f.Rgb?.Value == "FFFCE4D6"), "Report highlighting differs.");
        var link = sheet.HyperlinkRelationships.Single();
        Require(link.IsExternal && link.Uri.ToString() == "https://example.com/report?site=lodz&view=health", "Report hyperlink differs.");
        Require(XmlConvert.DecodeName(worksheet.Descendants<Hyperlink>().Single().Tooltip?.Value ?? "") == "Open Łódź _x0041_ report", "Literal hyperlink tooltip differs.");
        var anchor = sheet.DrawingsPart!.WorksheetDrawing!.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor>().Single();
        Require(anchor.FromMarker?.RowId?.Text == "5" && anchor.FromMarker.ColumnId?.Text == "0" && anchor.Extent?.Cx?.Value == 1714500 && anchor.Extent.Cy?.Value == 762000,
            "Report image anchor or size differs.");
        using var bytes = new MemoryStream();
        using (Stream image = sheet.DrawingsPart.ImageParts.Single().GetStream(FileMode.Open, FileAccess.Read)) image.CopyTo(bytes);
        Require(bytes.ToArray().SequenceEqual(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jfKsAAAAASUVORK5CYII=")), "Report image bytes differ.");
    }
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidDataException(message); }
}
#endif
