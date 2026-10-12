using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.IWork;
using OfficeIMO.Excel.IWork;
using OfficeIMO.Word.IWork;
using OfficeIMO.PowerPoint.IWork;
using A = DocumentFormat.OpenXml.Drawing;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

internal static class NativeNumericFormatWriterSmoke {
    internal static void Run() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "iwork-created-numeric-formats.json")));
        var root = manifest.RootElement;
        foreach (JsonElement family in root.GetProperty("families").EnumerateArray()) {
            string kind = family.GetProperty("kind").GetString()!;
            byte[] created = Create(kind, root);
            string expected = root.GetProperty("artifacts").GetProperty(family.GetProperty("created").GetString()!).GetProperty("sha256").GetString()!;
            Require(Convert.ToHexString(SHA256.HashData(created)).Equals(expected, StringComparison.OrdinalIgnoreCase)
                && created.SequenceEqual(Create(kind, root)), "Native numeric format creation differs from the Apple-accepted profile.");
            Validate(kind, created, root, false);
            string editedName = family.GetProperty("edited").GetString()!;
            byte[] edited = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, editedName));
            string editedHash = root.GetProperty("artifacts").GetProperty(editedName).GetProperty("sha256").GetString()!;
            Require(Convert.ToHexString(SHA256.HashData(edited)).Equals(editedHash, StringComparison.OrdinalIgnoreCase), "Native edited format fixture changed.");
            Validate(kind, edited, root, true);
        }
        Console.WriteLine("PASS | native explicit numeric format creation | three Apple-accepted deterministic graphs, 0-30 decimal precision, negative styles/grouping, native edits and saved DOCX/XLSX/PPTX");
    }

    private static byte[] Create(string kind, JsonElement root) {
        var data = IWorkTableData.Create(20, 2).SetText(1, 1, "Explicit numeric formats").SetText(1, 2, "Numeric values");
        foreach (JsonElement row in root.GetProperty("expectedRows").EnumerateArray()) {
            var format = row.GetProperty("format"); int decimals = format.GetProperty("decimalPlaces").GetInt32();
            bool grouping = format.GetProperty("thousandsSeparator").GetBoolean();
            var negative = (IWorkNegativeNumberStyle)format.GetProperty("negativeStyle").GetInt32();
            IWorkNumberFormat number = format.GetProperty("kind").GetString() switch {
                "Number" => IWorkNumberFormat.CreateNumber(decimals, grouping, negative),
                "Percentage" => IWorkNumberFormat.CreatePercentage(decimals, grouping, negative),
                "Scientific" => IWorkNumberFormat.CreateScientific(decimals), _ => throw new InvalidDataException()
            };
            int index = row.GetProperty("row").GetInt32();
            data.SetText(index, 1, row.GetProperty("label").GetString()!).SetNumber(index, 2, row.GetProperty("value").GetDouble(), number);
        }
        if (kind == "Numbers") {
            var document = IWorkNumbersDocument.Create(); document.AddSheet("Formats").AddTable(data, "Formats", 20, 20, 280, 24);
            return document.SaveBytes();
        }
        if (kind == "Pages") {
            var document = IWorkPagesDocument.Create(); document.AddTable(data, "Formats", 234, 24, fontSizePoints: 10);
            return document.SaveBytes();
        }
        var deck = IWorkKeynoteDocument.Create(); deck.AddSlide().AddTable(data, 120, 40, 720, 480, "Formats", fontSizePoints: 12);
        return deck.SaveBytes();
    }

    private static void Validate(string kind, byte[] bytes, JsonElement root, bool edited) {
        var source = IWorkSourceDocument.Open(bytes);
        IWorkTable table = kind switch {
            "Numbers" => source.ReadNumbers().Sheets[0].Tables.Single(),
            "Pages" => source.ReadPages().Tables.Single(), _ => source.ReadKeynote().Slides[0].Tables.Single()
        };
        foreach (JsonElement row in root.GetProperty("expectedRows").EnumerateArray()) {
            IWorkTableCell cell = table.GetCell(row.GetProperty("row").GetInt32(), 2)!;
            JsonElement expected = row.GetProperty("format"); IWorkNumberFormat? format = cell.NumberFormat;
            Require(cell.Kind == IWorkCellKind.Number && (double)cell.Value! == row.GetProperty(edited ? "editedValue" : "value").GetDouble()
                && format != null && format.Kind.ToString() == expected.GetProperty("kind").GetString()
                && format.DecimalPlaces == expected.GetProperty("decimalPlaces").GetInt32()
                && (int)format.NegativeStyle == expected.GetProperty("negativeStyle").GetInt32()
                && format.ThousandsSeparator == expected.GetProperty("thousandsSeparator").GetBoolean(), "Native numeric source value or format changed.");
        }
        using var output = new MemoryStream();
        if (kind == "Numbers") {
            using var result = source.ToExcelDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
            result.Value.Save(output); output.Position = 0;
            using var saved = SpreadsheetDocument.Open(output, false);
            Require(!new OpenXmlValidator().Validate(saved).Any(), "Saved numeric XLSX is invalid.");
            var cells = saved.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<S.Cell>().ToArray();
            var styles = saved.WorkbookPart.WorkbookStylesPart!.Stylesheet!;
            var codes = styles.NumberingFormats!.Elements<S.NumberingFormat>().ToDictionary(item => item.NumberFormatId!.Value,
                item => item.FormatCode!.Value);
            foreach (JsonElement row in root.GetProperty("expectedRows").EnumerateArray()) {
                var cell = cells.Single(item => item.CellReference!.Value == "B" + row.GetProperty("row").GetInt32());
                uint id = styles.CellFormats!.Elements<S.CellFormat>().ElementAt((int)cell.StyleIndex!.Value).NumberFormatId!.Value;
                Require((cell.DataType == null || cell.DataType.Value == S.CellValues.Number)
                    && double.Parse(cell.CellValue!.Text, System.Globalization.CultureInfo.InvariantCulture)
                        == row.GetProperty(edited ? "editedValue" : "value").GetDouble()
                    && codes[id] == row.GetProperty("excelFormatCode").GetString(), "Saved numeric XLSX value or format changed.");
            }
        } else if (kind == "Pages") {
            using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
            result.RequireCompleteEditableReconstruction(); result.Value.Save(output); output.Position = 0;
            using var saved = WordprocessingDocument.Open(output, false);
            Require(!new OpenXmlValidator().Validate(saved).Any(), "Saved numeric DOCX is invalid.");
            VerifyText(root, edited, saved.MainDocumentPart!.Document!.Body!.Elements<W.Table>().Single()
                .Elements<W.TableRow>().Select(row => row.Elements<W.TableCell>().Last().InnerText).ToArray());
        } else {
            using var result = source.ToPowerPointPresentationResult(new() { Mode = IWorkConversionMode.EditableOnly });
            result.RequireCompleteEditableReconstruction(); result.Value.Save(output); output.Position = 0;
            using var saved = PresentationDocument.Open(output, false);
            Require(!new OpenXmlValidator().Validate(saved).Any(), "Saved numeric PPTX is invalid.");
            VerifyText(root, edited, saved.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.Table>().Single()
                .Elements<A.TableRow>().Select(row => row.Elements<A.TableCell>().Last().InnerText).ToArray());
        }
    }

    private static void VerifyText(JsonElement root, bool edited, string[] values) {
        Require(values.Length == 20, "Saved numeric text table dimensions changed.");
        foreach (JsonElement row in root.GetProperty("expectedRows").EnumerateArray())
            Require(values[row.GetProperty("row").GetInt32() - 1]
                == row.GetProperty(edited ? "invariantEditedDisplay" : "invariantDisplay").GetString(), "Saved numeric display changed.");
    }

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
