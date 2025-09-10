using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        [Trait("Category","ExcelLinks")]
        public void Excel_LinkByHeader_InternalLinks_Styled() {
            string filePath = Path.Combine(_directoryWithFiles, "LinksByHeader_Internal.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath))
            {
                var s = doc.AddWorkSheet("Summary");
                // Header
                s.Cell(1, 1, "Domain");
                // Two rows
                s.Cell(2, 1, "domain-001.example");
                s.Cell(3, 1, "domain-002.example");
                // Create destination sheets
                doc.AddWorkSheet("domain-001.example");
                doc.AddWorkSheet("domain-002.example");

                // Link by header (styled)
                s.LinkByHeaderToInternalSheets("Domain", rowFrom: 2, rowTo: 3, targetA1: "A1", styled: true);
                doc.Save(false);
            }

            using (var pkg = SpreadsheetDocument.Open(filePath, false))
            {
                var wb = pkg.WorkbookPart!;
                var sheets = wb.Workbook.Sheets!.Elements<Sheet>();
                var summary = sheets.First(x => x.Name!.Value == "Summary");
                var wsPart = (WorksheetPart)wb.GetPartById(summary.Id!);
                var ws = wsPart.Worksheet;
                var hyperlinks = ws.Elements<Hyperlinks>().FirstOrDefault();
                Assert.NotNull(hyperlinks);
                // Expect two hyperlinks for A2 and A3
                var items = hyperlinks!.Elements<Hyperlink>().ToList();
                Assert.Equal(2, items.Count);
                Assert.Contains(items, h => h.Reference!.Value == "A2" && h.Location!.Value == "'domain-001.example'!A1");
                Assert.Contains(items, h => h.Reference!.Value == "A3" && h.Location!.Value == "'domain-002.example'!A1");

                // Verify styled link (blue + underline) for A2
                var sd = ws.GetFirstChild<SheetData>();
                Assert.NotNull(sd);
                var row2 = sd!.Elements<Row>().First(r => r.RowIndex!.Value == 2U);
                var cellA2 = row2.Elements<Cell>().First(c => c.CellReference!.Value == "A2");
                Assert.NotNull(cellA2.StyleIndex);
                uint styleIdx = cellA2.StyleIndex!.Value;
                var ss = wb.WorkbookStylesPart!.Stylesheet!;
                var cellFormats = ss.CellFormats!;
                var cellFormat = cellFormats.Elements<CellFormat>().ElementAt((int)styleIdx);
                Assert.NotNull(cellFormat.FontId);
                uint fontId = cellFormat.FontId!.Value;
                var font = ss.Fonts!.Elements<Font>().ElementAt((int)fontId);
                // Underline element present and hyperlink blue (FF0563C1)
                Assert.NotNull(font.Underline);
                string rgb = font.Color?.Rgb?.Value ?? string.Empty;
                Assert.Equal("FF0563C1", rgb);
            }
        }

        [Fact]
        [Trait("Category","ExcelLinks")]
        public void Excel_LinkByHeader_ExternalLinks_SmartDisplay() {
            string filePath = Path.Combine(_directoryWithFiles, "LinksByHeader_External.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath))
            {
                var s = doc.AddWorkSheet("Summary");
                s.Cell(1, 1, "RFC");
                s.Cell(2, 1, "rfc7208");

                s.LinkByHeaderToUrls(
                    header: "RFC",
                    rowFrom: 2,
                    rowTo: 2,
                    urlForCellText: rfc => $"https://datatracker.ietf.org/doc/html/{rfc}",
                    titleForCellText: null,
                    styled: true);

                doc.Save(false);
            }

            using (var pkg = SpreadsheetDocument.Open(filePath, false))
            {
                var wb = pkg.WorkbookPart!;
                var sheets = wb.Workbook.Sheets!.Elements<Sheet>();
                var summary = sheets.First(x => x.Name!.Value == "Summary");
                var wsPart = (WorksheetPart)wb.GetPartById(summary.Id!);
                var ws = wsPart.Worksheet;
                var hyperlinks = ws.Elements<Hyperlinks>().FirstOrDefault();
                Assert.NotNull(hyperlinks);
                var items = hyperlinks!.Elements<Hyperlink>().ToList();
                Assert.Single(items);
                var h1 = items[0];
                Assert.Equal("A2", h1.Reference!.Value);
                // External link uses Id (relationship)
                Assert.False(string.IsNullOrEmpty(h1.Id!.Value));
                Assert.NotEmpty(wsPart.HyperlinkRelationships);

                // Display text should be smart "RFC 7208"
                // Cell is shared string; read SharedString item
                var sd = ws.GetFirstChild<SheetData>();
                var row2 = sd!.Elements<Row>().First(r => r.RowIndex!.Value == 2U);
                var cell = row2.Elements<Cell>().First(c => c.CellReference!.Value == "A2");
                Assert.Equal(CellValues.SharedString, cell.DataType!.Value);
                int ssid = int.Parse(cell.CellValue!.InnerText);
                var sst = wb.SharedStringTablePart!.SharedStringTable!;
                string text = sst.Elements<SharedStringItem>().ElementAt(ssid).InnerText;
                Assert.Equal("RFC 7208", text);
            }
        }

        [Fact]
        [Trait("Category","ExcelLinks")]
        public void Excel_LinkByHeader_In_Table_Internal_And_External() {
            string filePath = Path.Combine(_directoryWithFiles, "LinksByHeader_InTable.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath))
            {
                var s = doc.AddWorkSheet("Summary");
                // Build a 2-column table with headers Domain, RFC and two data rows
                s.Cell(1, 1, "Domain"); s.Cell(1, 2, "RFC");
                s.Cell(2, 1, "domain-001.example"); s.Cell(2, 2, "rfc7208");
                s.Cell(3, 1, "domain-002.example"); s.Cell(3, 2, "rfc6376");
                s.AddTable("A1:B3", hasHeader: true, name: "TblSummary", style: OfficeIMO.Excel.TableStyle.TableStyleMedium9, includeAutoFilter: true);
                // Destination sheets
                doc.AddWorkSheet("domain-001.example");
                doc.AddWorkSheet("domain-002.example");

                // Link inside table by header
                s.LinkByHeaderToInternalSheetsInTable("TblSummary", "Domain", targetA1: "A1", styled: true);
                s.LinkByHeaderToUrlsInTable("TblSummary", "RFC", urlForCellText: rfc => $"https://datatracker.ietf.org/doc/html/{rfc}", styled: true);
                doc.Save(false);
            }

            using (var pkg = SpreadsheetDocument.Open(filePath, false))
            {
                var wb = pkg.WorkbookPart!;
                var sheets = wb.Workbook.Sheets!.Elements<Sheet>();
                var summary = sheets.First(x => x.Name!.Value == "Summary");
                var wsPart = (WorksheetPart)wb.GetPartById(summary.Id!);
                var ws = wsPart.Worksheet;
                var hyperlinks = ws.Elements<Hyperlinks>().FirstOrDefault();
                Assert.NotNull(hyperlinks);
                var items = hyperlinks!.Elements<Hyperlink>().ToList();
                // Expect 4 links total: A2,A3 internal; B2,B3 external
                Assert.Equal(4, items.Count);
                Assert.Contains(items, h => h.Reference!.Value == "A2" && h.Location!.Value == "'domain-001.example'!A1");
                Assert.Contains(items, h => h.Reference!.Value == "A3" && h.Location!.Value == "'domain-002.example'!A1");
                Assert.Contains(items, h => h.Reference!.Value == "B2" && !string.IsNullOrEmpty(h.Id?.Value));
                Assert.Contains(items, h => h.Reference!.Value == "B3" && !string.IsNullOrEmpty(h.Id?.Value));
            }
        }

        [Fact]
        [Trait("Category","ExcelLinks")]
        public void Excel_LinkByHeader_In_Range_Internal_And_External() {
            string filePath = Path.Combine(_directoryWithFiles, "LinksByHeader_InRange.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath))
            {
                var s = doc.AddWorkSheet("Summary");
                // Header + rows (no table)
                s.Cell(1, 1, "Domain"); s.Cell(1, 2, "RFC");
                s.Cell(2, 1, "domain-001.example"); s.Cell(2, 2, "rfc7208");
                s.Cell(3, 1, "domain-002.example"); s.Cell(3, 2, "rfc6376");
                // Destination sheets
                doc.AddWorkSheet("domain-001.example");
                doc.AddWorkSheet("domain-002.example");

                // Link within the rectangular range
                s.LinkByHeaderToInternalSheetsInRange("A1:B3", "Domain", targetA1: "A1", styled: true);
                s.LinkByHeaderToUrlsInRange("A1:B3", "RFC", urlForCellText: rfc => $"https://datatracker.ietf.org/doc/html/{rfc}", styled: true);
                doc.Save(false);
            }

            using (var pkg = SpreadsheetDocument.Open(filePath, false))
            {
                var wb = pkg.WorkbookPart!;
                var sheets = wb.Workbook.Sheets!.Elements<Sheet>();
                var summary = sheets.First(x => x.Name!.Value == "Summary");
                var wsPart = (WorksheetPart)wb.GetPartById(summary.Id!);
                var ws = wsPart.Worksheet;
                var hyperlinks = ws.Elements<Hyperlinks>().FirstOrDefault();
                Assert.NotNull(hyperlinks);
                var items = hyperlinks!.Elements<Hyperlink>().ToList();
                // Expect 4 links total: A2,A3 internal; B2,B3 external
                Assert.Equal(4, items.Count);
                Assert.Contains(items, h => h.Reference!.Value == "A2" && h.Location!.Value == "'domain-001.example'!A1");
                Assert.Contains(items, h => h.Reference!.Value == "A3" && h.Location!.Value == "'domain-002.example'!A1");
                Assert.Contains(items, h => h.Reference!.Value == "B2" && !string.IsNullOrEmpty(h.Id?.Value));
                Assert.Contains(items, h => h.Reference!.Value == "B3" && !string.IsNullOrEmpty(h.Id?.Value));
            }
        }
    }
}
