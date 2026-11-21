using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

/// <summary>
/// The Open XML SDK takes part only where its own API does the work: streaming a large sheet with
/// <see cref="OpenXmlWriter"/> and editing one cell of a loaded package. It also writes the recalculation fixture,
/// a neutral producer every library then opens.
/// </summary>
internal static class SdkLanes {
    internal static byte[] Write1M() {
        var output = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook)) {
            WorkbookPart workbookPart = document.AddWorkbookPart();
            WorksheetPart sheetPart = workbookPart.AddNewPart<WorksheetPart>();
            using (OpenXmlWriter writer = OpenXmlWriter.Create(sheetPart)) {
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());
                WriteRow(writer, 1, Workloads.LargeHeader.Select(InlineText).ToArray());
                for (int i = 1; i <= Workloads.LargeRows; i++) {
                    WriteRow(writer, i + 1, new[] {
                        Number(i), InlineText(Workloads.LargeName(i)), Number(Workloads.LargeAmount(i)),
                        new Cell { DataType = CellValues.Boolean, CellValue = new CellValue(Workloads.LargeFlag(i) ? "1" : "0") }
                    });
                }
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            AddWorkbook(workbookPart, ("Data", sheetPart));
        }
        return output.ToArray();
    }

    internal static byte[] Preserve(byte[] input) {
        var output = new MemoryStream();
        output.Write(input, 0, input.Length);
        using (var document = SpreadsheetDocument.Open(output, isEditable: true)) {
            WorkbookPart workbookPart = document.WorkbookPart!;
            Sheet first = workbookPart.Workbook.Sheets!.Elements<Sheet>().First();
            var sheetPart = (WorksheetPart)workbookPart.GetPartById(first.Id!);
            SheetData data = sheetPart.Worksheet.GetFirstChild<SheetData>()!;
            Row? row = data.Elements<Row>().FirstOrDefault(candidate => candidate.RowIndex!.Value >= Workloads.PreserveRow);
            if (row == null || row.RowIndex!.Value != Workloads.PreserveRow) {
                var created = new Row { RowIndex = (uint)Workloads.PreserveRow };
                if (row == null) data.Append(created); else data.InsertBefore(created, row);
                row = created;
            }
            string reference = "Z" + Workloads.PreserveRow.ToString(CultureInfo.InvariantCulture);
            Cell? after = row.Elements<Cell>().FirstOrDefault(cell => ColumnOf(cell.CellReference!) >= Workloads.PreserveColumn);
            Cell edited = InlineText(Workloads.PreserveText);
            edited.CellReference = reference;
            if (after == null) row.Append(edited);
            else if (ColumnOf(after.CellReference!) == Workloads.PreserveColumn) {
                edited.StyleIndex = after.StyleIndex;
                row.ReplaceChild(edited, after);
            } else row.InsertBefore(edited, after);
        }
        return output.ToArray();
    }

    /// <summary>
    /// Data: an id, a lookup key, VLOOKUP, IF/ROUND and INDEX/MATCH formulas and a category; Lookup: 1,000 keys and
    /// values; Summary: a SUMIFS per category. No cached values, so every library has to calculate.
    /// </summary>
    internal static byte[] RecalcFixture() {
        int last = Workloads.RecalcRows + 1;
        int lookupLast = Workloads.LookupRows + 1;
        var output = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook)) {
            WorkbookPart workbookPart = document.AddWorkbookPart();
            WorksheetPart dataPart = workbookPart.AddNewPart<WorksheetPart>();
            using (OpenXmlWriter writer = OpenXmlWriter.Create(dataPart)) {
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());
                WriteRow(writer, 1, new[] { "Id", "Key", "Lookup", "Adjusted", "Matched", "Category" }.Select(InlineText).ToArray());
                for (int i = 1; i <= Workloads.RecalcRows; i++) {
                    int r = i + 1;
                    string row = r.ToString(CultureInfo.InvariantCulture);
                    WriteRow(writer, r, new[] {
                        Number(i), Number(i % Workloads.LookupRows + 1),
                        Formula($"VLOOKUP(B{row},Lookup!$A$2:$B${lookupLast},2,FALSE)*A{row}"),
                        Formula($"IF(C{row}>1000,ROUND(C{row}/3,2),C{row}*2)"),
                        Formula($"INDEX(Lookup!$B$2:$B${lookupLast},MATCH(B{row},Lookup!$A$2:$A${lookupLast},0))"),
                        InlineText("Cat" + (i % Workloads.Categories).ToString(CultureInfo.InvariantCulture))
                    });
                }
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            WorksheetPart lookupPart = workbookPart.AddNewPart<WorksheetPart>();
            using (OpenXmlWriter writer = OpenXmlWriter.Create(lookupPart)) {
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());
                WriteRow(writer, 1, new[] { InlineText("Key"), InlineText("Value") });
                for (int key = 1; key <= Workloads.LookupRows; key++) WriteRow(writer, key + 1, new[] { Number(key), Number(key) });
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            WorksheetPart summaryPart = workbookPart.AddNewPart<WorksheetPart>();
            using (OpenXmlWriter writer = OpenXmlWriter.Create(summaryPart)) {
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());
                for (int category = 0; category < Workloads.Categories; category++) {
                    string row = (category + 1).ToString(CultureInfo.InvariantCulture);
                    WriteRow(writer, category + 1, new[] {
                        InlineText("Cat" + category.ToString(CultureInfo.InvariantCulture)),
                        Formula($"SUMIFS(Data!$D$2:$D${last},Data!$F$2:$F${last},A{row})")
                    });
                }
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            AddWorkbook(workbookPart, ("Data", dataPart), ("Lookup", lookupPart), ("Summary", summaryPart));
        }
        return output.ToArray();
    }

    private static void AddWorkbook(WorkbookPart workbookPart, params (string Name, WorksheetPart Part)[] sheets) {
        var list = new Sheets();
        uint id = 1;
        foreach ((string name, WorksheetPart part) in sheets) {
            list.Append(new Sheet { Name = name, SheetId = id++, Id = workbookPart.GetIdOfPart(part) });
        }
        workbookPart.Workbook = new Workbook(list);
    }

    private static void WriteRow(OpenXmlWriter writer, int rowIndex, Cell[] cells) {
        writer.WriteStartElement(new Row { RowIndex = (uint)rowIndex });
        for (int column = 0; column < cells.Length; column++) {
            cells[column].CellReference = ColumnName(column + 1) + rowIndex.ToString(CultureInfo.InvariantCulture);
            writer.WriteElement(cells[column]);
        }
        writer.WriteEndElement();
    }

    private static Cell Number(double value) => new Cell { CellValue = new CellValue(value.ToString("R", CultureInfo.InvariantCulture)) };
    private static Cell Formula(string formula) => new Cell { CellFormula = new CellFormula(formula) };
    private static Cell InlineText(string text) => new Cell { DataType = CellValues.InlineString, InlineString = new InlineString(new Text(text)) };

    internal static string ColumnName(int column) {
        string name = string.Empty;
        for (; column > 0; column = (column - 1) / 26) name = (char)('A' + (column - 1) % 26) + name;
        return name;
    }

    private static int ColumnOf(string reference) {
        int column = 0;
        foreach (char character in reference) {
            if (!char.IsLetter(character)) break;
            column = column * 26 + (char.ToUpperInvariant(character) - 'A' + 1);
        }
        return column;
    }
}
