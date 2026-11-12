using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

/// <summary>
/// Reopens each output with the Open XML SDK and checks that it holds what the lane wrote, so a library cannot win
/// by doing less work. Returns a short summary; throws when the output is wrong.
/// </summary>
internal static class Validation {
    internal static string Check(string lane, byte[] output, byte[]? input) {
        using var document = SpreadsheetDocument.Open(new MemoryStream(output, writable: false), isEditable: false);
        WorkbookPart workbook = document.WorkbookPart ?? throw Fail("no workbook part");
        return lane switch {
            "write1m" => CheckLarge(workbook),
            "styles" => CheckStyles(workbook),
            "merged" => Expect("merges", FirstSheet(workbook).Worksheet.Descendants<MergeCell>().Count(), Workloads.MergedRows),
            "cfdv" => CheckRules(workbook),
            "namestables" => CheckNamesAndTables(workbook),
            "autofit" => CheckWidths(workbook),
            "preserve" => CheckPreserved(workbook, input!),
            _ => throw Fail("unknown lane " + lane)
        };
    }

    private static string CheckLarge(WorkbookPart workbook) {
        long rows = 0, idSum = 0, flags = 0;
        using (OpenXmlReader reader = OpenXmlReader.Create(FirstSheet(workbook))) {
            while (reader.Read()) {
                if (reader.ElementType == typeof(Row) && reader.IsStartElement) rows++;
                if (reader.ElementType != typeof(Cell) || !reader.IsStartElement) continue;
                var cell = (Cell)reader.LoadCurrentElement()!;
                string reference = cell.CellReference?.Value ?? string.Empty;
                if (reference.StartsWith("A", StringComparison.Ordinal) && reference != "A1") idSum += long.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture);
                if (reference.StartsWith("D", StringComparison.Ordinal) && cell.DataType?.Value == CellValues.Boolean && cell.CellValue!.Text == "1") flags++;
            }
        }
        long n = Workloads.LargeRows;
        if (rows != n + 1 || idSum != n * (n + 1) / 2 || flags != n / 3) throw Fail($"rows {rows}, id sum {idSum}, true flags {flags}");
        return $"rows {rows}";
    }

    private static string CheckStyles(WorkbookPart workbook) {
        var used = new HashSet<uint>();
        long cells = 0;
        using (OpenXmlReader reader = OpenXmlReader.Create(FirstSheet(workbook))) {
            while (reader.Read()) {
                if (reader.ElementType != typeof(Cell) || !reader.IsStartElement) continue;
                var cell = (Cell)reader.LoadCurrentElement()!;
                cells++;
                used.Add(cell.StyleIndex?.Value ?? 0);
            }
        }
        if (cells != (long)Workloads.StyleRows * Workloads.StyleColumns || used.Count < Workloads.StyleCount) throw Fail($"cells {cells}, distinct styles {used.Count}");
        return $"distinct styles {used.Count}";
    }

    private static string CheckRules(WorkbookPart workbook) {
        Worksheet sheet = FirstSheet(workbook).Worksheet;
        // A library may write one rule per block or fold identical adjacent rules into one rule over a longer range, so
        // count the cells each kind of rule covers rather than the rule elements.
        int rules = sheet.Descendants<ConditionalFormatting>()
            .Where(formatting => formatting.Elements<ConditionalFormattingRule>().Any(rule => rule.Operator?.Value == ConditionalFormattingOperatorValues.GreaterThan))
            .Sum(formatting => CellCount(formatting.SequenceOfReferences));
        int validations = sheet.Descendants<DataValidation>()
            .Where(validation => validation.Type?.Value == DataValidationValues.Whole)
            .Sum(validation => CellCount(validation.SequenceOfReferences));
        if (rules != Workloads.RuleRows || validations != Workloads.RuleRows) throw Fail($"rule cells {rules}, validation cells {validations}");
        return $"rule cells {rules}, validation cells {validations}";
    }

    private static string CheckNamesAndTables(WorkbookPart workbook) {
        int names = workbook.Workbook.DefinedNames?.Elements<DefinedName>().Count(name => name.Name?.Value?.StartsWith("Name", StringComparison.Ordinal) == true) ?? 0;
        int tables = workbook.WorksheetParts.Sum(part => part.TableDefinitionParts.Count());
        if (names != Workloads.DefinedNames || tables != Workloads.Tables) throw Fail($"names {names}, tables {tables}");
        return $"names {names}, tables {tables}";
    }

    private static string CheckWidths(WorkbookPart workbook) {
        var widths = new double[Workloads.AutoFitColumns + 1];
        foreach (Column column in FirstSheet(workbook).Worksheet.Descendants<Column>()) {
            for (uint index = column.Min!.Value; index <= column.Max!.Value && index <= Workloads.AutoFitColumns; index++) widths[index] = column.Width?.Value ?? 0;
        }
        if (widths.Skip(1).Any(width => width <= 0)) throw Fail("columns without a width: " + string.Join(",", widths.Skip(1)));
        return "widths " + string.Join("/", widths.Skip(1).Select(width => width.ToString("0.#", CultureInfo.InvariantCulture)));
    }

    private static string CheckPreserved(WorkbookPart workbook, byte[] input) {
        using var original = SpreadsheetDocument.Open(new MemoryStream(input, writable: false), isEditable: false);
        string before = PartCounts(original.WorkbookPart!);
        string after = PartCounts(workbook);
        if (before != after) throw Fail($"parts before {before}, after {after}");
        string? text = CellText(workbook, FirstSheet(workbook), SdkLanes.ColumnName(Workloads.PreserveColumn) + Workloads.PreserveRow.ToString(CultureInfo.InvariantCulture));
        if (text != Workloads.PreserveText) throw Fail("edited cell holds " + (text ?? "nothing"));
        return after;
    }

    private static string PartCounts(WorkbookPart workbook) {
        IEnumerable<OpenXmlPart> parts = AllParts(workbook);
        int Count<T>() => parts.OfType<T>().Count();
        return $"charts {Count<ChartPart>()}, images {Count<ImagePart>()}, pivot caches {Count<PivotTableCacheDefinitionPart>()}, pivot tables {Count<PivotTablePart>()}, drawings {Count<DrawingsPart>()}";
    }

    private static IEnumerable<OpenXmlPart> AllParts(OpenXmlPart root) {
        var seen = new HashSet<OpenXmlPart>();
        var pending = new Stack<OpenXmlPart>();
        pending.Push(root);
        while (pending.Count > 0) {
            OpenXmlPart part = pending.Pop();
            if (!seen.Add(part)) continue;
            foreach (IdPartPair child in part.Parts) pending.Push(child.OpenXmlPart);
        }
        return seen;
    }

    private static string? CellText(WorkbookPart workbook, WorksheetPart sheet, string reference) {
        Cell? cell = sheet.Worksheet.Descendants<Cell>().FirstOrDefault(candidate => candidate.CellReference?.Value == reference);
        if (cell == null) return null;
        if (cell.DataType?.Value == CellValues.SharedString) {
            int index = int.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture);
            return workbook.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(index).InnerText;
        }
        return cell.DataType?.Value == CellValues.InlineString ? cell.InlineString?.InnerText : cell.CellValue?.Text;
    }

    private static WorksheetPart FirstSheet(WorkbookPart workbook) {
        Sheet first = workbook.Workbook.Sheets!.Elements<Sheet>().First();
        return (WorksheetPart)workbook.GetPartById(first.Id!);
    }

    private static int CellCount(ListValue<StringValue>? references) {
        int cells = 0;
        foreach (StringValue item in references?.Items ?? Enumerable.Empty<StringValue>()) {
            foreach (string range in (item.Value ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries)) {
                string[] ends = range.Split(':');
                (int firstRow, int firstColumn) = Position(ends[0]);
                (int lastRow, int lastColumn) = Position(ends[^1]);
                cells += (lastRow - firstRow + 1) * (lastColumn - firstColumn + 1);
            }
        }
        return cells;
    }

    private static (int Row, int Column) Position(string reference) {
        int column = 0, index = 0;
        while (index < reference.Length && char.IsLetter(reference[index])) column = column * 26 + (char.ToUpperInvariant(reference[index++]) - 'A' + 1);
        return (int.Parse(reference.AsSpan(index), NumberStyles.None, CultureInfo.InvariantCulture), column);
    }

    private static string Expect(string what, int actual, int expected) =>
        actual == expected ? $"{what} {actual}" : throw Fail($"{what} {actual}, expected {expected}");

    private static InvalidDataException Fail(string message) => new InvalidDataException("Validation failed: " + message);
}
