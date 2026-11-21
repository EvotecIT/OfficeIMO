using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Conditional formatting and differential formats that the native XLSB writers encode and the reader projects.</summary>
public sealed class ExcelXlsbConditionalFormattingTests {
    private static readonly OfficeColor Fill = OfficeColor.Parse("#FFC7CE");

    // Each case adds one rule to C3:D12 (data starts at A1) and checks the reloaded rule.
    private static readonly Dictionary<string, (Action<ExcelRangeConditionalFormattingBuilder> Add, Action<ExcelConditionalFormattingInfo> Check)> Cases = new() {
        ["cellIs between"] = (cf => cf.CellIs(ExcelConditionalFormattingOperator.Between, "3", "7", Fill), rule => {
            Assert.Equal("CellIs", rule.Type);
            Assert.Equal("Between", rule.Operator);
            Assert.Equal(new[] { "3", "7" }, rule.Formulas);
            Assert.Equal("FFFFC7CE", rule.DifferentialFillColorArgb);
        }),
        ["cellIs notEqual absolute stopIfTrue"] = (cf => cf.CellIs(ExcelConditionalFormattingOperator.NotEqual, "$A$1", null, Fill, stopIfTrue: true), rule => {
            Assert.Equal("NotEqual", rule.Operator);
            Assert.Equal(new[] { "$A$1" }, rule.Formulas);
            Assert.True(rule.StopIfTrue);
        }),
        ["cellIs lessThanOrEqual"] = (cf => cf.CellIs(ExcelConditionalFormattingOperator.LessThanOrEqual, "A1*2", null, Fill), rule => {
            Assert.Equal("LessThanOrEqual", rule.Operator);
            Assert.Equal(new[] { "A1*2" }, rule.Formulas);
        }),
        ["expression with relative and mixed references"] = (cf => cf.Formula("AND(C3>A3,$B3<D$1)", Fill), rule => {
            Assert.Equal("Expression", rule.Type);
            Assert.Equal(new[] { "AND(C3>A3,$B3<D$1)" }, rule.Formulas);
        }),
        ["containsText"] = (cf => cf.Text(ExcelConditionalFormatType.ContainsText, "ab", Fill), rule => {
            Assert.Equal("ContainsText", rule.Type);
            Assert.Equal("ab", rule.Text);
            Assert.Equal(new[] { "NOT(ISERROR(SEARCH(\"ab\",C3)))" }, rule.Formulas);
        }),
        ["notContainsText"] = (cf => cf.Text(ExcelConditionalFormatType.NotContainsText, "ab", Fill), rule => {
            Assert.Equal("NotContainsText", rule.Type);
            Assert.Equal("NotContains", rule.Operator);
            Assert.Equal("ab", rule.Text);
        }),
        ["beginsWith"] = (cf => cf.Text(ExcelConditionalFormatType.BeginsWith, "x", Fill), rule => {
            Assert.Equal("BeginsWith", rule.Type);
            Assert.Equal("x", rule.Text);
        }),
        ["endsWith"] = (cf => cf.Text(ExcelConditionalFormatType.EndsWith, "z", Fill), rule => {
            Assert.Equal("EndsWith", rule.Type);
            Assert.Equal("z", rule.Text);
        }),
        ["containsBlanks"] = (cf => cf.Blanks(true), rule => {
            Assert.Equal("ContainsBlanks", rule.Type);
            Assert.Equal(new[] { "LEN(TRIM(C3))=0" }, rule.Formulas);
            Assert.Null(rule.DifferentialFormatId);
        }),
        ["notContainsBlanks"] = (cf => cf.Blanks(false), rule => Assert.Equal("NotContainsBlanks", rule.Type)),
        ["containsErrors"] = (cf => cf.Errors(true), rule => Assert.Equal("ContainsErrors", rule.Type)),
        ["notContainsErrors"] = (cf => cf.Errors(false), rule => {
            Assert.Equal("NotContainsErrors", rule.Type);
            Assert.Equal(new[] { "NOT(ISERROR(C3))" }, rule.Formulas);
        }),
        ["timePeriod"] = (cf => cf.TimePeriod(ExcelConditionalTimePeriod.NextMonth, Fill), rule => {
            Assert.Equal("TimePeriod", rule.Type);
            Assert.Equal("NextMonth", rule.TimePeriod);
            Assert.Single(rule.Formulas);
        }),
        ["duplicateValues"] = (cf => cf.DuplicateValues(Fill), rule => Assert.Equal("DuplicateValues", rule.Type)),
        ["uniqueValues"] = (cf => cf.UniqueValues(Fill), rule => Assert.Equal("UniqueValues", rule.Type)),
        ["aboveAverage with standard deviations"] = (cf => cf.AboveAverage(true, false, 2, Fill), rule => {
            Assert.Equal("AboveAverage", rule.Type);
            Assert.True(rule.AboveAverageAbove);
            Assert.False(rule.AboveAverageEqual);
            Assert.Equal(2, rule.AboveAverageStdDev);
        }),
        ["below or equal average"] = (cf => cf.AboveAverage(false, true, null, Fill), rule => {
            Assert.False(rule.AboveAverageAbove);
            Assert.True(rule.AboveAverageEqual);
            Assert.Null(rule.AboveAverageStdDev);
        }),
        ["top items"] = (cf => cf.TopBottom(3, bottom: false, percent: false, Fill), rule => {
            Assert.Equal("Top10", rule.Type);
            Assert.Equal(3U, rule.TopBottomRank);
            Assert.False(rule.TopBottomBottom);
            Assert.False(rule.TopBottomPercent);
        }),
        ["bottom percent"] = (cf => cf.TopBottom(15, bottom: true, percent: true, Fill), rule => {
            Assert.Equal(15U, rule.TopBottomRank);
            Assert.True(rule.TopBottomBottom);
            Assert.True(rule.TopBottomPercent);
        }),
        ["two-color scale"] = (cf => cf.ColorScale(OfficeColor.Parse("#F8696B"), OfficeColor.Parse("#63BE7B")), rule => {
            Assert.Equal("ColorScale", rule.Type);
            Assert.Equal(new[] { "FFF8696B", "FF63BE7B" }, rule.ColorScaleColors);
            Assert.Equal(new[] { "min", "max" }, rule.ColorScaleThresholds.Select(threshold => threshold.Type));
        }),
        ["three-color scale"] = (cf => cf.ColorScale(OfficeColor.Parse("#F8696B"), OfficeColor.Parse("#FFEB84"), OfficeColor.Parse("#63BE7B")), rule => {
            Assert.Equal(3, rule.ColorScaleColors.Count);
            Assert.Equal("percentile", rule.ColorScaleThresholds[1].Type);
            Assert.Equal("50", rule.ColorScaleThresholds[1].Value);
        }),
        ["data bar"] = (cf => cf.DataBar(OfficeColor.Parse("#638EC6")), rule => {
            Assert.Equal("DataBar", rule.Type);
            Assert.Equal("FF638EC6", rule.DataBarColor);
            Assert.True(rule.DataBarShowValue);
            Assert.Equal(new[] { "min", "max" }, rule.DataBarThresholds.Select(threshold => threshold.Type));
        }),
        ["icon set"] = (cf => cf.IconSet(ExcelIconSet.FourArrows, showValue: false, reverseIconOrder: true, numberThresholds: new[] { 0D, 2.5, 5, 7.5 }), rule => {
            Assert.Equal("IconSet", rule.Type);
            Assert.Equal("FourArrows", rule.IconSet);
            Assert.False(rule.IconSetShowValue);
            Assert.True(rule.IconSetReverse);
            Assert.Equal(new[] { "0", "2.5", "5", "7.5" }, rule.IconSetThresholds.Select(threshold => threshold.Value));
            Assert.All(rule.IconSetThresholds, threshold => Assert.Equal("num", threshold.Type));
        }),
    };

    public static TheoryData<string> CaseNames => new(Cases.Keys);

    [Theory]
    [MemberData(nameof(CaseNames))]
    public void Conditional_formatting_rule_survives_native_xlsb_save(string name) {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = AddData(document);
            Cases[name].Add(sheet.Range("C3:D12").ConditionalFormat());
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx)));
        ExcelConditionalFormattingInfo rule = Assert.Single(reloaded["Data"].GetConditionalFormattingRules());
        Assert.Equal("C3:D12", rule.Range);
        Assert.Equal(1, rule.Priority);
        Cases[name].Check(rule);
    }

    [Fact]
    public void Rule_order_priorities_and_multi_area_ranges_survive_native_xlsb_save() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = AddData(document);
            sheet.Range("A1:A5").ConditionalFormat().CellIs(ExcelConditionalFormattingOperator.GreaterThan, "4", null, Fill, priority: 7);
            sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                Range = "B2:B4 D6:E8", Type = "expression", Formulas = new[] { "B2=D6" }, Priority = 3, DifferentialFillColorArgb = "FFFFEB9C"
            });
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx)));
        ExcelConditionalFormattingInfo[] rules = reloaded["Data"].GetConditionalFormattingRules().OrderBy(rule => rule.Priority).ToArray();
        Assert.Equal(new[] { "B2:B4 D6:E8", "A1:A5" }, rules.Select(rule => rule.Range));
        // Relative references are stored as offsets from B2, the top-left cell of the first area.
        Assert.Equal(new[] { "B2=D6" }, rules[0].Formulas);
        Assert.Equal("FFFFEB9C", rules[0].DifferentialFillColorArgb);
    }

    [Fact]
    public void Differential_format_properties_survive_native_xlsb_save() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = AddData(document);
            sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                Range = "A1:B4", Type = "expression", Formulas = new[] { "A1>2" },
                DifferentialFontBold = true, DifferentialFontItalic = true, DifferentialFontUnderline = true, DifferentialFontStrike = true,
                DifferentialFontColorArgb = "FF9C0006", DifferentialFontName = "Arial", DifferentialFontSize = OfficeLength.Points(14),
                DifferentialFillColorArgb = "FFFFEB9C",
                DifferentialBorder = new ExcelCellBorderInfo(left: new ExcelBorderSideInfo("thin", "FF0000FF"), top: new ExcelBorderSideInfo("double", "FF00B050"))
            });
            // Number format, alignment and protection have no rule-model setters; attach them to a second rule's dxf.
            Stylesheet styles = document.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!;
            styles.DifferentialFormats!.Append(new DifferentialFormat(
                new NumberingFormat { NumberFormatId = 170U, FormatCode = "0.000\"x\"" },
                new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Top, WrapText = true, Indent = 2U },
                new Protection { Locked = false, Hidden = true }));
            ExcelConditionalFormattingInfo second = sheet.Range("C1:C4").ConditionalFormat().Formula("C1<0");
            second.DifferentialFormatId = (uint)styles.DifferentialFormats.Elements<DifferentialFormat>().Count() - 1;
            sheet.UpdateConditionalFormattingRule(second);
        });

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(SaveAsXlsb(xlsx)));
        ExcelConditionalFormattingInfo[] rules = reloaded["Data"].GetConditionalFormattingRules().OrderBy(rule => rule.Range).ToArray();
        ExcelConditionalFormattingInfo styled = rules[0];
        Assert.True(styled.DifferentialFontBold);
        Assert.True(styled.DifferentialFontItalic);
        Assert.True(styled.DifferentialFontUnderline);
        Assert.True(styled.DifferentialFontStrike);
        Assert.Equal("FF9C0006", styled.DifferentialFontColorArgb);
        Assert.Equal("Arial", styled.DifferentialFontName);
        Assert.Equal(14D, styled.DifferentialFontSize!.Value.ToPoints(), 3);
        Assert.Equal("FFFFEB9C", styled.DifferentialFillColorArgb);
        Assert.Equal("thin", styled.DifferentialBorder!.Left!.Style);
        Assert.Equal("FF0000FF", styled.DifferentialBorder.Left.ColorArgb);
        Assert.Equal("double", styled.DifferentialBorder.Top!.Style);
        Assert.Equal("FF00B050", styled.DifferentialBorder.Top.ColorArgb);

        Stylesheet stylesheet = reloaded.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!;
        DifferentialFormat other = stylesheet.DifferentialFormats!.Elements<DifferentialFormat>().ElementAt((int)rules[1].DifferentialFormatId!.Value);
        Assert.Equal(170U, other.NumberingFormat!.NumberFormatId!.Value);
        Assert.Equal("0.000\"x\"", other.NumberingFormat.FormatCode!.Value);
        Assert.Equal(HorizontalAlignmentValues.Center, other.Alignment!.Horizontal!.Value);
        Assert.Equal(VerticalAlignmentValues.Top, other.Alignment.Vertical!.Value);
        Assert.True(other.Alignment.WrapText!.Value);
        Assert.Equal(2U, other.Alignment.Indent!.Value);
        Assert.False(other.Protection!.Locked!.Value);
        Assert.True(other.Protection.Hidden!.Value);
    }

    [Fact]
    public void Adding_a_rule_to_a_loaded_xlsb_rewrites_conditional_formatting_and_keeps_other_native_records() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = AddData(document);
            sheet.Range("A1:A10").ConditionalFormat().CellIs(ExcelConditionalFormattingOperator.GreaterThan, "5", null, Fill);
            sheet.Range("B1:B10").ConditionalFormat().DataBar(OfficeColor.Parse("#638EC6"));
            sheet.AddManualRowPageBreak(5, save: false);
        });
        byte[] source = SaveAsXlsb(xlsx);

        byte[] edited;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            loaded["Data"].Range("C2:C10").ConditionalFormat().Formula("C2>B2", OfficeColor.Parse("#C6EFCE"));
            edited = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(edited));
        ExcelConditionalFormattingInfo[] rules = reloaded["Data"].GetConditionalFormattingRules().OrderBy(rule => rule.Range).ToArray();
        Assert.Equal(new[] { "A1:A10", "B1:B10", "C2:C10" }, rules.Select(rule => rule.Range));
        Assert.Equal("FFFFC7CE", rules[0].DifferentialFillColorArgb);
        Assert.Equal("FF638EC6", rules[1].DataBarColor);
        Assert.Equal(new[] { "C2>B2" }, rules[2].Formulas);
        Assert.Equal("FFC6EFCE", rules[2].DifferentialFillColorArgb);
        // The page breaks and the source differential format are copied record for record.
        Assert.Equal(ReadRecords(source, "xl/worksheets/sheet1.bin", 392, 393), ReadRecords(edited, "xl/worksheets/sheet1.bin", 392, 393));
        Assert.Equal(ReadRecords(source, "xl/styles.bin", 507, 507), ReadRecords(edited, "xl/styles.bin", 507, 507).Take(1));
    }

    [Fact]
    public void Saving_a_loaded_xlsb_without_changes_keeps_its_conditional_formatting_records() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => {
            ExcelSheet sheet = AddData(document);
            sheet.Range("A1:A10").ConditionalFormat().IconSet(ExcelIconSet.ThreeArrows);
            sheet.Range("B1:B10").ConditionalFormat().Formula("B1>A1", Fill);
        }));

        byte[] saved;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            loaded["Data"].Cell(20, 1).SetValue("edited");
            saved = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }

        Assert.Equal(ReadRecords(source, "xl/worksheets/sheet1.bin", 461, 462, all: true), ReadRecords(saved, "xl/worksheets/sheet1.bin", 461, 462, all: true));
    }

    [Fact]
    public void Office_2010_extension_rules_are_refused_by_native_xlsb_save() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = AddData(document);
            sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                Source = ExcelConditionalFormattingSource.Office2010Extension,
                Range = "A1:A10", Type = "dataBar", DataBarColor = "FF638EC6", DataBarGradient = false,
                DataBarThresholds = new[] { new ExcelConditionalFormatThreshold { Type = "autoMin" }, new ExcelConditionalFormatThreshold { Type = "autoMax" } }
            });
        });

        Assert.Throws<NotSupportedException>(() => SaveAsXlsb(xlsx));
    }

    [Theory]
    [InlineData("rule")]
    [InlineData("threshold")]
    public void Array_constants_in_conditional_formatting_formulas_are_refused(string place) {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = AddData(document);
            ConditionalFormattingRule rule = place == "rule"
                ? new ConditionalFormattingRule(new Formula("SUM({1,2,3})>A1")) { Type = ConditionalFormatValues.Expression, Priority = 1 }
                : new ConditionalFormattingRule(new ColorScale(
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Formula, Val = "SUM({1,2})" },
                    new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max },
                    new Color { Rgb = "FFF8696B" }, new Color { Rgb = "FF63BE7B" })) { Type = ConditionalFormatValues.ColorScale, Priority = 1 };
            sheet.WorksheetPart.Worksheet!.Append(new ConditionalFormatting(rule) {
                SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = "A1:A10" }
            });
            sheet.EnsureWorksheetElementOrder();
        });

        NotSupportedException error = Assert.Throws<NotSupportedException>(() => SaveAsXlsb(xlsx));
        Assert.Contains("array constants", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Minimum_and_maximum_thresholds_ignore_their_formula_length_when_loading() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => AddData(document).Range("A1:A10").ConditionalFormat().DataBar(OfficeColor.Parse("#638EC6"))));
        // MS-XLSB 2.4.334 says cbFmla is ignored for CFVOMIN and CFVOMAX; state a nonzero length with no formula after it.
        byte[] patched = RewriteSheetRecords(source, record => record.Type == 471 && BitConverter.ToUInt32(record.Data, 0) is 2U or 3U
            ? PatchUInt32(record.Data, 20, 9U)
            : record.Data);

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(patched));

        ExcelConditionalFormattingInfo rule = Assert.Single(reloaded["Data"].GetConditionalFormattingRules());
        Assert.Equal("DataBar", rule.Type);
    }

    private static byte[] PatchUInt32(byte[] data, int offset, uint value) {
        byte[] copy = (byte[])data.Clone();
        BitConverter.GetBytes(value).CopyTo(copy, offset);
        return copy;
    }

    private static byte[] RewriteSheetRecords(byte[] xlsb, Func<XlsbRecord, byte[]> payload) {
        using var output = new MemoryStream();
        using (var source = new ZipArchive(new MemoryStream(xlsb), ZipArchiveMode.Read))
        using (var target = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ZipArchiveEntry entry in source.Entries) {
                using var content = new MemoryStream();
                using (Stream input = entry.Open()) input.CopyTo(content);
                byte[] bytes = content.ToArray();
                if (entry.FullName == "xl/worksheets/sheet1.bin") {
                    using var rewritten = new MemoryStream();
                    content.Position = 0;
                    foreach (XlsbRecord record in XlsbRecordReader.ReadAll(content)) XlsbRecordWriter.Write(rewritten, record.Type, payload(record));
                    bytes = rewritten.ToArray();
                }
                using Stream written = target.CreateEntry(entry.FullName).Open();
                written.Write(bytes, 0, bytes.Length);
            }
        }
        return output.ToArray();
    }

    private static ExcelSheet AddData(ExcelDocument document) {
        ExcelSheet sheet = document.AddSheet("Data");
        for (int row = 1; row <= 12; row++) {
            for (int column = 1; column <= 5; column++) sheet.Cell(row, column).SetValue(row * column);
        }
        return sheet;
    }

    private static byte[] CreateXlsx(Action<ExcelDocument> build) {
        using ExcelDocument document = ExcelDocument.Create();
        build(document);
        return document.ToBytes();
    }

    private static byte[] SaveAsXlsb(byte[] xlsx) {
        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(xlsx));
        return loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
    }

    private static string[] ReadRecords(byte[] xlsb, string partName, int beginType, int endType, bool all = false) {
        using var archive = new ZipArchive(new MemoryStream(xlsb), ZipArchiveMode.Read);
        using var part = new MemoryStream();
        using (Stream entry = archive.GetEntry(partName)!.Open()) entry.CopyTo(part);
        part.Position = 0;
        var records = XlsbRecordReader.ReadAll(part).ToList();
        int begin = records.FindIndex(record => record.Type == beginType);
        int end = all ? records.FindLastIndex(record => record.Type == endType) : records.FindIndex(record => record.Type == endType);
        Assert.True(begin >= 0 && end >= begin);
        return records.Skip(begin).Take(end - begin + 1).Select(record => record.Type + ":" + Convert.ToBase64String(record.Data)).ToArray();
    }
}
