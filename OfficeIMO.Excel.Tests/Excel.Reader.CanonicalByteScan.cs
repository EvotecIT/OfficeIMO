using System.Text;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    private const string ByteScanRows =
        "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\"><v>1.5</v></c></row>"
        + "<row r=\"2\"><c r=\"C2\" t=\"inlineStr\"><is><t>x</t></is></c></row>";

    /// <summary>
    /// The byte scan may only answer with the XmlReader scan's answer. Every worksheet below, and every
    /// truncated prefix of it, is measured both ways; the shapes the byte scan is meant to cover must answer.
    /// </summary>
    [Fact]
    public void CanonicalByteScan_AnswersOnlyWithTheXmlReaderScansExtent() {
        (string Xml, bool MustAnswer)[] cases = {
            (ByteScanSheet(ByteScanRows), true),
            ("﻿<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" + ByteScanSheet(ByteScanRows), true),
            (ByteScanSheet("<row r='1'>\r\n  <c r = 'A1' s='3'><v>2</v></c>\r\n</row>"), true),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\"><v/></c><c r=\"B1\"><v></v></c><c r=\"C1\" t=\"inlineStr\"><is><t/></is></c></row>"), true),
            (ByteScanSheet("<row r=\"1\"><c r=\"a1\"/><c r=\"xfd1\"/></row>"), true),
            (ByteScanSheet(ByteScanRows, rootAttributes: " xmlns:mc=\"urn:mc\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"urn:x14ac\"",
                after: "<extLst><ext uri=\"{1}\" xmlns:x14=\"urn:x14\"><x14:sparklineGroups x14ac:dyDescent=\"0.25\"/></ext></extLst>"), true),
            (ByteScanSheet(ByteScanRows, after: "<extLst><ext xmlns=\"urn:other\"><foo bar=\"1\"/></ext></extLst>"), true),

            (ByteScanSheet(ByteScanManyRows(null)), true),
            (ByteScanSheet(ByteScanManyRows("<!-- note -->")), false),
            (ByteScanSheet(ByteScanManyRows("<c r=\"B{0}\"><v> </v></c>")), false),
            (ByteScanSheet(ByteScanManyRows("<c r=\"B{1}\"/>")), false),
            (ByteScanSheet(ByteScanManyRows("<c r=\"B{0}\"><row r=\"{0}\"/></c>")), false),
            (ByteScanSheet(ByteScanManyRows("<c r=\"B{0}\"><is><t>a</t></is></c></row></sheetData><sheetData><row r=\"{0}\">")), false),
            (ByteScanSheet(ByteScanManyRows("<p:c r=\"B{0}\" xmlns:p=\"urn:p\"/>")), true),
            (ByteScanSheet(ByteScanManyRows(null) + "<row r=\"41\"><c r=\"A41\"/>"), false),
            (ByteScanSheet(ByteScanManyRows(null), after: "<extLst><sheetData><row r=\"9\"/></sheetData></extLst>"), false),
            (ByteScanSheet(ByteScanManyRows(null), after: "<extLst><ext xmlns:x14=\"urn:x14\"><x14:c r=\"A40\"/></ext></extLst>"), true),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\"><v> </v></c></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t xml:space=\"preserve\"> </t></is></c></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\"><v>\n<x/></v></c></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\"><v>&amp;</v></c></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\"><v>&#32;</v></c></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A&#49;\"/></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\" A1\"/></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\t\"/></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\" t=\"s \"/></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\" t=\"\"/></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\" t=\"x\"/></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\" r=\"B1\"/></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\" p:id=\"1\" q:id=\"2\"/></row>", rootAttributes: " xmlns:p=\"urn:a\" xmlns:q=\"urn:a\""), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\" xml:space=\"bogus\"/></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\"r=\"B1\"/></row>"), false),
            (ByteScanSheet("<row r=\"1\"><u:c r=\"A1\"/></row>"), false),
            (ByteScanSheet(ByteScanRows, rootAttributes: " xmlns:p=\"\""), false),
            ("<x:worksheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:sheetData>"
                + "<x:row r=\"1\"><x:c r=\"A1\"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>", false),
            (ByteScanSheet(ByteScanRows, after: "<extLst><ext xmlns:x14=\"urn:x14\"><x14:c r=\"ZZ9\"/></ext></extLst>"), false),
            (ByteScanSheet("<row><c r=\"A1\"/></row>"), false),
            (ByteScanSheet("<row/><row r=\"0\"/>"), false),
            (ByteScanSheet("<row r=\"1048577\"/>"), false),
            (ByteScanSheet("<row r=\"1\"><c><v>1</v></c></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A2\"/></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"XFE1\"/></row>"), false),
            (ByteScanSheet("<row r=\"1\"><!-- note --><c r=\"A1\"/></row>"), false),
            (ByteScanSheet("<row r=\"1\"><?pi data?><c r=\"A1\"/></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\" t=\"str\"><v><![CDATA[ ]]></v></c></row>"), false),
            ("<!DOCTYPE worksheet>" + ByteScanSheet(ByteScanRows), false),
            ("<?xml version=\"1.0\" encoding=\"windows-1252\"?>" + ByteScanSheet(ByteScanRows), false),
            (" <?xml version=\"1.0\"?>" + ByteScanSheet(ByteScanRows), false),
            (ByteScanSheet(ByteScanRows) + "x", false),
            (ByteScanSheet(ByteScanRows) + "<worksheet/>", false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\"><v>1</c></v></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\" t=\"str\"><v>\u0001</v></c></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\" t=\"str\"><v>￿</v></c></row>"), false),
            (ByteScanSheet("<row r=\"1\"><c r=\"A1\" t=\"str\"><v>a]]>b</v></c></row>"), false),
        };

        foreach ((string xml, bool mustAnswer) in cases) {
            byte[] document = Encoding.UTF8.GetBytes(xml);
            AssertByteScanAgrees(document, document.Length, xml, mustAnswer);
            for (int length = 0; length < document.Length; length++) {
                AssertByteScanAgrees(document, length, xml, mustAnswer: false);
            }
        }

        byte[] invalidUtf8 = Encoding.UTF8.GetBytes(ByteScanSheet("<row r=\"1\"><c r=\"A1\" t=\"str\"><v>#</v></c></row>"));
        invalidUtf8[Array.IndexOf(invalidUtf8, (byte)'#')] = 0xC0;
        AssertByteScanAgrees(invalidUtf8, invalidUtf8.Length, "invalid UTF-8", mustAnswer: false);
    }

    /// <summary>Forty rows; a marked row (the thirtieth) carries <paramref name="extra"/>, with {0} its number and {1} the next.</summary>
    private static string ByteScanManyRows(string? extra) {
        var rows = new StringBuilder();
        for (int row = 1; row <= 40; row++) {
            rows.Append("<row r=\"").Append(row).Append("\" spans=\"1:3\"><c r=\"A").Append(row).Append("\"><v>").Append(row).Append("</v></c>");
            if (row == 30 && extra != null) rows.AppendFormat(System.Globalization.CultureInfo.InvariantCulture, extra, row, row + 1);
            rows.Append("<c r=\"C").Append(row).Append("\" t=\"inlineStr\"><is><t>x</t></is></c></row>\r\n");
        }
        return rows.ToString();
    }

    private static string ByteScanSheet(string rows, string rootAttributes = "", string after = "") =>
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
        + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"" + rootAttributes + ">"
        + "<dimension ref=\"A1\"/><sheetData>" + rows + "</sheetData>" + after + "</worksheet>";

    private static void AssertByteScanAgrees(byte[] document, int length, string label, bool mustAnswer) {
        // Without helpers the rows are walked in order; with them, in chunks split at row tags.
        foreach (int helpers in new[] { 0, 2 }) {
            AssertByteScanAgrees(document, length, label, mustAnswer, helpers);
        }
    }

    private static void AssertByteScanAgrees(byte[] document, int length, string label, bool mustAnswer, int helpers) {
        bool answered = ExcelSheetReader.TryMeasureCanonicalWorksheet(document, length, helpers, CancellationToken.None,
            out WorksheetExtent bytes);
        if (mustAnswer) Assert.True(answered, $"The byte scan ({helpers} helpers) left this shape to the XmlReader scan: {label}");
        if (!answered) return;

        WorksheetExtent? scanned;
        try {
            using var stream = new MemoryStream(document, 0, length, writable: false);
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true
            });
            scanned = ExcelSheetReader.MeasureCanonicalWorksheet(reader, CancellationToken.None, out _);
        } catch (XmlException) {
            scanned = null;
        }
        Assert.True(scanned.HasValue, $"The byte scan qualified what the XmlReader scan does not ({length} bytes): {label}");
        Assert.Equal((scanned!.Value.Rows, scanned.Value.Columns), (bytes.Rows, bytes.Columns));
    }
}
