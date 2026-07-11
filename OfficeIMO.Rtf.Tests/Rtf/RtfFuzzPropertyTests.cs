using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using System.Globalization;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfFuzzPropertyTests {
    [Fact]
    public void Seeded_Valid_Rtf_Preserves_Every_Source_Character_Losslessly() {
        var random = new Random(0x52_54_46);
        for (int caseIndex = 0; caseIndex < 100; caseIndex++) {
            int parameter = random.Next(-10_000, 10_001);
            string text = EscapeText(RandomText(random, random.Next(4, 32)));
            string binary = RandomBinaryText(random, random.Next(0, 18));
            string rtf = "{\\rtf1\\ansi{\\*\\officeimofuzz\\value" + parameter.ToString(CultureInfo.InvariantCulture) +
                " " + text + "}{\\bin" + binary.Length.ToString(CultureInfo.InvariantCulture) + " " + binary +
                "}\\pard Case " + caseIndex.ToString(CultureInfo.InvariantCulture) + "\\par}\r\n";

            RtfReadResult result = RtfDocument.Read(rtf, RtfReadOptions.CreateUntrustedProfile());

            Assert.Equal(rtf, result.ToRtfLossless());
            Assert.Equal(rtf, result.EditLossless().ToRtf());
        }
    }

    [Fact]
    public void Seeded_Malformed_Groups_Normalize_Deterministically_With_Diagnostics() {
        var random = new Random(0x47_52_50);
        for (int caseIndex = 0; caseIndex < 64; caseIndex++) {
            string valid = "{\\rtf1\\ansi{\\b Bold " + caseIndex.ToString(CultureInfo.InvariantCulture) +
                "}{\\i Italic}\\pard Body\\par}";
            int[] bracePositions = valid.Select((character, index) => new { character, index })
                .Where(item => item.character == '{' || item.character == '}')
                .Select(item => item.index)
                .ToArray();
            string malformed = valid.Remove(bracePositions[random.Next(bracePositions.Length)], 1);

            RtfReadResult first = RtfDocument.Read(malformed, RtfReadOptions.CreateUntrustedProfile());
            string normalized = first.ToRtfLossless();
            RtfReadResult second = RtfDocument.Read(normalized, RtfReadOptions.CreateUntrustedProfile());

            Assert.NotEmpty(first.Diagnostics);
            Assert.Equal(normalized, second.ToRtfLossless());
        }
    }

    [Fact]
    public void Seeded_Control_Parameters_Never_Overflow_The_Reader() {
        var random = new Random(0x43_54_52_4C);
        for (int caseIndex = 0; caseIndex < 100; caseIndex++) {
            int digits = random.Next(1, 48);
            var parameter = new StringBuilder(caseIndex % 3 == 0 ? "-" : string.Empty);
            for (int index = 0; index < digits; index++) parameter.Append((char)('0' + random.Next(10)));
            string rtf = "{\\rtf1\\ansi\\fs" + parameter + " Body\\par}";

            RtfReadResult result = RtfDocument.Read(rtf, RtfReadOptions.CreateUntrustedProfile());

            Assert.Equal(rtf, result.ToRtfLossless());
            Assert.NotNull(result.Document);
        }
    }

    [Fact]
    public void Seeded_Unicode_Fallback_Lengths_Produce_Exactly_One_Unicode_Scalar() {
        int[] codePoints = { 0x00E9, 0x017C, 0x03A9, 0x0416, 0x20AC, 0x4E2D, 0x65E5, 0xD55C };
        var random = new Random(0x55_43);
        for (int caseIndex = 0; caseIndex < 80; caseIndex++) {
            int codePoint = codePoints[random.Next(codePoints.Length)];
            int fallbackLength = random.Next(0, 8);
            int signedValue = codePoint <= short.MaxValue ? codePoint : codePoint - 0x10000;
            string fallback = new string('?', fallbackLength);
            string expected = char.ConvertFromUtf32(codePoint) + "X";
            string rtf = "{\\rtf1\\ansi\\uc" + fallbackLength.ToString(CultureInfo.InvariantCulture) +
                "\\pard \\u" + signedValue.ToString(CultureInfo.InvariantCulture) + fallback + "X\\par}";

            RtfReadResult result = RtfDocument.Read(rtf, RtfReadOptions.CreateUntrustedProfile());

            Assert.Equal(expected, Assert.Single(result.Document.Paragraphs).ToPlainText());
            Assert.Equal(rtf, result.ToRtfLossless());
        }
    }

    [Fact]
    public void Seeded_Binary_Lengths_Preserve_Payload_And_Enforce_PerPayload_Limit() {
        var random = new Random(0x42_49_4E);
        for (int caseIndex = 0; caseIndex < 80; caseIndex++) {
            string payload = RandomBinaryText(random, random.Next(0, 64));
            string rtf = "{\\rtf1\\ansi{\\*\\payload\\bin" + payload.Length.ToString(CultureInfo.InvariantCulture) +
                " " + payload + "}\\pard Visible\\par}";
            RtfReadResult result = RtfDocument.Read(rtf, RtfReadOptions.CreateUntrustedProfile());
            Assert.Equal(rtf, result.ToRtfLossless());
        }

        var options = RtfReadOptions.CreateUntrustedProfile();
        options.MaxBinaryBytesPerPayload = 8;
        RtfReadLimitException exception = Assert.Throws<RtfReadLimitException>(() =>
            RtfDocument.Read("{\\rtf1\\ansi\\bin9 123456789}", options));
        Assert.Equal("RtfBinaryPayloadLimitExceeded", exception.Code);
        Assert.Equal(nameof(RtfReadOptions.MaxBinaryBytesPerPayload), exception.LimitSource);
    }

    [Fact]
    public void Seeded_Semantic_Documents_Retain_Text_And_Block_Shape_After_Normalization() {
        var random = new Random(0x53_45_4D);
        for (int caseIndex = 0; caseIndex < 40; caseIndex++) {
            RtfDocument source = RtfDocument.Create();
            int paragraphCount = random.Next(1, 12);
            for (int paragraphIndex = 0; paragraphIndex < paragraphCount; paragraphIndex++) {
                RtfParagraph paragraph = source.AddParagraph();
                paragraph.AddText(RandomText(random, random.Next(1, 24))).Bold = random.Next(2) == 0;
                paragraph.AddText(" żΩЖ中 ").Italic = random.Next(2) == 0;
                paragraph.AddText(RandomText(random, random.Next(1, 24))).Underline = random.Next(2) == 0;
                if (paragraphIndex % 5 == 0) paragraph.SetList(1, paragraphIndex % 3, RtfListKind.Bullet);
                if (paragraphIndex % 4 == 0) {
                    RtfTable table = source.AddTable(2, 2);
                    foreach (RtfTableRow row in table.Rows) {
                        foreach (RtfTableCell cell in row.Cells) cell.AddParagraph(RandomText(random, 8));
                    }
                }
            }

            string firstRtf = source.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
            RtfDocument first = RtfDocument.Read(firstRtf, RtfReadOptions.CreateUntrustedProfile()).Document;
            string secondRtf = first.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
            RtfDocument second = RtfDocument.Read(secondRtf, RtfReadOptions.CreateUntrustedProfile()).Document;

            Assert.Equal(first.Blocks.Select(BlockText), second.Blocks.Select(BlockText));
            Assert.Equal(first.Blocks.Select(block => block.GetType()), second.Blocks.Select(block => block.GetType()));
        }
    }

    private static string BlockText(IRtfBlock block) {
        if (block is RtfParagraph paragraph) return paragraph.ToPlainText();
        if (block is RtfTable table) {
            return string.Join("|", table.Rows.SelectMany(row => row.Cells).Select(cell =>
                string.Join(" ", cell.Paragraphs.Select(paragraph => paragraph.ToPlainText()))));
        }
        return string.Empty;
    }

    private static string RandomText(Random random, int length) {
        const string alphabet = "abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
        var builder = new StringBuilder(length);
        for (int index = 0; index < length; index++) builder.Append(alphabet[random.Next(alphabet.Length)]);
        return builder.ToString();
    }

    private static string EscapeText(string value) => value
        .Replace("\\", "\\\\")
        .Replace("{", "\\{")
        .Replace("}", "\\}");

    private static string RandomBinaryText(Random random, int length) {
        const string alphabet = "abcXYZ09{}\\";
        var builder = new StringBuilder(length);
        for (int index = 0; index < length; index++) builder.Append(alphabet[random.Next(alphabet.Length)]);
        return builder.ToString();
    }
}
