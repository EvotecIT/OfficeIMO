using System;
using System.IO;
using System.Linq;
using Xunit;
using System.Text;
using OfficeIMO.Provenance;

namespace OfficeIMO.Shared.Tests;

public sealed class TextIntegrityReviewContracts {
    [Theory]
    [InlineData("utf8")]
    [InlineData("utf16le")]
    [InlineData("utf16be")]
    [InlineData("utf32le")]
    [InlineData("utf32be")]
    public void SelectedRemovalPreservesEncodingBomLineEndingsAndLanguage(string name) {
        Encoding encoding = name switch { "utf8" => new UTF8Encoding(true, true), "utf16le" => new UnicodeEncoding(false, true, true),
            "utf16be" => new UnicodeEncoding(true, true, true), "utf32le" => new UTF32Encoding(false, true, true), _ => new UTF32Encoding(true, true, true) };
        string text = "العربية\u200D❤️\u00A0\r\nreview\u202Ethis\r\n";
        byte[] data = encoding.GetPreamble().Concat(encoding.GetBytes(text)).ToArray();
        var review = OfficeTextIntegrityReview.Inspect(data);
        int selected = review.Report.Findings.ToList().FindIndex(item => item.CodePoint == 0x202E);
        byte[] expected = encoding.GetPreamble().Concat(encoding.GetBytes(text.Replace("\u202E", ""))).ToArray();
        Assert.Equal(expected, review.ExportSelected(review.Text, [selected]));
        Assert.Equal(text, review.Text);
        Assert.Equal(data, review.ExportSelected(review.Text, []));
        Assert.Equal(OfficeTextIntegrityReview.ComputeSha256(data), review.InputSha256);
    }
    [Fact]
    public void ChangedSourceAndInvalidSelectionsCannotReuseAReview() {
        var review = OfficeTextIntegrityReview.Inspect("a\u202Eb");
        Assert.Throws<InvalidOperationException>(() => review.RemoveSelected("x\u202Ey", [0]));
        Assert.Throws<ArgumentOutOfRangeException>(() => review.RemoveSelected(review.Text, [1]));
        Assert.Equal("ab", review.RemoveSelected(review.Text, [0, 0]));
        Assert.Throws<InvalidDataException>(() => OfficeTextIntegrityReview.Inspect(new byte[] { 0xC0, 0x80 }));
    }
}
