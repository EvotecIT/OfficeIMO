using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFontResourceNamesTests {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(true, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void ConverterLookupKeysStaySeparateFromProgramNamesAndPublicNamesRemainExplicit(
        bool programNames, bool standardSlot, bool openTypeCff) {
        const string alias = "Private lookup family";
        byte[] data = openTypeCff
            ? File.ReadAllBytes(PdfComplianceTestFonts.FindBundledOpenTypeCffFont()
                ?? throw new InvalidOperationException("The bundled OpenType/CFF fixture is required."))
            : ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B', 'C');
        var family = new PdfEmbeddedFontFamily(alias, data);
        if (programNames) family = family.WithProgramFontNames();
        var options = new PdfCreateOptions();
        if (standardSlot) options.RegisterFontFamily(PdfStandardFont.Helvetica, family);
        else options.RegisterNamedFontFamily(family);
        using var document = PdfDocument.Create(options).Paragraph(paragraph => {
            if (!standardSlot) paragraph.FontFamily(alias);
            paragraph.Text("ABC");
        });

        using var independent = UglyToad.PdfPig.PdfDocument.Open(document.ToBytes());

        string expectedName = programNames ? PdfOpenTypeFontInspector.Inspect(data).FontName : alias.Replace(" ", "") + "-Regular";
        var letters = independent.GetPage(1).Letters;
        Assert.Equal("ABC", string.Concat(letters.Select(letter => letter.Value)));
        Assert.All(letters, letter => Assert.EndsWith(expectedName, letter.FontName, StringComparison.Ordinal));
        Assert.Equal(alias, family.FamilyName);
    }
}
