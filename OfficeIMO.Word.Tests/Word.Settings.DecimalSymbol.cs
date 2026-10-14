using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(".")]
    [InlineData(",")]
    public void Decimal_symbol_survives_save_load_and_can_be_cleared_without_invalid_values(string value) {
        using var document = WordDocument.Create();
        document.AddParagraph("One\t123.45").AddTabStop(720, WordTabAlignment.Decimal);
        document.Settings.DecimalSymbol = value;
        foreach (string invalid in new[] { "", "..", "\t", "\uD800" })
            Assert.Throws<ArgumentException>(() => document.Settings.DecimalSymbol = invalid);
        Assert.Equal(value, document.Settings.DecimalSymbol);
        using var output = new MemoryStream(); document.Save(output); output.Position = 0;
        using (var xml = WordprocessingDocument.Open(output, false)) Assert.Empty(new OpenXmlValidator().Validate(xml));
        output.Position = 0;
        using var reopened = WordDocument.Load(output);
        Assert.Equal(value, reopened.Settings.DecimalSymbol);
        reopened.Settings.DecimalSymbol = null;
        using var cleared = new MemoryStream(); reopened.Save(cleared); cleared.Position = 0;
        using var without = WordDocument.Load(cleared);
        Assert.Null(without.Settings.DecimalSymbol);
    }
}
