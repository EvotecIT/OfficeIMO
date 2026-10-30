using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfRendererFidelityTests {
    [Fact]
    public void MultilineWidgetChargesPrefixMeasurementsBeforeRendering() {
        string value = string.Concat(Enumerable.Repeat("A ", 200));
        byte[] bytes = BuildPdf("", "<< >>", objects: new List<(string, byte[])> {
            ("<< /Type /Annot /Subtype /Widget /FT /Tx /T (Notes) /Ff 4096 /Rect [0 0 80 20] /DA (/Helv 12 Tf) /V (" + value + ") >>", Array.Empty<byte>())
        }, pageEntries: "/Annots [5 0 R]", catalogEntries: "/AcroForm << /Fields [5 0 R] >>");
        var readOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxDecodedTextCharacters = 1_000 } };
        PdfReadDocument document = PdfReadDocument.Open(bytes, readOptions);

        PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());
        Assert.Equal(PdfReadLimitKind.DecodedTextCharacters, error.Kind);
    }

    [Fact]
    public void WideMultilineWidgetKeepsFittingTextOnOneAppearanceLine() {
        string value = string.Concat(Enumerable.Repeat("A ", 8_192));
        string content = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(
            1_000_000_000D, 20D, value, 12D, new PdfFormFieldStyle { IsMultiline = true });

        Assert.Equal(1, content.Split(new[] { "BT /Helv" }, StringSplitOptions.None).Length - 1);
        Assert.Contains(" Tj ET", content, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MultilineWidgetMeasurementBudgetFollowsDocumentLimitsDuringFormUpdates(bool incremental) {
        string value = string.Concat(Enumerable.Repeat("A ", 200));
        byte[] bytes = BuildPdf("", "<< >>", objects: new List<(string, byte[])> {
            ("<< /Type /Annot /Subtype /Widget /FT /Tx /T (Notes) /Ff 4096 /Rect [0 0 80 20] /DA (/Helv 12 Tf) >>", Array.Empty<byte>())
        }, pageEntries: "/Annots [5 0 R]", catalogEntries: "/AcroForm << /Fields [5 0 R] >>");
        if (incremental)
            bytes = PdfDocumentObjectGraphRewriter.Rewrite(bytes, sourceReadOptions: null, outputEncryption: null);
        PdfDocument document = PdfDocument.Load(bytes, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxDecodedTextCharacters = 1_000 }
        });
        var values = new Dictionary<string, string> { ["Notes"] = value };

        PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() => {
            if (incremental)
                document.Forms.AppendRevision(values, new PdfIncrementalFormFieldUpdateOptions {
                    GenerateAppearanceStreams = true, KeepNeedAppearances = false
                });
            else document.Forms.Fill(values);
        });
        Assert.Equal(PdfReadLimitKind.DecodedTextCharacters, error.Kind);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MultilineWidgetMeasurementBudgetFollowsDocumentLimitsDuringFlatten(bool selective) {
        string value = string.Concat(Enumerable.Repeat("A ", 200));
        byte[] bytes = BuildPdf("", "<< >>", objects: new List<(string, byte[])> {
            ("<< /Type /Annot /Subtype /Widget /FT /Tx /T (Notes) /Ff 4096 /Rect [0 0 80 20] /DA (/Helv 12 Tf) /V (" + value + ") >>", Array.Empty<byte>())
        }, pageEntries: "/Annots [5 0 R]", catalogEntries: "/AcroForm << /Fields [5 0 R] >>");
        PdfDocument document = PdfDocument.Load(bytes, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxDecodedTextCharacters = 1_000 }
        });

        PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() => {
            if (selective) document.Forms.Flatten("Notes");
            else document.Forms.Flatten();
        });
        Assert.Equal(PdfReadLimitKind.DecodedTextCharacters, error.Kind);
    }
}
