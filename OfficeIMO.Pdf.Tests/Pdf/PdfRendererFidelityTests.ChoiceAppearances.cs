using System;
using System.Collections.Generic;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfRendererFidelityTests {
    [Theory]
    [InlineData(0, false)]
    [InlineData(1, true)]
    public void TransientListBoxPaintsOptionRowsAndAllSelectedHighlights(int topIndex, bool inherited) {
        string field = $"/FT /Ch /Opt [(None) [(PL) (Poland)] [(US) (United States)]] /V [(PL) (US)] /I [1 2] /TI {topIndex} /Ff 2097152 /DA (/Helv 12 Tf)";
        var objects = new List<(string, byte[])> {
            ("<< /Type /Annot /Subtype /Widget /Rect [10 10 90 90] " + (inherited ? "/Parent 6 0 R" : field) + " >>", Array.Empty<byte>()),
            ("<< " + field + " >>", Array.Empty<byte>())
        };
        byte[] bytes = BuildPdf("", "<< >>", objects: objects, pageEntries: "/Annots [5 0 R]",
            catalogEntries: "/AcroForm << /Fields [5 0 R] >>");
        PdfReadDocument document = PdfReadDocument.Open(bytes);
        Assert.True(PdfFormFiller.TryCreateTransientWidgetAppearance(document.Objects,
            (PdfDictionary)document.Objects[5].Value, null, out PdfStream appearance));
        string content = Encoding.ASCII.GetString(appearance.Data);
        Assert.Contains("506F6C616E64", content);
        Assert.Contains("556E6974656420537461746573", content);
        Assert.Equal(topIndex == 0, content.Contains("4E6F6E65", StringComparison.Ordinal));
        Assert.Equal(2, content.Split(new[] { "0.153 0.392 0.8 rg" }, StringSplitOptions.None).Length - 1);
        // Exact glyph-edge assertions require a known face, not the host's
        // substitute for the widget's unembedded Helvetica resource.
        var exportOptions = new PdfImageExportOptions { TargetDpi = 72D };
        exportOptions.Fonts.Add("Helvetica", System.IO.File.ReadAllBytes(System.IO.Path.Combine(
            RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf")));
        var image = PdfRendererQualitySupport.Decode(document.Pages[0].ExportImage(
            OfficeIMO.Drawing.OfficeImageExportFormat.Png, exportOptions).Bytes);
        // Ascenders leave a clear pixel below the widget's inner clip rather than
        // intersecting its top edge. Check both unselected and selected first rows.
        for (int x = 15; x < 80; x++) {
            const int y = 12;
            if (topIndex == 0) AssertColor(image, x, y, 255, 255, 255);
            else AssertColor(image, x, y, 39, 100, 204);
        }
        AssertColor(image, 85, topIndex == 0 ? 30 : 16, 39, 100, 204);
        AssertColor(image, 85, topIndex == 0 ? 45 : 30, 39, 100, 204);
        AssertColor(image, 85, 80, 255, 255, 255);
    }

    [Theory]
    [InlineData("Ch", false)]
    [InlineData("Ch", true)]
    [InlineData("Tx", false)]
    public void IncrementalAppearanceEncodesInheritedType0RowsWithoutViewerRegeneration(string type, bool unicodeOption) {
        string mapping = "2 beginbfchar\n<0041> <0041>\n<1234> <03A9>\nendbfchar";
        string options = unicodeOption ? "/Opt [(A) <FEFF03A9>]" : "/Opt [(A)]";
        byte[] source = BuildPdf("", "<< >>", objects: new List<(string, byte[])> {
            ($"<< /Type /Annot /Subtype /Widget /FT /{type} /T (Selection) /Rect [10 10 90 90] /DA (/F1 12 Tf) /V (A) {options} >>", Array.Empty<byte>()),
            ("<< /Type /Font /Subtype /Type0 /BaseFont /Custom /Encoding /Identity-H /ToUnicode 7 0 R /DescendantFonts [<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Custom /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /DW 1000 /W [65 [600] 4660 [700]] >>] >>", Array.Empty<byte>()),
            ($"<< /Length {mapping.Length} >>", Encoding.ASCII.GetBytes(mapping)),
            ("<< /Fields [5 0 R] /DR << /Font << /F1 6 0 R >> >> >>", Array.Empty<byte>())
        }, catalogEntries: "/AcroForm 8 0 R", pageEntries: "/Annots [5 0 R]");
        source = PdfDocumentObjectGraphRewriter.Rewrite(source, sourceReadOptions: null, outputEncryption: null);
        PdfDocument updated = PdfDocument.Load(source).Forms.AppendRevision(new Dictionary<string, string> { ["Selection"] = "A" },
            new PdfIncrementalFormFieldUpdateOptions { GenerateAppearanceStreams = true, KeepNeedAppearances = false });
        PdfReadDocument document = PdfReadDocument.Open(updated.ToBytes());
        var widget = (PdfDictionary)document.Objects[Assert.Single(updated.Inspect().FormFields).ObjectNumber!.Value].Value;
        var normal = (PdfReference)((PdfDictionary)widget.Items["AP"]).Items["N"];
        string content = Encoding.ASCII.GetString(((PdfStream)document.Objects[normal.ObjectNumber].Value).Data);
        Assert.Contains("<0041>", content);
        Assert.DoesNotContain("<41>", content);
        if (unicodeOption) Assert.Contains("<1234>", content);
        Assert.False(updated.Inspect().AcroFormNeedAppearances);
        if (unicodeOption) {
            PdfDocument bounded = PdfDocument.Load(source, new PdfLoadOptions {
                Limits = new PdfReadLimits { MaxDecodedStreamBytes = 8 }
            });
            Assert.Throws<PdfReadLimitException>(() => bounded.Forms.AppendRevision(
                new Dictionary<string, string> { ["Selection"] = "A" },
                new PdfIncrementalFormFieldUpdateOptions { GenerateAppearanceStreams = true, KeepNeedAppearances = false }));
        }

    }

    [Fact]
    public void TransientListBoxUsesInheritedType0GlyphCodesAndPreservesReadLimits() {
        PdfReadDocument document = CreateTransientType0TextWidget("Identity-H", "(A)", 2097152, 0,
            "2 beginbfchar\n<0041> <0041>\n<1234> <03A9>\nendbfchar", "[65 [600] 4660 [700]]");
        var widget = (PdfDictionary)document.Objects[5].Value;
        widget.Items["FT"] = new PdfName("Ch");
        widget.Items["Rect"] = CreateChoiceArray(new PdfNumber(10), new PdfNumber(10), new PdfNumber(90), new PdfNumber(90));
        var options = new PdfArray();
        options.Items.Add(new PdfStringObj("A"));
        options.Items.Add(new PdfStringObj("Ω", useTextStringEncoding: true));
        widget.Items["Opt"] = options;
        widget.Items["V"] = CreateChoiceArray(new PdfStringObj("A"), new PdfStringObj("Ω", useTextStringEncoding: true));
        Assert.True(PdfFormFiller.TryCreateTransientWidgetAppearance(document.Objects, widget,
            (PdfDictionary)document.Objects[8].Value, out PdfStream appearance));
        string content = Encoding.ASCII.GetString(appearance.Data);
        Assert.Contains("<0041>", content);
        Assert.Contains("<1234>", content);
        Assert.DoesNotContain("<00411234>", content);
        Assert.Throws<PdfReadLimitException>(() => PdfFormFiller.TryCreateTransientWidgetAppearance(document.Objects, widget,
            (PdfDictionary)document.Objects[8].Value, out _, maximumDecodedBytes: 8));
        using var cancelled = new System.Threading.CancellationTokenSource();
        cancelled.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => PdfFormFiller.TryCreateTransientWidgetAppearance(document.Objects,
            widget, null, out _, cancellationToken: cancelled.Token));
    }

    [Fact]
    public void TransientComboBoxKeepsOnlyItsSelectedDisplayValue() {
        PdfReadDocument document = PdfReadDocument.Open(BuildPdf("", "<< >>", objects: new List<(string, byte[])> {
            ("<< /Type /Annot /Subtype /Widget /FT /Ch /Ff 131072 /Rect [10 60 90 80] /Opt [(None) [(PL) (Poland)]] /V (PL) >>", Array.Empty<byte>())
        }));
        Assert.True(PdfFormFiller.TryCreateTransientWidgetAppearance(document.Objects,
            (PdfDictionary)document.Objects[5].Value, null, out PdfStream appearance));
        string content = Encoding.ASCII.GetString(appearance.Data);
        Assert.Contains("506F6C616E64", content);
        Assert.DoesNotContain("4E6F6E65", content);
        Assert.DoesNotContain("0.153 0.392 0.8", content);
    }
    private static PdfArray CreateChoiceArray(params PdfObject[] items) {
        var array = new PdfArray();
        foreach (PdfObject item in items) array.Items.Add(item);
        return array;
    }

}
