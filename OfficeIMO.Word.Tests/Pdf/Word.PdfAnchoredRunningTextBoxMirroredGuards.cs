using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_SymmetricMirroredRunningTextBoxRemainsQualified() =>
        AssertStableMirroredRunningTextBoxAdmission(1440, null);

    [Fact]
    public void SaveAsPdf_ExplicitMarginsKeepAsymmetricMirroredRunningTextBoxQualified() =>
        AssertStableMirroredRunningTextBoxAdmission(2880, PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(72D)));

    private static void AssertStableMirroredRunningTextBoxAdmission(uint authoredRightMargin,
        PdfCore.PdfPageMargins? exportMargins) {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        using WordDocument control = CreateAnchoredRunningTextBoxControl(includeTextBox: false);
        foreach (WordDocument source in new[] { document, control }) {
            source.Settings.MirrorMargins = true;
            WordSection selectingSection = source.Sections[1];
            selectingSection.Margins.Left = global::OfficeIMO.OfficeLength.Points(72D);
            selectingSection.Margins.Right = global::OfficeIMO.OfficeLength.Points((authoredRightMargin) / 20D);
        }
        string[] before = AnchoredRunningStoryXml(document);
        string[] controlBefore = AnchoredRunningStoryXml(control);
        string settingsBefore = document._wordprocessingDocument.MainDocumentPart!
            .DocumentSettingsPart!.Settings!.OuterXml;
        string controlSettingsBefore = control._wordprocessingDocument.MainDocumentPart!
            .DocumentSettingsPart!.Settings!.OuterXml;
        WordPdfExportOptions options = EmptyRunningOptions();
        options.Margins = exportMargins;
        WordPdfExportOptions controlOptions = EmptyRunningOptions();
        controlOptions.Margins = exportMargins;

        var result = document.ConvertToPdfWithEvidence(options);
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        using var controlPdf = PdfPigDocument.Open(control.ToPdf(controlOptions).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Equal(controlPdf.NumberOfPages, pdf.NumberOfPages);
        foreach (string token in new[] { "PREFIX", "SUFFIX" })
            Assert.Equal(AnchoredRunningLetters(controlPdf.GetPage(2), token).Select(letter => letter.StartBaseLine),
                AnchoredRunningLetters(pdf.GetPage(2), token).Select(letter => letter.StartBaseLine));
        Assert.Equal(72D, AnchoredRunningLetters(pdf.GetPage(2), "PREFIX")[0].StartBaseLine.X, 3);
        Assert.Equal(AnchoredRunningLetters(controlPdf.GetPage(2), "S2")[0].StartBaseLine.Y,
            AnchoredRunningLetters(pdf.GetPage(2), "S2")[0].StartBaseLine.Y, 3);
        AnchoredRunningLetters(pdf.GetPage(2), "BOX");
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.DoesNotContain("{sectionpages}", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
        Assert.Equal(controlBefore, AnchoredRunningStoryXml(control));
        Assert.Equal(settingsBefore, document._wordprocessingDocument.MainDocumentPart!
            .DocumentSettingsPart!.Settings!.OuterXml);
        Assert.Equal(controlSettingsBefore, control._wordprocessingDocument.MainDocumentPart!
            .DocumentSettingsPart!.Settings!.OuterXml);
    }
}
