using OfficeIMO.ContentSafety;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfHiddenContentInspectionTests {
    [Fact]
    public void ContentSafetySkipsHiddenLayerReparseWhenDocumentHasNoOptionalContent() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("VISIBLE-CONTENT"))
            .ToBytes();
        int hiddenInspectionCalls = 0;
        PdfReadPage.HiddenOptionalContentInspectionObserverForTesting = () => hiddenInspectionCalls++;

        try {
            OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(pdf);

            Assert.Equal(0, hiddenInspectionCalls);
            Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Contains("Optional-content", StringComparison.Ordinal));
        } finally {
            PdfReadPage.HiddenOptionalContentInspectionObserverForTesting = null;
        }
    }

    [Fact]
    public void ContentSafetySurfacesHiddenLayerTextIncludingHiddenFormContent() {
        byte[] pdf = BuildHiddenOptionalContentPdf();

        string visibleText = PdfDocument.Load(pdf).Read().Text;
        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(pdf);
        OfficeContentSafetyFinding[] hidden = report.Findings
            .Where(finding => finding.Kind == OfficeContentConcealmentKind.HiddenContainer)
            .ToArray();

        Assert.Contains("VISIBLE-CONTENT", visibleText, StringComparison.Ordinal);
        Assert.DoesNotContain("HIDDEN-PAGE-CONTENT", visibleText, StringComparison.Ordinal);
        Assert.DoesNotContain("HIDDEN-FORM-CONTENT", visibleText, StringComparison.Ordinal);
        Assert.Contains(hidden, finding => finding.TextPreview.Contains("HIDDEN-PAGE-CONTENT", StringComparison.Ordinal));
        Assert.Contains(hidden, finding => finding.TextPreview.Contains("HIDDEN-FORM-CONTENT", StringComparison.Ordinal));
        Assert.All(hidden, finding => Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability));
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Contains("default layer configuration", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetyTreatsUnsupportedOptionalContentViewIntentAsInconclusive() {
        byte[] pdf = BuildHiddenOptionalContentPdf(unsupportedViewIntent: true);

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(pdf);

        Assert.DoesNotContain(
            report.Findings,
            finding => finding.Kind == OfficeContentConcealmentKind.HiddenContainer &&
                finding.TextPreview.Contains("HIDDEN-", StringComparison.Ordinal));
        Assert.Contains(
            report.Diagnostics,
            diagnostic => diagnostic.Contains("inconclusive", StringComparison.OrdinalIgnoreCase) &&
                diagnostic.Contains("view intent", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ContentSafetyExcludesAnnotationsAndFormValuesWhenNonPrimaryContentIsDisabled() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 8 0 R >> >> /Contents 4 0 R /Annots [5 0 R 7 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "5 0 obj\n<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (HIDDEN-ANNOTATION) /F 2 >>\nendobj",
            "6 0 obj\n<< /FT /Tx /T (HiddenField) /V (HIDDEN-FIELD-VALUE) /Kids [7 0 R] >>\nendobj",
            "7 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [60 20 180 40] /P 3 0 R /F 2 >>\nendobj",
            "8 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(
            Encoding.ASCII.GetBytes(pdf),
            new OfficeContentSafetyOptions { IncludeNonPrimaryContent = false });
        OfficeContentSafetyReport layered = PdfDocument.InspectContentSafety(
            BuildHiddenOptionalContentPdf(),
            new OfficeContentSafetyOptions { IncludeNonPrimaryContent = false });

        Assert.DoesNotContain(report.Findings, finding => finding.TextPreview.Contains("HIDDEN-ANNOTATION", StringComparison.Ordinal));
        Assert.DoesNotContain(report.Findings, finding => finding.TextPreview.Contains("HIDDEN-FIELD-VALUE", StringComparison.Ordinal));
        Assert.Contains(
            layered.Findings,
            finding => finding.Kind == OfficeContentConcealmentKind.HiddenContainer &&
                finding.TextPreview.Contains("HIDDEN-PAGE-CONTENT", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetySurfacesHiddenAnnotationAndWidgetValues() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 8 0 R >> >> /Contents 4 0 R /Annots [5 0 R 7 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "5 0 obj\n<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (HIDDEN-ANNOTATION) /F 2 >>\nendobj",
            "6 0 obj\n<< /FT /Tx /T (HiddenField) /V (HIDDEN-FIELD-VALUE) /Kids [7 0 R] >>\nendobj",
            "7 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [60 20 180 40] /P 3 0 R /F 2 >>\nendobj",
            "8 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        Assert.Contains(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.Contains("HIDDEN-ANNOTATION", StringComparison.Ordinal));
        Assert.Contains(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.Contains("HIDDEN-FIELD-VALUE", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetyDoesNotClassifyValueAsHiddenWhenAnotherWidgetIsVisible() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 9 0 R >> >> /Contents 4 0 R /Annots [7 0 R 8 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "6 0 obj\n<< /FT /Tx /T (SharedField) /V (VISIBLE-SHARED-VALUE) /Kids [7 0 R 8 0 R] >>\nendobj",
            "7 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [20 20 100 40] /P 3 0 R /F 2 >>\nendobj",
            "8 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [120 20 220 40] /P 3 0 R /F 4 >>\nendobj",
            "9 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 10 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        Assert.DoesNotContain(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.Contains("VISIBLE-SHARED-VALUE", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetyTreatsZeroAreaWidgetAsConcealed() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 8 0 R >> >> /Contents 4 0 R /Annots [7 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "6 0 obj\n<< /FT /Tx /T (ZeroAreaField) /V (ZERO-AREA-SECRET) /Kids [7 0 R] >>\nendobj",
            "7 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [20 20 20 40] /P 3 0 R /F 4 >>\nendobj",
            "8 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        Assert.Contains(report.Findings, finding =>
            finding.Location.EndsWith("/HiddenWidgetValue", StringComparison.Ordinal) &&
            finding.TextPreview.Contains("ZERO-AREA-SECRET", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetySurfacesChoiceExportValueBehindDisplayText() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 8 0 R >> >> /Contents 4 0 R /Annots [7 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "6 0 obj\n<< /FT /Ch /T (ChoiceField) /V (SECRET-CODE) /Opt [[(SECRET-CODE) (Public label)]] /Kids [7 0 R] >>\nendobj",
            "7 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [20 20 220 40] /P 3 0 R /F 4 >>\nendobj",
            "8 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        Assert.Contains(report.Findings, finding =>
            finding.Location.EndsWith("/HiddenChoiceExportValue", StringComparison.Ordinal) &&
            finding.TextPreview.Contains("SECRET-CODE", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetySurfacesIndependentRichFormValue() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 8 0 R >> >> /Contents 4 0 R /Annots [7 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "6 0 obj\n<< /FT /Tx /Ff 33554432 /T (RichField) /V (Public value) /RV (<body><p>RICH-SECRET</p></body>) /Kids [7 0 R] >>\nendobj",
            "7 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [20 20 220 40] /P 3 0 R /F 4 >>\nendobj",
            "8 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        byte[] bytes = Encoding.ASCII.GetBytes(pdf);
        PdfFormField field = Assert.Single(PdfReadDocument.Open(bytes).FormFields);
        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(bytes);

        Assert.True(field.HasRichValue);
        Assert.Contains("RICH-SECRET", field.RichValuePlainText, StringComparison.Ordinal);
        Assert.Contains(report.Findings, finding =>
            finding.Location.EndsWith("/HiddenRichValue", StringComparison.Ordinal) &&
            finding.TextPreview.Contains("RICH-SECRET", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetyDoesNotClassifyEquivalentRichFormValueAsHidden() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Contents 4 0 R /Annots [7 0 R] >>\nendobj",
            StreamObject(4, string.Empty, string.Empty),
            "6 0 obj\n<< /FT /Tx /Ff 33554432 /T (RichField) /V (Public value) /RV (<body><p>Public value</p></body>) /Kids [7 0 R] >>\nendobj",
            "7 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [20 20 220 40] /P 3 0 R /F 4 >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        Assert.DoesNotContain(report.Findings, finding =>
            finding.Location.EndsWith("/HiddenRichValue", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetyTreatsOffPageAnnotationAsConcealed() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /CropBox [10 20 210 140] /MediaBox [0 0 240 180] /Contents 4 0 R /Annots [5 0 R] >>\nendobj",
            StreamObject(4, string.Empty, string.Empty),
            "5 0 obj\n<< /Type /Annot /Subtype /Text /Rect [220 30 235 45] /Contents (OFF-PAGE-SECRET) /F 4 >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        Assert.Contains(report.Findings, finding =>
            finding.Location.Contains("/HiddenAnnotation[", StringComparison.Ordinal) &&
            finding.TextPreview.Contains("OFF-PAGE-SECRET", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetySurfacesDistinctStoredDefaultBehindVisibleCurrentValue() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 8 0 R >> >> /Contents 4 0 R /Annots [7 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "6 0 obj\n<< /FT /Tx /T (VisibleField) /V (PUBLIC-CURRENT-VALUE) /DV (SENSITIVE-RESET-DEFAULT) /Kids [7 0 R] >>\nendobj",
            "7 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [20 20 220 40] /P 3 0 R /F 4 >>\nendobj",
            "8 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        Assert.Contains(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.Location.EndsWith("/HiddenWidgetDefaultValue", StringComparison.Ordinal) &&
            finding.TextPreview.Contains("SENSITIVE-RESET-DEFAULT", StringComparison.Ordinal));
        Assert.DoesNotContain(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.Contains("PUBLIC-CURRENT-VALUE", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetyDoesNotClassifySoleDefaultAsHiddenWhenItsWidgetIsVisible() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 8 0 R >> >> /Contents 4 0 R /Annots [7 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "6 0 obj\n<< /FT /Tx /T (DefaultOnlyField) /DV (VISIBLE-INITIAL-DEFAULT) /Kids [7 0 R] >>\nendobj",
            "7 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [20 20 220 40] /P 3 0 R /F 4 >>\nendobj",
            "8 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        Assert.DoesNotContain(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.Contains("VISIBLE-INITIAL-DEFAULT", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("7 0 R 8 0 R")]
    [InlineData("8 0 R 7 0 R")]
    public void ContentSafetyReportsInheritedDefaultIndependentlyOfTerminalFieldOrder(string childReferences) {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 11 0 R >> >> /Contents 4 0 R /Annots [9 0 R 10 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "6 0 obj\n<< /FT /Tx /T (Shared) /DV (INHERITED-SECRET-DEFAULT) /Kids [" + childReferences + "] >>\nendobj",
            "7 0 obj\n<< /T (HiddenChild) /Kids [9 0 R] >>\nendobj",
            "8 0 obj\n<< /T (VisibleChild) /V (PUBLIC-CURRENT-VALUE) /Kids [10 0 R] >>\nendobj",
            "9 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 20 100 40] /P 3 0 R /F 2 >>\nendobj",
            "10 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 8 0 R /Rect [120 20 220 40] /P 3 0 R /F 4 >>\nendobj",
            "11 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 12 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        OfficeContentSafetyFinding finding = Assert.Single(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            item.TextPreview.Contains("INHERITED-SECRET-DEFAULT", StringComparison.Ordinal));
        Assert.EndsWith("/HiddenWidgetDefaultValue", finding.Location, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            item.TextPreview.Contains("PUBLIC-CURRENT-VALUE", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetyReportsInheritedDefaultWhenVisibleSiblingExplicitlySelectsNoValues() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 11 0 R >> >> /Contents 4 0 R /Annots [9 0 R 10 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "6 0 obj\n<< /FT /Ch /Ff 2097152 /T (Shared) /DV [(INHERITED-SECRET-DEFAULT)] /Kids [7 0 R 8 0 R] >>\nendobj",
            "7 0 obj\n<< /T (HiddenChild) /Kids [9 0 R] >>\nendobj",
            "8 0 obj\n<< /T (VisibleChild) /V [] /Kids [10 0 R] >>\nendobj",
            "9 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 20 100 40] /P 3 0 R /F 2 >>\nendobj",
            "10 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 8 0 R /Rect [120 20 220 40] /P 3 0 R /F 4 >>\nendobj",
            "11 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 12 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        OfficeContentSafetyFinding finding = Assert.Single(report.Findings, item =>
            item.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            item.TextPreview.Contains("INHERITED-SECRET-DEFAULT", StringComparison.Ordinal));
        Assert.EndsWith("/HiddenWidgetDefaultValue", finding.Location, StringComparison.Ordinal);
    }

    [Fact]
    public void ContentSafetySurfacesWidgetlessFieldValues() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>\nendobj",
            StreamObject(4, string.Empty, content),
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "6 0 obj\n<< /FT /Tx /T (DataOnlyField) /V (WIDGETLESS-SECRET) /DV (WIDGETLESS-DEFAULT) >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        Assert.Contains(report.Findings, finding => finding.TextPreview.Contains("WIDGETLESS-SECRET", StringComparison.Ordinal));
        Assert.Contains(report.Findings, finding => finding.TextPreview.Contains("WIDGETLESS-DEFAULT", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetyTreatsWidgetMissingFromPageAnnotationsAsConcealed() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>\nendobj",
            StreamObject(4, string.Empty, content),
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "6 0 obj\n<< /FT /Tx /T (DetachedWidgetField) /V (DETACHED-WIDGET-SECRET) /Kids [7 0 R] >>\nendobj",
            "7 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [20 20 120 40] /P 3 0 R /F 4 >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        Assert.Contains(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.Contains("DETACHED-WIDGET-SECRET", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetyAggregatesVisibilityAcrossTerminalFieldsSharingInheritedValue() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R /Annots [9 0 R 10 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "6 0 obj\n<< /FT /Tx /T (Shared) /V (INHERITED-SHARED-VALUE) /Kids [7 0 R 8 0 R] >>\nendobj",
            "7 0 obj\n<< /T (HiddenChild) /Kids [9 0 R] >>\nendobj",
            "8 0 obj\n<< /T (VisibleChild) /Kids [10 0 R] >>\nendobj",
            "9 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 20 100 40] /P 3 0 R /F 2 >>\nendobj",
            "10 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 8 0 R /Rect [120 20 220 40] /P 3 0 R /F 4 >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 11 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        Assert.DoesNotContain(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.Contains("INHERITED-SHARED-VALUE", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetyAggregatesVisibilityAcrossDirectValueOwnerChildren() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [<< /FT /Tx /T (Shared) /V (DIRECT-OWNER-VALUE) /Kids [7 0 R 8 0 R] >>] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R /Annots [9 0 R 10 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "7 0 obj\n<< /T (HiddenChild) /Kids [9 0 R] >>\nendobj",
            "8 0 obj\n<< /T (VisibleChild) /Kids [10 0 R] >>\nendobj",
            "9 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 20 100 40] /P 3 0 R /F 2 >>\nendobj",
            "10 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 8 0 R /Rect [120 20 220 40] /P 3 0 R /F 4 >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 11 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        Assert.DoesNotContain(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.Contains("DIRECT-OWNER-VALUE", StringComparison.Ordinal));
    }

    [Fact]
    public void ContentSafetyTreatsPasswordWidgetValueAsConcealed() {
        byte[] pdf = BuildSingleWidgetFieldPdf(
            fieldEntries: "/FT /Tx /Ff 8192 /T (PasswordField) /V (PASSWORD-SECRET)",
            widgetRectangle: "20 20 220 40");

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(pdf);

        Assert.Contains(report.Findings, finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.Contains("PASSWORD-SECRET", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("260 20 320 40", true)]
    [InlineData("-80 20 0 40", true)]
    [InlineData("-20 20 20 40", false)]
    public void ContentSafetyUsesPageIntersectionForWidgetVisibility(string widgetRectangle, bool shouldBeConcealed) {
        byte[] pdf = BuildSingleWidgetFieldPdf(
            fieldEntries: "/FT /Tx /T (PositionedField) /V (POSITIONED-VALUE)",
            widgetRectangle);

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(pdf);
        bool concealed = report.Findings.Any(finding =>
            finding.Kind == OfficeContentConcealmentKind.HiddenByProperty &&
            finding.TextPreview.Contains("POSITIONED-VALUE", StringComparison.Ordinal));

        Assert.Equal(shouldBeConcealed, concealed);
    }

    [Fact]
    public void ContentSafetySurfacesOptionalContentAnnotationPayloadsAndDefaultWidgetValues() {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [8 0 R] >> /OCProperties << /OCGs [6 0 R] /D << /BaseState /ON /OFF [6 0 R] >> >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R /Annots [7 0 R 9 0 R 11 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "6 0 obj\n<< /Type /OCG /Name (Hidden annotations) >>\nendobj",
            "7 0 obj\n<< /Type /Annot /Subtype /FreeText /Rect [20 20 80 50] /OC 6 0 R /RC (<body>RICH-OPTIONAL-SECRET</body>) /Contents (PLAIN-OPTIONAL-SECRET) /F 4 >>\nendobj",
            "8 0 obj\n<< /FT /Tx /T (DefaultHiddenField) /DV (DEFAULT-HIDDEN-FIELD) /Kids [9 0 R] >>\nendobj",
            "9 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 8 0 R /Rect [90 20 210 50] /P 3 0 R /OC 6 0 R /F 4 >>\nendobj",
            "10 0 obj\n<< /Type /OCMD /OCGs [6 0 R] /P /AllOn >>\nendobj",
            "11 0 obj\n<< /Type /Annot /Subtype /Text /Rect [20 70 40 90] /OC 10 0 R /Contents (MEMBERSHIP-OPTIONAL-SECRET) /F 4 >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 12 >>",
            "%%EOF"
        });

        OfficeContentSafetyReport report = PdfDocument.InspectContentSafety(Encoding.ASCII.GetBytes(pdf));

        Assert.Contains(report.Findings, finding => finding.TextPreview.Contains("RICH-OPTIONAL-SECRET", StringComparison.Ordinal));
        Assert.Contains(report.Findings, finding => finding.TextPreview.Contains("PLAIN-OPTIONAL-SECRET", StringComparison.Ordinal));
        Assert.Contains(report.Findings, finding => finding.TextPreview.Contains("MEMBERSHIP-OPTIONAL-SECRET", StringComparison.Ordinal));
        Assert.Contains(report.Findings, finding => finding.TextPreview.Contains("DEFAULT-HIDDEN-FIELD", StringComparison.Ordinal));
        Assert.All(
            report.Findings.Where(finding => finding.Kind == OfficeContentConcealmentKind.HiddenByProperty),
            finding => Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability));
    }

    private static byte[] BuildHiddenOptionalContentPdf(bool unsupportedViewIntent = false) {
        const string pageContent = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET\n/OC /Hidden BDC BT /F1 12 Tf 20 100 Td (HIDDEN-PAGE-CONTENT) Tj ET EMC\n/OuterForm Do";
        const string outerFormContent = "/HiddenForm Do";
        const string formContent = "BT /F1 12 Tf 20 60 Td (HIDDEN-FORM-CONTENT) Tj ET";
        string pageStream = StreamObject(4, string.Empty, pageContent);
        string outerFormStream = StreamObject(
            7,
            "/Type /XObject /Subtype /Form /BBox [0 0 240 180] /Resources << /XObject << /HiddenForm 8 0 R >> >>",
            outerFormContent);
        string hiddenFormStream = StreamObject(
            8,
            "/Type /XObject /Subtype /Form /BBox [0 0 240 180] /OC 6 0 R /Resources << /Font << /F1 5 0 R >> >>",
            formContent);
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [6 0 R] /D << /BaseState /ON /OFF [6 0 R] >> >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> /Properties << /Hidden 6 0 R >> /XObject << /OuterForm 7 0 R >> >> /Contents 4 0 R >>\nendobj",
            pageStream,
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "6 0 obj\n<< /Type /OCG /Name (Hidden layer)" + (unsupportedViewIntent ? " /Intent /Design" : string.Empty) + " >>\nendobj",
            outerFormStream,
            hiddenFormStream,
            "trailer\n<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSingleWidgetFieldPdf(string fieldEntries, string widgetRectangle) {
        const string content = "BT /F1 12 Tf 20 150 Td (VISIBLE-CONTENT) Tj ET";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R /Annots [7 0 R] >>\nendobj",
            StreamObject(4, string.Empty, content),
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "6 0 obj\n<< " + fieldEntries + " /Kids [7 0 R] >>\nendobj",
            "7 0 obj\n<< /Type /Annot /Subtype /Widget /Parent 6 0 R /Rect [" + widgetRectangle + "] /P 3 0 R /F 4 >>\nendobj",
            "trailer\n<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static string StreamObject(int objectNumber, string dictionaryEntries, string content) {
        int length = Encoding.ASCII.GetByteCount(content);
        string entries = string.IsNullOrWhiteSpace(dictionaryEntries) ? string.Empty : dictionaryEntries + " ";
        return objectNumber + " 0 obj\n<< " + entries + "/Length " + length + " >>\nstream\n" + content + "\nendstream\nendobj";
    }
}
