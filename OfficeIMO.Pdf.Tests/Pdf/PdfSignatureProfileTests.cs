using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSignatureProfileTests {
    [Fact]
    public void CertificationProfileEmitsDocMdpCatalogAndTransformPermissions() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Certification source"))
            .ToBytes();

        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                Profile = PdfSignatureProfile.Certification,
                CertificationPermission = PdfCertificationPermissionLevel.FormFillingAndSignatures,
                FieldName = "CertificationSignature",
                ReservedSignatureContentsBytes = 512
            });
        PdfDocumentSecurityInfo security = PdfInspector.Inspect(preparation.PreparedPdf).Security;
        string raw = PdfEncoding.Latin1GetString(preparation.PreparedPdf);

        Assert.Equal(PdfSignatureProfile.Certification, preparation.Profile);
        Assert.True(security.HasDocMDPPermissions);
        Assert.Equal(2, security.DocMDPPermissionLevel);
        Assert.Contains("/Perms << /DocMDP", raw, StringComparison.Ordinal);
        Assert.Contains("/TransformMethod /DocMDP", raw, StringComparison.Ordinal);
        Assert.Contains("/P 2 /V /1.2", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void VisibleApprovalProfileCreatesWidgetAndAppearanceOnSelectedPage() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Visible approval source"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Signature target page"))
            .ToBytes();
        var appearance = new PdfVisibleSignatureAppearanceOptions {
            PageNumber = 2,
            X = 42,
            Y = 54,
            Width = 210,
            Height = 60,
            Text = "Approved by external signer"
        };

        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                Profile = PdfSignatureProfile.Approval,
                FieldName = "VisibleApproval",
                VisibleAppearance = appearance,
                ReservedSignatureContentsBytes = 512
            });
        PdfReadDocument document = PdfReadDocument.Open(preparation.PreparedPdf);
        PdfFormField field = Assert.Single(document.FormFields, formField => formField.Name == "VisibleApproval");
        PdfFormWidget widget = Assert.Single(field.Widgets);
        string raw = PdfEncoding.Latin1GetString(preparation.PreparedPdf);

        Assert.Equal(2, widget.PageNumber);
        Assert.True(widget.IsPrint);
        Assert.Equal(42, widget.X1);
        Assert.Equal(54, widget.Y1);
        Assert.Equal(252, widget.X2);
        Assert.Equal(114, widget.Y2);
        Assert.Contains("Approved by external signer", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Widget", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Form", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentTimestampProfileSelectsRfc3161SubFilter() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Timestamp source"))
            .ToBytes();

        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                Profile = PdfSignatureProfile.DocumentTimestamp,
                FieldName = "DocumentTimestamp",
                ReservedSignatureContentsBytes = 512
            });
        string raw = PdfEncoding.Latin1GetString(preparation.PreparedPdf);

        Assert.Equal(PdfSignatureProfile.DocumentTimestamp, preparation.Profile);
        Assert.Equal("ETSI.RFC3161", preparation.SubFilter);
        Assert.Contains("/Type /DocTimeStamp", raw, StringComparison.Ordinal);
        Assert.Contains("/SubFilter /ETSI.RFC3161", raw, StringComparison.Ordinal);
    }
}
