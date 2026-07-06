using System;
using System.Collections.Generic;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

internal static class PdfRewritePreservationMatrixScenarioSupport {
    public const int PremiumScenarioCount = 8;

    public static IReadOnlyList<PdfRewritePreservationMatrixScenario> BuildPremiumFeatureScenarios() {
        byte[] preservationSource = PdfRewritePreservationTestSupport.BuildPreservationProofPdf();
        byte[] sourceStructure = PdfRewritePreservationTestSupport.BuildSourceStructurePreservationProofPdf();
        byte[] optionalContent = PdfOptionalContentSupport.BuildOptionalContentMetadataPdf();
        byte[] formSource = PdfFormAppearanceProofTestSupport.BuildFormAppearanceProofPdf();
        byte[] taggedSource = PdfRewritePreservationTestSupport.BuildTaggedPreservationProofPdf();
        byte[] activeContentSource = BuildPageActiveContentProofPdf();
        byte[] signedSource = PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf();

        return new[] {
            new PdfRewritePreservationMatrixScenario(
                    "metadata-update-safe",
                    "MetadataUpdate",
                    preservationSource,
                    pdf => PdfMetadataEditor.UpdateMetadata(pdf, title: "Updated preservation title")) {
                    PreservationOptions = new PdfRewritePreservationOptions()
                        .AllowMetadataChanges("Title")
                        .RequireTextMarkers("PreservationMarker", "SecondPageMarker")
                }
                .WithSourceFeatures("metadata", "xmp", "attachments", "output-intents"),
            new PdfRewritePreservationMatrixScenario(
                    "source-structure-drift-detected",
                    "ByteMutation",
                    sourceStructure,
                    pdf => ReplaceFirstAscii(pdf, "/Type /XRef", "/Type /Xbad")) {
                    ExpectedClassification = PdfRewritePreservationMatrixClassification.PreservationFailed
                }
                .WithSourceFeatures("source-structure", "xref-stream", "object-stream", "incremental"),
            new PdfRewritePreservationMatrixScenario(
                    "optional-content-drift-detected",
                    "ByteMutation",
                    optionalContent,
                    pdf => ReplaceFirstAscii(pdf, "Print layer", "Proof layer")) {
                    ExpectedClassification = PdfRewritePreservationMatrixClassification.PreservationFailed
                }
                .WithSourceFeatures("optional-content", "layers", "ocg-order", "visibility-state"),
            new PdfRewritePreservationMatrixScenario(
                    "form-fill-safe",
                    "FormFill",
                    formSource,
                    pdf => PdfFormFiller.FillFields(pdf, new Dictionary<string, string> {
                        ["Name"] = "Visible Value",
                        ["Country"] = "PL"
                    }))
                .WithSourceFeatures("forms", "text-field", "choice-field", "appearances"),
            new PdfRewritePreservationMatrixScenario(
                    "forms-full-rewrite-blocked",
                    "MetadataUpdate",
                    formSource,
                    pdf => PdfMetadataEditor.UpdateMetadata(pdf, title: "Blocked form rewrite")) {
                    ExpectedClassification = PdfRewritePreservationMatrixClassification.Blocked
                }
                .WithSourceFeatures("forms", "acroform", "widget-annotations"),
            new PdfRewritePreservationMatrixScenario(
                    "tagged-full-rewrite-blocked",
                    "MetadataUpdate",
                    taggedSource,
                    pdf => PdfMetadataEditor.UpdateMetadata(pdf, title: "Blocked tagged rewrite")) {
                    ExpectedClassification = PdfRewritePreservationMatrixClassification.Blocked
                }
                .WithSourceFeatures("tagged-content", "structure-tree", "accessibility"),
            new PdfRewritePreservationMatrixScenario(
                    "active-content-full-rewrite-blocked",
                    "MetadataUpdate",
                    activeContentSource,
                    pdf => PdfMetadataEditor.UpdateMetadata(pdf, title: "Blocked active rewrite")) {
                    ExpectedClassification = PdfRewritePreservationMatrixClassification.Blocked
                }
                .WithSourceFeatures("active-content", "page-actions", "javascript"),
            new PdfRewritePreservationMatrixScenario(
                    "signed-full-rewrite-blocked",
                    "MetadataUpdate",
                    signedSource,
                    pdf => PdfMetadataEditor.UpdateMetadata(pdf, title: "Blocked")) {
                    ExpectedClassification = PdfRewritePreservationMatrixClassification.Blocked
                }
                .WithSourceFeatures("signature", "doc-mdp", "dss", "incremental")
        };
    }

    private static byte[] BuildPageActiveContentProofPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /AA << /O 5 0 R >> >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /S /JavaScript /JS (app.alert('Page open')) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] ReplaceFirstAscii(byte[] source, string oldValue, string newValue) {
        Assert.Equal(oldValue.Length, newValue.Length);

        byte[] oldBytes = Encoding.ASCII.GetBytes(oldValue);
        byte[] newBytes = Encoding.ASCII.GetBytes(newValue);
        int index = IndexOf(source, oldBytes);
        Assert.True(index >= 0, "Expected PDF test fixture to contain marker: " + oldValue);

        byte[] rewritten = (byte[])source.Clone();
        Array.Copy(newBytes, 0, rewritten, index, newBytes.Length);
        return rewritten;
    }

    private static int IndexOf(byte[] source, byte[] value) {
        for (int i = 0; i <= source.Length - value.Length; i++) {
            bool match = true;
            for (int j = 0; j < value.Length; j++) {
                if (source[i + j] != value[j]) {
                    match = false;
                    break;
                }
            }

            if (match) {
                return i;
            }
        }

        return -1;
    }
}
