using System;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRewritePreservationTests {
    [Fact]
    public void AssertPreserved_AllowsDeclaredMetadataUpdateAndPreservesDocumentSignals() {
        byte[] source = PdfRewritePreservationTestSupport.BuildPreservationProofPdf();
        byte[] updated = PdfMetadataEditor.UpdateMetadata(source, title: "Updated preservation title");

        PdfRewritePreservationReport report = PdfRewritePreservation.AssertPreserved(
            source,
            updated,
            new PdfRewritePreservationOptions()
                .AllowMetadataChanges("Title")
                .RequireTextMarkers("PreservationMarker", "SecondPageMarker"));

        Assert.True(report.IsPreserved);
        Assert.Empty(report.Issues);
        Assert.Equal(2, report.Original.PageCount);
        Assert.Equal(2, report.Rewritten.PageCount);
        Assert.Equal("Updated preservation title", report.Rewritten.Metadata.Title);
        Assert.Equal(report.Original.LinkAnnotationCount, report.Rewritten.LinkAnnotationCount);
        Assert.Equal(report.Original.NamedDestinations.Count, report.Rewritten.NamedDestinations.Count);
        Assert.Equal(report.Original.Attachments.Count, report.Rewritten.Attachments.Count);
        Assert.True(report.Rewritten.HasXmpMetadata);
        Assert.True(report.Rewritten.HasOutputIntents);
        Assert.Equal("UseThumbs", report.Rewritten.CatalogPageMode);
        Assert.Equal("SinglePage", report.Rewritten.CatalogPageLayout);
        Assert.Equal("en-US", report.Rewritten.CatalogLanguage);
    }

    [Fact]
    public void Assess_ReportsLostMarkersAndFeaturesAfterUnexpectedPageDeletion() {
        byte[] source = PdfRewritePreservationTestSupport.BuildPreservationProofPdf();
        byte[] deleted = PdfPageEditor.DeletePages(source, 2);

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(
            source,
            deleted,
            new PdfRewritePreservationOptions().RequireTextMarkers("SecondPageMarker"));

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue => issue.Feature == "PageCount");
        Assert.Contains(report.Issues, issue => issue.Feature == "TextMarker" && issue.Expected == "SecondPageMarker");
        Assert.Contains(report.Issues, issue => issue.Feature == "NamedDestinations");
        Assert.Contains(report.Issues, issue => issue.Feature == "LinkAnnotations");
        Assert.Contains("PDF rewrite preservation failed", report.Summary, StringComparison.Ordinal);

        var exception = Assert.Throws<InvalidOperationException>(() => report.ThrowIfFailed());
        Assert.Contains("PageCount", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_ReportsReadableXmpMetadataContentDrift() {
        byte[] source = PdfRewritePreservationTestSupport.BuildPreservationProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "OfficeIMO.Pdf", "OfficeIMO.Bad");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "XmpMetadata.Producer" &&
            issue.Expected == "OfficeIMO.Pdf" &&
            issue.Actual == "OfficeIMO.Bad");
        Assert.Contains("XmpMetadata.Producer", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_ReportsAttachmentMetadataContentDrift() {
        byte[] source = PdfRewritePreservationTestSupport.BuildPreservationProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "text#2Fplain", "text#2Fwrong");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "EmbeddedFiles[0].MimeType" &&
            issue.Expected == "text/plain" &&
            issue.Actual == "text/wrong");
        Assert.Contains("EmbeddedFiles[0].MimeType", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_ReportsOutputIntentMetadataContentDrift() {
        byte[] source = PdfRewritePreservationTestSupport.BuildPreservationProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "73524742", "73426164");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "OutputIntents[0].OutputConditionIdentifier" &&
            issue.Expected == "sRGB IEC61966-2.1" &&
            issue.Actual == "sBad IEC61966-2.1");
        Assert.Contains("OutputIntents[0].OutputConditionIdentifier", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_PreservesNavigationMetadataForUnchangedPdf() {
        byte[] source = PdfRewritePreservationTestSupport.BuildNavigationPreservationProofPdf();

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, source);

        Assert.True(report.IsPreserved);
        Assert.Empty(report.Issues);
        Assert.Equal(2, report.Original.NamedDestinations.Count);
        Assert.Equal(new[] { "Chapter1", "Chapter2" }, report.Rewritten.NamedDestinationNames);
        Assert.Equal(2, report.Original.PageLabels.Count);
        Assert.Equal(new[] { "r", "D" }, report.Rewritten.PageLabels.Select(label => label.Style).ToArray());
        Assert.Equal("front-", report.Rewritten.PageLabels[0].Prefix);
        Assert.Equal(2, report.Rewritten.PageLabels[1].StartNumber);
    }

    [Fact]
    public void Assess_ReportsNamedDestinationContentDrift() {
        byte[] source = PdfRewritePreservationTestSupport.BuildNavigationPreservationProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "Chapter1", "Section1");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "NamedDestinations[0].Name" &&
            issue.Expected == "Chapter1" &&
            issue.Actual == "Section1");
        Assert.Contains("NamedDestinations[0].Name", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_ReportsPageLabelContentDrift() {
        byte[] source = PdfRewritePreservationTestSupport.BuildNavigationPreservationProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "front-", "draft-");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "PageLabels[0].Prefix" &&
            issue.Expected == "front-" &&
            issue.Actual == "draft-");
        Assert.Contains("PageLabels[0].Prefix", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_PreservesViewerAndActionStateForUnchangedPdf() {
        byte[] source = PdfRewritePreservationTestSupport.BuildViewerActionPreservationProofPdf();

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, source);

        Assert.True(report.IsPreserved);
        Assert.Empty(report.Issues);
        Assert.NotNull(report.Original.OpenAction);
        Assert.Equal("Destination", report.Rewritten.OpenAction!.ActionType);
        Assert.Equal(1, report.Rewritten.OpenAction.PageNumber);
        Assert.Equal(PdfOpenActionDestinationMode.Xyz, report.Rewritten.OpenAction.DestinationMode);
        Assert.Equal(180, report.Rewritten.OpenAction.DestinationTop);
        Assert.NotNull(report.Rewritten.ViewerPreferences);
        Assert.True(report.Rewritten.ViewerPreferences!.GetBoolean("HideToolbar"));
        Assert.True(report.Rewritten.ViewerPreferences.GetBoolean("DisplayDocTitle"));
        Assert.Equal(2, report.Rewritten.CatalogActions.Count);
        Assert.Contains(report.Rewritten.CatalogActions, action => action.Name == "Open" && action.Source == "Names/JavaScript");
        Assert.Contains(report.Rewritten.CatalogActions, action => action.Name == "AA.WC" && action.TriggerName == "WC");
        Assert.Equal(2, report.Rewritten.PageActionCount);
        Assert.Contains(report.Rewritten.PageActions, action => action.TriggerName == "O" && action.ActionType == "JavaScript");
        Assert.Contains(report.Rewritten.PageActions, action => action.TriggerName == "C" && action.ActionType == "Launch");
    }

    [Fact]
    public void Assess_ReportsOpenActionContentDrift() {
        byte[] source = PdfRewritePreservationTestSupport.BuildViewerActionPreservationProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "/XYZ", "/Fit");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "OpenAction.DestinationMode" &&
            issue.Expected == "Xyz" &&
            issue.Actual == "Fit");
        Assert.Contains("OpenAction.DestinationMode", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_ReportsViewerPreferenceContentDrift() {
        byte[] source = PdfRewritePreservationTestSupport.BuildViewerActionPreservationProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "HideToolbar true", "HideToolbar null");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "ViewerPreferences.Values" &&
            issue.Expected.Contains("HideToolbar=true", StringComparison.Ordinal) &&
            issue.Actual.Contains("HideToolbar=null", StringComparison.Ordinal));
        Assert.Contains("ViewerPreferences.Values", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_ReportsCatalogActionContentDrift() {
        byte[] source = PdfRewritePreservationTestSupport.BuildViewerActionPreservationProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "(Open)", "(Load)");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature.EndsWith(".Name", StringComparison.Ordinal) &&
            issue.Expected == "Open" &&
            issue.Actual == "Load");
        Assert.Contains("CatalogActions", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_ReportsPageActionContentDrift() {
        byte[] source = PdfRewritePreservationTestSupport.BuildViewerActionPreservationProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "/C 10 0 R", "/D 10 0 R");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature.EndsWith(".TriggerName", StringComparison.Ordinal) &&
            issue.Expected == "C" &&
            issue.Actual == "D");
        Assert.Contains("PageActions", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_PreservesSourceStructureForUnchangedPdf() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSourceStructurePreservationProofPdf();

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, source);

        Assert.True(report.IsPreserved);
        Assert.Empty(report.Issues);
        Assert.Equal("1.4", report.Original.HeaderVersion);
        Assert.Equal("1.7", report.Original.CatalogVersion);
        Assert.Equal("1.7", report.Original.EffectiveVersion);
        Assert.True(report.Original.Security.HasXrefStreams);
        Assert.True(report.Original.Security.HasObjectStreams);
        Assert.True(report.Original.Security.HasPreviousRevision);
        Assert.True(report.Original.Security.HasIncrementalUpdates);
        Assert.True(report.Original.Security.StartXrefCount > 1);
    }

    [Fact]
    public void Assess_ReportsDocumentHeaderVersionDrift() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSourceStructurePreservationProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "%PDF-1.4", "%PDF-1.3");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "SourceStructure.HeaderVersion" &&
            issue.Expected == "1.4" &&
            issue.Actual == "1.3");
        Assert.Contains("SourceStructure.HeaderVersion", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_ReportsCatalogVersionDrift() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSourceStructurePreservationProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "/Version /1.7", "/Version /1.6");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "SourceStructure.CatalogVersion" &&
            issue.Expected == "1.7" &&
            issue.Actual == "1.6");
        Assert.Contains("SourceStructure.CatalogVersion", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_ReportsXrefAndObjectStreamMarkerLoss() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSourceStructurePreservationProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "/Type /XRef", "/Type /Xbad");
        rewritten = ReplaceFirstAscii(rewritten, "/Type /ObjStm", "/Type /ObjBad");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue => issue.Feature == "SourceStructure.XrefStreams");
        Assert.Contains(report.Issues, issue => issue.Feature == "SourceStructure.ObjectStreams");
        Assert.Contains("SourceStructure.XrefStreams", report.Summary, StringComparison.Ordinal);
        Assert.Contains("SourceStructure.ObjectStreams", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_ReportsPreviousRevisionMarkerLoss() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSourceStructurePreservationProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "/Prev 100", "/Pbad 100");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue => issue.Feature == "SourceStructure.PreviousRevision");
        Assert.Contains("SourceStructure.PreviousRevision", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_ReportsOptionalContentMetadataContentDrift() {
        byte[] source = PdfOptionalContentSupport.BuildOptionalContentMetadataPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "Print layer", "Proof layer");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "OptionalContent.Groups[0].Name" &&
            issue.Expected == "Print layer" &&
            issue.Actual == "Proof layer");
        Assert.Contains("OptionalContent.Groups[0].Name", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_PreservesTaggedContentForUnchangedTaggedPdf() {
        byte[] source = PdfRewritePreservationTestSupport.BuildTaggedPreservationProofPdf();

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, source);

        Assert.True(report.IsPreserved);
        Assert.Empty(report.Issues);
        Assert.True(report.Original.HasTaggedContent);
        Assert.NotNull(report.Original.TaggedContent);
        Assert.Equal(report.Original.TaggedContent!.StructureElementCount, report.Rewritten.TaggedContent!.StructureElementCount);
        Assert.Contains("Document", report.Rewritten.TaggedContent.StructureTypes);
        Assert.Contains("H1", report.Rewritten.TaggedContent.StructureTypes);
        Assert.Contains("P", report.Rewritten.TaggedContent.StructureTypes);
        Assert.True(report.Rewritten.TaggedContent.MarkedContentReferenceCount > 0);
    }

    [Fact]
    public void Assess_ReportsTaggedStructureTypeDrift() {
        byte[] source = PdfRewritePreservationTestSupport.BuildTaggedPreservationProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "/S /H1", "/S /H2");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "TaggedContent.StructureTypes" &&
            issue.Expected == "Document,H1,P" &&
            issue.Actual == "Document,H2,P");
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "TaggedContent.StructureTypeCounts" &&
            issue.Expected.Contains("H1=1", StringComparison.Ordinal) &&
            issue.Actual.Contains("H2=1", StringComparison.Ordinal));
        Assert.Contains("TaggedContent.StructureTypes", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_PreservesSignatureSecurityStateForUnchangedSignedPdf() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf();

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, source);

        Assert.True(report.IsPreserved);
        Assert.Empty(report.Issues);
        Assert.True(report.Original.Security.HasSignatures);
        Assert.True(report.Original.Security.HasDocMDPPermissions);
        Assert.True(report.Original.Security.HasLongTermValidationEvidence);
        Assert.True(report.Original.Security.RequiresAppendOnlyMutation);
    }

    [Fact]
    public void Assess_ReportsSignatureSignerContentDrift() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "(Alice)", "(Carol)");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "Security.Signatures[0].SignerName" &&
            issue.Expected == "Alice" &&
            issue.Actual == "Carol");
        Assert.Contains("Security.Signatures[0].SignerName", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_ReportsDocMdpPermissionContentDrift() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "/V /1.2 /P 2", "/V /1.2 /P 3");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "Security.DocMDPPermissionLevel" &&
            issue.Expected == "2" &&
            issue.Actual == "3");
        Assert.Contains("Security.DocMDPPermissionLevel", report.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Assess_ReportsDssValidationEvidenceContentDrift() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf();
        byte[] rewritten = ReplaceFirstAscii(source, "ABCDEF", "ABC123");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, rewritten);

        Assert.False(report.IsPreserved);
        Assert.Contains(report.Issues, issue =>
            issue.Feature == "Security.DSS.VriKeys" &&
            issue.Expected == "ABCDEF" &&
            issue.Actual == "ABC123");
        Assert.Contains("Security.DSS.VriKeys", report.Summary, StringComparison.Ordinal);
    }

    private static byte[] ReplaceFirstAscii(byte[] source, string oldValue, string newValue) {
        Assert.Equal(oldValue.Length, newValue.Length);

        byte[] oldBytes = System.Text.Encoding.ASCII.GetBytes(oldValue);
        byte[] newBytes = System.Text.Encoding.ASCII.GetBytes(newValue);
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
