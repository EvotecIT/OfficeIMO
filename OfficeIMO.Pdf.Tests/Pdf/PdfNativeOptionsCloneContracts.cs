using OfficeIMO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfNativeOptionsCloneContracts {
    [Fact]
    public void FormSnapshotRetainsFontConfigurationDiagnosticsSinkAndCancellation() {
        byte[] fontBytes = { 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        using var cancellation = new CancellationTokenSource();
        var report = new PdfConversionReport();
        var original = new PdfFormFillerOptions {
            KeepNeedAppearances = true, MaxXfdfDocumentCharacters = 1234,
            MaxXfdfFields = 12, MaxXfdfValueCharacters = 321,
            CancellationToken = cancellation.Token,
            AppearanceFontFallbacks = new PdfEmbeddedFontFallbackSet(new[] {
                new PdfEmbeddedFontFallbackCandidate("Fallback", fontBytes)
            })
        }.UseAppearanceFont("Original", fontBytes).ReportDiagnosticsTo(report, "form-snapshot");
        PdfFormFillerOptions copy = original.Clone();
        Assert.True(copy.KeepNeedAppearances);
        Assert.Equal(1234, copy.MaxXfdfDocumentCharacters);
        Assert.Equal(12, copy.MaxXfdfFields);
        Assert.Equal(321, copy.MaxXfdfValueCharacters);
        Assert.Equal(cancellation.Token, copy.CancellationToken);
        Assert.Equal("Original", copy.AppearanceFontFamily!.FamilyName);
        Assert.Equal("Fallback", Assert.Single(copy.AppearanceFontFallbacks!.Candidates).FontName);
        byte[] returnedFont = copy.AppearanceFontFamily!.Regular;
        returnedFont[0] = 255;
        Assert.Equal(0, original.AppearanceFontFamily!.Regular[0]);
        copy.UseAppearanceFont("Override", fontBytes);
        copy.AppearanceFontFallbacks = null;
        Assert.Equal("Original", original.AppearanceFontFamily!.FamilyName);
        Assert.True(original.HasAppearanceFontFallbacks);
        copy.AddTextDiagnostics(new[] { new PdfTextEncodingDiagnostic("field", 0, "U+0100", "Ā", false) });
        Assert.Equal("form-snapshot", Assert.Single(report.Warnings).Converter);
        var operation = (PdfFormFillerOptions)((IPdfOperationCancellation)original).CloneForOperation();
        Assert.NotSame(original, operation);
        Assert.Equal("Original", operation.AppearanceFontFamily!.FamilyName);
        Assert.True(operation.HasAppearanceFontFallbacks);
        Assert.Equal(cancellation.Token, operation.CancellationToken);
    }

    [Fact]
    public void RewriteSnapshotSeparatesAllCollectionsAndBothAuthenticatedReadPolicies() {
        var policy = new PdfLoadOptions {
            Password = "rewrite-owner", IncludeImages = false,
            Limits = new PdfReadLimits { MaxInputBytes = 1_000_000, MaxPages = 5 }
        };
        var original = new PdfRewritePreservationOptions {
            OriginalReadOptions = policy, RewrittenReadOptions = policy,
            FilterActionsByPreservedTypes = true, PreserveLinkAnnotations = false
        }.RequireTextMarkers("CloneMarker").AllowMetadataChanges("Title");
        original.PreservedActionTypes.Add("URI");
        original.ExcludedAnnotationSubtypes.Add("Widget");
        original.ExcludedLinkAnnotationUris.Add("https://example.test/link");
        original.ExcludedActionUris.Add("https://example.test/action");
        PdfRewritePreservationOptions copy = original.Clone();
        Assert.True(copy.FilterActionsByPreservedTypes);
        Assert.False(copy.PreserveLinkAnnotations);
        Assert.True(copy.AllowedMetadataChanges.Contains("title"));
        Assert.False(copy.PreservedActionTypes.Contains("uri"));
        ICollection<string>[] originals = { original.RequiredTextMarkers, original.AllowedMetadataChanges,
            original.PreservedActionTypes, original.ExcludedAnnotationSubtypes, original.ExcludedLinkAnnotationUris, original.ExcludedActionUris };
        ICollection<string>[] copies = { copy.RequiredTextMarkers, copy.AllowedMetadataChanges,
            copy.PreservedActionTypes, copy.ExcludedAnnotationSubtypes, copy.ExcludedLinkAnnotationUris, copy.ExcludedActionUris };
        for (int i = 0; i < originals.Length; i++) {
            string[] expected = originals[i].ToArray();
            copies[i].Add("copy-only");
            Assert.Equal(expected, originals[i]);
        }
        Assert.NotSame(policy, copy.OriginalReadOptions);
        Assert.NotSame(policy, copy.RewrittenReadOptions);
        Assert.NotSame(copy.OriginalReadOptions, copy.RewrittenReadOptions);
        Assert.NotSame(policy.Limits, copy.OriginalReadOptions!.Limits);
        Assert.NotSame(copy.OriginalReadOptions.Limits, copy.RewrittenReadOptions!.Limits);
        Assert.Equal("rewrite-owner", copy.OriginalReadOptions.Password);
        Assert.Equal(5, copy.RewrittenReadOptions.Limits.MaxPages);
        Assert.False(copy.OriginalReadOptions.IncludeImages);
        Assert.Equal(PdfPermissionPolicy.Enforce, copy.OriginalReadOptions.PermissionPolicy);

        byte[] source = PdfDocument.Create().Paragraph(p => p.AddText("CloneMarker")).ToBytes();
        byte[] encrypted = PdfSecurityEditor.Encrypt(source, new PdfStandardEncryptionOptions("rewrite-open") {
            OwnerPassword = "rewrite-owner"
        }).Pdf;
        Assert.True(PdfRewritePreservation.Assess(encrypted, encrypted, original).IsPreserved);
        Assert.False(PdfRewritePreservation.Assess(encrypted, encrypted, copy).IsPreserved);
    }

    [Fact]
    public void LayoutSnapshotKeepsBoundedSettingsAndColorsIndependent() {
        var original = new PdfLayoutDebugOverlayOptions {
            ShowWords = false, ShowLines = false, ShowRegions = false, ShowReadingOrder = false,
            MaxElements = 12, MaxRasterPixels = 3456,
            WordColor = OfficeColor.FromRgba(10, 20, 30, 40),
            LineColor = OfficeColor.Blue, RegionColor = OfficeColor.Green
        };
        PdfLayoutDebugOverlayOptions copy = original.Clone();
        Assert.NotSame(original, copy);
        Assert.False(copy.ShowWords || copy.ShowLines || copy.ShowRegions || copy.ShowReadingOrder);
        Assert.Equal(12, copy.MaxElements);
        Assert.Equal(3456, copy.MaxRasterPixels);
        Assert.Equal(original.WordColor, copy.WordColor);
        Assert.Equal(original.LineColor, copy.LineColor);
        Assert.Equal(original.RegionColor, copy.RegionColor);
        copy.MaxElements = 3;
        copy.WordColor = OfficeColor.Red;
        Assert.Equal(12, original.MaxElements);
        Assert.Equal(OfficeColor.FromRgba(10, 20, 30, 40), original.WordColor);
    }
}
