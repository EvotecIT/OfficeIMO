using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    [Fact]
    public void Preflight_AllowsSimpleViewerPreferencePdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildViewerPreferencePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasViewerPreferences);
        Assert.False(report.Probe.HasOpenActions);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasViewerPreferences);
        Assert.False(report.DocumentInfo.HasOpenActions);
        AssertViewerPreferences(report.DocumentInfo);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsComplexViewerPreferencePdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexViewerPreferencePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasViewerPreferences);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasViewerPreferences);
        Assert.False(report.DocumentInfo.HasReadableViewerPreferences);
        Assert.Null(report.DocumentInfo.ViewerPreferences);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.ViewerPreferences, "PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsTaggedPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildTaggedPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasTaggedContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasTaggedContent);
        Assert.True(report.DocumentInfo.HasReadableTaggedContent);
        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(report.DocumentInfo.TaggedContent);
        Assert.True(tagged.Marked);
        Assert.Equal(5, tagged.StructTreeRootObjectNumber);
        Assert.Equal(7, tagged.ParentTreeObjectNumber);
        Assert.Null(tagged.ParentTreeNextKey);
        Assert.Empty(tagged.RoleMap);
        Assert.Equal(new[] { 6 }, tagged.RootElementObjectNumbers);
        Assert.Equal(new[] { 0 }, tagged.ParentTreeStructParentIndexes);
        Assert.Equal(1, tagged.ParentTreeEntryCount);
        PdfStructureElementInfo element = Assert.Single(tagged.StructureElements);
        Assert.Equal(6, element.ObjectNumber);
        Assert.Equal("Document", element.StructureType);
        Assert.Equal(5, element.ParentObjectNumber);
        Assert.Equal(0, element.MarkedContentReferenceCount);
        Assert.Equal(0, element.ObjectReferenceCount);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF tagged content structure is not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.TaggedContent, "PDF tagged content structure is not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Inspect_ReadsParentTreeIndexesFromNumberTreeKids() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildTaggedPdfWithParentTreeKids());

        Assert.True(info.HasTaggedContent);
        Assert.True(info.HasReadableTaggedContent);
        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(info.TaggedContent);
        Assert.Equal(7, tagged.ParentTreeObjectNumber);
        Assert.Equal(new[] { 0 }, tagged.ParentTreeStructParentIndexes);
        Assert.Equal(1, tagged.ParentTreeEntryCount);
    }

    [Fact]
    public void Inspect_ReadsGeneratedTaggedContentMetadata() {
        byte[] pdf = PdfDocument.Create()
            .TaggedPdfCatalogMarkers()
            .Language("en-US")
            .H1("Tagged heading")
            .Paragraph(p => p.Text("Tagged paragraph."))
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);

        Assert.True(info.HasTaggedContent);
        Assert.True(info.HasReadableTaggedContent);
        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(info.TaggedContent);
        Assert.True(tagged.Marked);
        Assert.False(tagged.Suspects.HasValue);
        Assert.NotNull(tagged.StructTreeRootObjectNumber);
        Assert.NotNull(tagged.ParentTreeObjectNumber);
        Assert.True(tagged.ParentTreeNextKey > 0);
        Assert.Empty(tagged.RoleMap);
        Assert.NotEmpty(tagged.RootElementObjectNumbers);
        Assert.NotEmpty(tagged.ParentTreeStructParentIndexes);
        Assert.True(tagged.ParentTreeEntryCount > 0);
        Assert.True(tagged.StructureElementCount >= 3);
        Assert.Contains("Document", tagged.StructureTypes);
        Assert.Contains("H1", tagged.StructureTypes);
        Assert.Contains("P", tagged.StructureTypes);
        PdfStructureElementInfo document = Assert.Single(tagged.StructureElements, element => element.StructureType == "Document");
        Assert.Equal("en-US", document.Language);
        Assert.True(document.HasChildElements);
        Assert.Contains(tagged.StructureElements, element => element.StructureType == "H1" && element.MarkedContentReferenceCount > 0);
        Assert.Contains(tagged.StructureElements, element => element.StructureType == "P" && element.MarkedContentReferenceCount > 0);
    }

    [Fact]
    public void Preflight_AllowsSimpleXmpMetadataPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildXmpMetadataPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasXmpMetadata);
        Assert.False(report.Probe.HasCatalogUri);
        Assert.False(report.Probe.HasOutputIntents);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasXmpMetadata);
        Assert.True(report.DocumentInfo.HasReadableXmpMetadata);
        Assert.NotNull(report.DocumentInfo.XmpMetadata);
        Assert.Equal(5, report.DocumentInfo.XmpMetadata!.ObjectNumber);
        Assert.Equal("XML", report.DocumentInfo.XmpMetadata.Subtype);
        Assert.Equal(12, report.DocumentInfo.XmpMetadata.StreamSizeBytes);
        Assert.Equal(12, report.DocumentInfo.XmpMetadata.DecodedSizeBytes);
        Assert.False(report.DocumentInfo.XmpMetadata.IsWellFormedXml);
        Assert.Contains("<x:xmpmeta/>", report.DocumentInfo.XmpMetadata.RawXml, StringComparison.Ordinal);
        Assert.Empty(report.DocumentInfo.XmpMetadata.UnsupportedFilters);
        Assert.False(report.DocumentInfo.HasCatalogUri);
        Assert.False(report.DocumentInfo.HasOutputIntents);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsComplexXmpMetadataPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexXmpMetadataPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasXmpMetadata);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasXmpMetadata);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.XmpMetadata, "PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Inspect_ReadsGeneratedXmpMetadataFields() {
        byte[] pdf = PdfDocument.Create(new PdfOptions()
                .SetPdfAIdentification(3, "B")
                .SetPdfUaIdentification()
                .SetElectronicInvoiceMetadata("BASIC"))
            .Meta(title: "XMP readback", author: "OfficeIMO", subject: "Metadata stream", keywords: "alpha, beta;gamma")
            .Paragraph(p => p.Text("Generated XMP readback."))
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);

        Assert.True(info.HasXmpMetadata);
        Assert.True(info.HasReadableXmpMetadata);
        PdfXmpMetadataInfo xmp = Assert.IsType<PdfXmpMetadataInfo>(info.XmpMetadata);
        Assert.NotNull(xmp.ObjectNumber);
        Assert.Equal("XML", xmp.Subtype);
        Assert.Null(xmp.Filter);
        Assert.Empty(xmp.UnsupportedFilters);
        Assert.True(xmp.StreamSizeBytes > 100);
        Assert.Equal(xmp.StreamSizeBytes, xmp.DecodedSizeBytes);
        Assert.True(xmp.IsWellFormedXml);
        Assert.Contains("x:xmpmeta", xmp.RawXml, StringComparison.Ordinal);
        Assert.Equal("XMP readback", xmp.Title);
        Assert.Equal("OfficeIMO", xmp.Creator);
        Assert.Equal("Metadata stream", xmp.Description);
        Assert.Equal(new[] { "alpha", "beta", "gamma" }, xmp.Subjects);
        Assert.Equal("OfficeIMO.Pdf", xmp.Producer);
        Assert.Equal("alpha, beta;gamma", xmp.Keywords);
        Assert.True(xmp.HasPdfAIdentification);
        Assert.Equal(3, xmp.PdfAPart);
        Assert.Equal("B", xmp.PdfAConformance);
        Assert.True(xmp.HasPdfUaIdentification);
        Assert.Equal(1, xmp.PdfUaPart);
        Assert.True(xmp.HasElectronicInvoiceMetadata);
        Assert.Equal("INVOICE", xmp.ElectronicInvoiceDocumentType);
        Assert.Equal("factur-x.xml", xmp.ElectronicInvoiceDocumentFileName);
        Assert.Equal("1.0", xmp.ElectronicInvoiceVersion);
        Assert.Equal("BASIC", xmp.ElectronicInvoiceConformanceLevel);
    }

    [Fact]
    public void Inspect_ReadsXmpIdentificationFieldsByNamespaceUri() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildXmpMetadataPdfWithAlternateIdentificationPrefixes());

        Assert.True(info.HasXmpMetadata);
        Assert.True(info.HasReadableXmpMetadata);
        PdfXmpMetadataInfo xmp = Assert.IsType<PdfXmpMetadataInfo>(info.XmpMetadata);
        Assert.True(xmp.IsWellFormedXml);
        Assert.True(xmp.HasPdfAIdentification);
        Assert.Equal(3, xmp.PdfAPart);
        Assert.Equal("B", xmp.PdfAConformance);
        Assert.True(xmp.HasPdfUaIdentification);
        Assert.Equal(1, xmp.PdfUaPart);
        Assert.True(xmp.HasElectronicInvoiceMetadata);
        Assert.Equal("INVOICE", xmp.ElectronicInvoiceDocumentType);
        Assert.Equal("factur-x.xml", xmp.ElectronicInvoiceDocumentFileName);
        Assert.Equal("1.0", xmp.ElectronicInvoiceVersion);
        Assert.Equal("BASIC", xmp.ElectronicInvoiceConformanceLevel);
    }

    [Fact]
    public void Preflight_AllowsSimpleCatalogUriPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildCatalogUriPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasCatalogUri);
        Assert.False(report.Probe.HasOutputIntents);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasCatalogUri);
        Assert.False(report.DocumentInfo.HasOutputIntents);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF catalog URI dictionaries are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_DoesNotTreatLinkAnnotationUriAsCatalogUri() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildAnnotatedPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasAnnotations);
        Assert.False(report.Probe.HasCatalogUri);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasAnnotations);
        Assert.False(report.DocumentInfo.HasCatalogUri);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
    }

    [Fact]
    public void Preflight_AllowsComplexCatalogUriPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexCatalogUriPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasCatalogUri);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasCatalogUri);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF catalog URI dictionaries are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.CatalogUri, "PDF catalog URI dictionaries are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsSimpleOutputIntentPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildOutputIntentPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasOutputIntents);
        Assert.False(report.Probe.HasXmpMetadata);
        Assert.False(report.Probe.HasCatalogUri);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOutputIntents);
        Assert.True(report.DocumentInfo.HasReadableOutputIntents);
        Assert.Equal(1, report.DocumentInfo.OutputIntentCount);
        Assert.Equal(new[] { "GTS_PDFA1" }, report.DocumentInfo.OutputIntentSubtypes);
        Assert.Equal(new[] { "sRGB" }, report.DocumentInfo.OutputConditionIdentifiers);
        PdfOutputIntentInfo outputIntent = Assert.Single(report.DocumentInfo.OutputIntents);
        Assert.Equal(5, outputIntent.ObjectNumber);
        Assert.Equal("GTS_PDFA1", outputIntent.Subtype);
        Assert.Equal("sRGB", outputIntent.OutputConditionIdentifier);
        Assert.False(outputIntent.HasDestinationOutputProfile);
        Assert.Same(outputIntent, Assert.Single(report.DocumentInfo.GetOutputIntentsBySubtype("GTS_PDFA1")));
        Assert.Same(outputIntent, Assert.Single(report.DocumentInfo.GetOutputIntentsByOutputConditionIdentifier("sRGB")));
        Assert.False(report.DocumentInfo.HasXmpMetadata);
        Assert.False(report.DocumentInfo.HasCatalogUri);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsComplexOutputIntentPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexOutputIntentPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasOutputIntents);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOutputIntents);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.OutputIntents, "PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Inspect_ReadsGeneratedOutputIntentProfileMetadata() {
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetSrgbOutputIntent())
            .Paragraph(p => p.Text("Output intent readback."))
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);

        Assert.True(info.HasOutputIntents);
        Assert.True(info.HasReadableOutputIntents);
        Assert.Equal(1, info.OutputIntentCount);
        Assert.Equal(new[] { "GTS_PDFA1" }, info.OutputIntentSubtypes);
        Assert.Equal(new[] { PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier }, info.OutputConditionIdentifiers);
        PdfOutputIntentInfo outputIntent = Assert.Single(info.OutputIntents);
        Assert.Equal("GTS_PDFA1", outputIntent.Subtype);
        Assert.Equal(PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier, outputIntent.OutputConditionIdentifier);
        Assert.Equal("IEC 61966-2-1 Default RGB Colour Space - sRGB", outputIntent.OutputCondition);
        Assert.Equal("https://www.color.org", outputIntent.RegistryName);
        Assert.True(outputIntent.HasDestinationOutputProfile);
        Assert.NotNull(outputIntent.DestinationOutputProfileObjectNumber);
        Assert.Equal(3, outputIntent.DestinationOutputProfileColorComponents);
        Assert.Null(outputIntent.DestinationOutputProfileAlternateColorSpace);
        Assert.Null(outputIntent.DestinationOutputProfileFilter);
        Assert.True(outputIntent.DestinationOutputProfileSizeBytes > 128);
        Assert.Equal(outputIntent.DestinationOutputProfileSizeBytes, outputIntent.DestinationOutputProfileDeclaredSizeBytes);
        Assert.Equal("RGB ", outputIntent.DestinationOutputProfileColorSpace);
        Assert.True(outputIntent.DestinationOutputProfileHasIccSignature);
        Assert.Same(outputIntent, Assert.Single(info.GetOutputIntentsBySubtype("GTS_PDFA1")));
        Assert.Same(outputIntent, Assert.Single(info.GetOutputIntentsByOutputConditionIdentifier(PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier)));
        Assert.Empty(info.GetOutputIntentsBySubtype("GTS_PDFX"));
        Assert.Empty(info.GetOutputIntentsByOutputConditionIdentifier("Office profile"));
        Assert.Throws<ArgumentException>(() => info.GetOutputIntentsBySubtype(""));
        Assert.Throws<ArgumentException>(() => info.GetOutputIntentsByOutputConditionIdentifier(""));
    }

    [Fact]
    public void Inspect_DecodesFilteredOutputIntentProfileBeforeReadingIccHeader() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildFilteredOutputIntentProfilePdf());

        Assert.True(info.HasOutputIntents);
        Assert.True(info.HasReadableOutputIntents);
        PdfOutputIntentInfo outputIntent = Assert.Single(info.OutputIntents);
        Assert.True(outputIntent.HasDestinationOutputProfile);
        Assert.Equal(6, outputIntent.DestinationOutputProfileObjectNumber);
        Assert.Equal(3, outputIntent.DestinationOutputProfileColorComponents);
        Assert.Equal("[ASCIIHexDecode FlateDecode]", outputIntent.DestinationOutputProfileFilter);
        Assert.Equal(128, outputIntent.DestinationOutputProfileSizeBytes);
        Assert.Equal(128, outputIntent.DestinationOutputProfileDeclaredSizeBytes);
        Assert.Equal("RGB ", outputIntent.DestinationOutputProfileColorSpace);
        Assert.True(outputIntent.DestinationOutputProfileHasIccSignature);
    }


}
