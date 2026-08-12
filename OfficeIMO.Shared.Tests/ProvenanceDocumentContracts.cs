using System.IO.Compression;
using OfficeIMO;
using OfficeIMO.Epub;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using OfficeIMO.Security;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlRemovesNativeCarriersAndEmbeddedImageProvenanceOffline() {
        byte[] manifest = CreateManifestStore();
        byte[] image = CreatePngWithManifest(manifest);
        string html = "<!doctype html><html><head>" +
            "<link rel=\"stylesheet c2pa-manifest\" href=\"https://example.test/claim.c2pa\">" +
            "</head><body>" +
            $"<img src=\"data:image/png;base64,{Convert.ToBase64String(image)}\" alt=\"kept\">" +
            "</body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());

        Assert.Equal(2, report.Evidence.Count);
        Assert.Contains(report.Evidence, item => item.Location.StartsWith("HTML/img[src]", StringComparison.Ordinal));
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain("c2pa-manifest", output, StringComparison.OrdinalIgnoreCase);
        Assert.StartsWith("<!DOCTYPE html>", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("rel=\"stylesheet\"", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("href=\"https://example.test/claim.c2pa\"", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("alt=\"kept\"", output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRemovesASingleEmbeddedManifestAssociation() {
        string html = $"<!doctype html><html><head><script type=\"application/c2pa\">{Convert.ToBase64String(CreateManifestStore())}</script></head><body>kept</body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain("application/c2pa", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("kept", output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPreservesMultipleManifestAssociationsByDefault() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string html = $"<html><head><script type=\"application/c2pa\">{manifest}</script><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.False(result.WasChanged);
        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, item => Assert.False(item.IsStructurallyValid));
        Assert.Contains(result.Before.Diagnostics, item => item.Contains("manifest.html.multipleManifests", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlPreservesMalformedNativeCarrierByDefault() {
        const string html = "<html><head><script type=\"application/c2pa\">not-base64</script></head><body>ok</body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.False(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.False(result.Before.Evidence[0].IsStructurallyValid);
        Assert.Equal(html, Encoding.UTF8.GetString(result.ToArray()));
    }

    [Fact]
    public void HtmlUsesOnlyHeadAssociationsAndAcceptsSafeRelativeReferences() {
        string html = "<html><head><link rel=\"c2pa-manifest\" href=\"claims/active.c2pa\"></head>" +
            "<body><script type=\"application/c2pa\">not-a-head-carrier</script></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());

        OfficeProvenanceEvidence evidence = Assert.Single(report.Evidence);
        Assert.True(evidence.IsStructurallyValid);
        Assert.Equal("claims/active.c2pa", evidence.Value);
        Assert.DoesNotContain("rel=\"c2pa-manifest\"", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("not-a-head-carrier", output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlFileRemovalPreservesDetectedLegacyEncoding() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1252 = Encoding.GetEncoding(1252);
        string html = "<!doctype html><html><head><meta charset=\"windows-1252\"><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>café</body></html>";
        string inputPath = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.html");
        string outputPath = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.html");
        try {
            File.WriteAllBytes(inputPath, windows1252.GetBytes(html));

            OfficeProvenanceReport report = HtmlProvenance.InspectFile(inputPath);
            OfficeProvenanceRemovalResult result = HtmlProvenance.RemoveFile(inputPath, outputPath);
            string output = windows1252.GetString(File.ReadAllBytes(outputPath));

            Assert.Single(report.Evidence);
            Assert.True(result.WasChanged);
            Assert.Contains("café", output, StringComparison.Ordinal);
            Assert.DoesNotContain("c2pa-manifest", output, StringComparison.OrdinalIgnoreCase);
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void HtmlFileRemovalEscapesCharactersOutsideTheLegacyEncoding() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1252 = Encoding.GetEncoding(1252);
        string html = "<!doctype html><html><head><meta charset=\"windows-1252\"><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>&#x2603;</body></html>";
        string inputPath = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.html");
        string outputPath = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.html");
        try {
            File.WriteAllBytes(inputPath, windows1252.GetBytes(html));

            HtmlProvenance.RemoveFile(inputPath, outputPath);
            string output = windows1252.GetString(File.ReadAllBytes(outputPath));

            Assert.Contains("&#x2603;", output, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("?", output, StringComparison.Ordinal);
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void HtmlFileInspectionUsesTheBoundedSourceEncodingSize() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1252 = Encoding.GetEncoding(1252);
        string html = "<html><head><meta charset=\"windows-1252\"></head><body>café</body></html>";
        byte[] data = windows1252.GetBytes(html);
        string inputPath = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.html");
        try {
            File.WriteAllBytes(inputPath, data);

            OfficeProvenanceReport report = HtmlProvenance.InspectFile(inputPath, new OfficeProvenanceOptions {
                MaxAssetBytes = data.Length,
                MaxManifestBytes = data.Length
            });

            Assert.Equal(OfficeProvenanceAssetFormat.Html, report.Format);
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
        }
    }

    [Fact]
    public void HtmlMalformedEmbeddedDataUriIsDiagnosticInsteadOfAnException() {
        const string html = "<html><head></head><body><img src=\"data:image/png,%ZZ\"></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Contains(report.Diagnostics, item => item.Contains("could not be decoded", StringComparison.Ordinal));
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void HtmlMalformedEmbeddedSvgIsDiagnosticInsteadOfAnException() {
        byte[] malformedSvg = Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\"><broken></svg>");
        string html = $"<html><head></head><body><img src=\"data:image/svg+xml;base64,{Convert.ToBase64String(malformedSvg)}\"></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Contains(report.Diagnostics, item => item.Contains("embedded image was preserved", StringComparison.Ordinal));
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void HtmlSanitizesEmbeddedImagesInResponsiveSourceSets() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(image);
        string html = $"<html><head></head><body><picture><source srcset=\"{dataUri} 1x, image.png 2x\"></picture></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());

        Assert.Contains(result.Before.Evidence, item => item.Location.Contains("[srcset]", StringComparison.Ordinal));
        Assert.Empty(result.After.Evidence);
        Assert.Contains("image.png 2x", output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlSanitizesConsecutiveDataUrisWhenTheFirstSrcsetCandidateHasNoDescriptor() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        string first = "data:image/png;base64," + Convert.ToBase64String(image);
        string second = "data:image/png;base64," + Convert.ToBase64String(image);
        string html = $"<html><head></head><body><source srcset=\"{first}, {second} 2x\"></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.Empty(result.After.Evidence);
        Assert.Equal(2, result.Changes.Count);
        Assert.All(result.Changes, change => Assert.Equal(0, change.RemovedBytes));
    }

    [Fact]
    public void HtmlRemovalSkipsEmbeddedAssetsWhenDisabled() {
        string html = "<html><head><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head>" +
            "<body><img src=\"data:image/png;base64," + new string('A', 512) + "\"></body></html>";
        var options = new OfficeProvenanceRemovalOptions { ProcessEmbeddedAssets = false };
        options.Limits.MaxAssetBytes = Encoding.UTF8.GetByteCount(html) + 32;
        options.Limits.MaxManifestBytes = 32;

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html, options);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void HtmlRemovalHonorsTheNestedEmbeddedAssetSwitch() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(image);
        string html = $"<html><head><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body><img src=\"{dataUri}\"></body></html>";
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.ProcessEmbeddedAssets = false;

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html, options);
        string output = Encoding.UTF8.GetString(result.ToArray());

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
        Assert.Contains(dataUri, output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlEmbeddedRewritePreservesDataUriMediaTypeParameters() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        string dataUri = "data:image/png;charset=utf-8;name=source.png;base64," + Convert.ToBase64String(image);
        string html = $"<html><head></head><body><img src=\"{dataUri}\"></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());

        Assert.Contains("data:image/png;charset=utf-8;name=source.png;base64,", output, StringComparison.Ordinal);
        Assert.Empty(result.After.Evidence);
    }

    [Theory]
    [InlineData("data-src", false)]
    [InlineData("data-original", false)]
    [InlineData("data-original-src", false)]
    [InlineData("data-lazy-src", false)]
    [InlineData("data-srcset", true)]
    [InlineData("data-original-srcset", true)]
    [InlineData("data-lazy-srcset", true)]
    public void HtmlSanitizesSupportedLazyImageAttributes(string attributeName, bool sourceSet) {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(image);
        string value = sourceSet ? dataUri + " 1x, retained.png 2x" : dataUri;
        string html = $"<html><head></head><body><img {attributeName}=\"{value}\"></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());

        Assert.Contains(result.Before.Evidence, item => item.Location.Contains("[" + attributeName + "]", StringComparison.Ordinal));
        Assert.Empty(result.After.Evidence);
        if (sourceSet) Assert.Contains("retained.png 2x", output, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("<html><body><video poster=\"{0}\"></video></body></html>", "poster")]
    [InlineData("<html><body><input type=\"image\" src=\"{0}\"></body></html>", "src")]
    [InlineData("<html><body><svg><image href=\"{0}\"/></svg></body></html>", "href")]
    [InlineData("<html><head><link rel=\"icon\" href=\"{0}\"></head><body></body></html>", "href")]
    [InlineData("<html><body><table background=\"{0}\"></table></body></html>", "background")]
    [InlineData("<html><head><link rel=\"preload\" as=\"image\" href=\"{0}\"></head><body></body></html>", "href")]
    [InlineData("<html><head><link rel=\"preload\" as=\"image\" href=\"keep.png\" imagesrcset=\"{0} 1x, keep2.png 2x\"></head><body></body></html>", "imagesrcset")]
    [InlineData("<html><body><div style=\"background-image:url('{0}')\"></div></body></html>", "style")]
    [InlineData("<html><head><style>.x{{background-image:image-set(\"{0}\" 1x)}}</style></head><body class=\"x\"></body></html>", "css")]
    public void HtmlSanitizesEverySupportedEmbeddedImageCarrier(string template, string attributeName) {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(image);
        string html = string.Format(System.Globalization.CultureInfo.InvariantCulture, template, dataUri);

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Contains(result.Before.Evidence, item => item.Location.Contains("[" + attributeName + "]", StringComparison.Ordinal));
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain(dataUri, Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlImageSetUrlReferenceIsEnumeratedOnlyOnce() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(image);
        string html = $"<html><head><style>.x{{background-image:image-set(url('{dataUri}') 1x)}}</style></head><body class=\"x\"></body></html>";
        var options = new OfficeProvenanceRemovalOptions { MaxEmbeddedAssets = 1 };
        options.Limits.MaxEmbeddedAssets = 1;

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html, options);

        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void HtmlSanitizesUsedCssCustomPropertyImage() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(image);
        string html = $"<html><head><style>:root{{--hero:url('{dataUri}')}}.x{{background-image:var(--hero)}}</style></head><body class=\"x\"></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain(dataUri, Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlSanitizesCssCustomPropertyUsedAcrossStyleBlocks() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(image);
        string html = $"<html><head><style>:root{{--hero:url('{dataUri}')}}</style><style>.x{{background-image:var(--hero)}}</style></head><body class=\"x\"></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain(dataUri, Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlSanitizesTransitiveCssCustomPropertyImageAcrossStyleBlocks() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(image);
        string html = $"<html><head><style>:root{{--source:url('{dataUri}')}}</style><style>:root{{--hero:var(--source)}}</style><style>.x{{background-image:var(--hero)}}</style></head><body class=\"x\"></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain(dataUri, Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlSanitizesImageInsideIframeSrcdoc() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(image);
        string nested = $"<html><body><img src=\"{dataUri}\"></body></html>";
        string html = $"<html><head></head><body><iframe srcdoc=\"{System.Net.WebUtility.HtmlEncode(nested)}\"></iframe></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain(dataUri, Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlEmbeddedSvgRewriteDeclaresTheUtf8OutputEncoding() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1252 = Encoding.GetEncoding(1252);
        string svg = "<?xml version=\"1.0\" encoding=\"windows-1252\"?><svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><title>café</title><metadata><rdf:Description " +
            "iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/></metadata></svg>";
        string dataUri = "data:image/svg+xml;charset=windows-1252;base64," + Convert.ToBase64String(windows1252.GetBytes(svg));
        string html = $"<html><head></head><body><img src=\"{dataUri}\"></body></html>";

        string output = Encoding.UTF8.GetString(HtmlProvenance.Remove(html).ToArray());
        int start = output.IndexOf("data:image/svg+xml", StringComparison.Ordinal);
        int end = output.IndexOf('"', start);
        Assert.True(HtmlDataUri.TryParse(output.Substring(start, end - start), out HtmlDataUri rewritten));

        Assert.Contains("charset=utf-8", rewritten.Metadata, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("windows-1252", rewritten.Metadata, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("café", rewritten.DecodeText(), StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownUsesTheSharedStructuredTextContract() {
        string markdown = "# Before\n\n-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\n" +
            "-----END C2PA MANIFEST-----\n\nAfter\n";

        OfficeProvenanceRemovalResult result = MarkdownProvenance.Remove(markdown);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
        Assert.Equal("# Before\n\n\nAfter\n", Encoding.UTF8.GetString(result.ToArray()));
    }

    [Theory]
    [InlineData("docx")]
    [InlineData("xlsx")]
    [InlineData("pptx")]
    public void OpenXmlOwnerApisSanitizeEmbeddedImages(string extension) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.{extension}");
        try {
            CreateOpenXmlPackage(path, extension);
            AddZipEntry(path, "media/provenance.png", CreatePngWithManifest(CreateManifestStore()));
            byte[] package = File.ReadAllBytes(path);

            OfficeProvenanceReport report = InspectOpenXml(path, extension);
            OfficeProvenanceRemovalResult result = RemoveOpenXml(package, extension);

            Assert.Single(report.Evidence);
            Assert.True(result.WasChanged);
            Assert.Empty(result.After.Evidence);
            Assert.Empty(OfficeProvenanceInspector.Inspect(ReadZipEntry(result.ToArray(), "media/provenance.png"), "image.png").Evidence);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void GenericZipPreservePolicyDoesNotParseMalformedOpcSignatureMetadata() {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteEntry(archive, "[Content_Types].xml", "<Types", CompressionLevel.Optimal);
            WriteEntry(archive, "media/provenance.png", CreatePngWithManifest(CreateManifestStore()), CompressionLevel.Optimal);
        }

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(output.ToArray(), "package.zip", new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.PreserveSignatureMarkup
        });

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void OpenXmlOwnerFailsClosedWhenOrphanSignatureEvidenceCannotBeRemoved() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.docx");
        try {
            CreateOpenXmlPackage(path, "docx");
            AddZipEntry(path, "_xmlsignatures/orphan.xml", Encoding.UTF8.GetBytes("<signature/>"));
            AddZipEntry(path, "media/provenance.png", CreatePngWithManifest(CreateManifestStore()));
            byte[] package = File.ReadAllBytes(path);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                WordDocument.RemoveProvenance(package, options: new OfficeProvenanceRemovalOptions {
                    SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
                }));

            Assert.Contains("could not remove", exception.Message, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData("odt", "META-INF/customsignatures.xml")]
    [InlineData("epub", "META-INF/signatures.xml")]
    public void ZipDocumentOwnersRemoveInvalidatedNativeSignatures(string extension, string signaturePath) {
        byte[] package = CreateZipPackage(extension, signaturePath, CreatePngWithManifest(CreateManifestStore()));
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };

        OfficeProvenanceRemovalResult result = extension == "odt"
            ? OdfDocument.RemoveProvenance(package, "document.odt", options)
            : EpubDocument.RemoveProvenance(package, "publication.epub", options);

        using var archive = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);
        Assert.True(result.WereInvalidatedSignaturesRemoved);
        Assert.DoesNotContain(archive.Entries, entry => entry.FullName.Equals(signaturePath, StringComparison.OrdinalIgnoreCase));
        Assert.Empty(result.After.Evidence);
        Assert.Equal("mimetype", archive.Entries[0].FullName);
        Assert.Equal(CompressionMethodStored, ReadCompressionMethod(result.ToArray(), archive.Entries[0].FullName));
    }

    [Fact]
    public void OdfDefaultPolicyBlocksProducerSpecificNativeSignature() {
        byte[] package = CreateZipPackage("odt", "META-INF/customsignatures.xml", CreatePngWithManifest(CreateManifestStore()));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            OdfDocument.RemoveProvenance(package, "document.odt"));

        Assert.Contains("invalidate package signatures", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void EpubIgnoresNonstandardSignatureLikeResourceNames() {
        byte[] package = CreateZipPackage("epub", "META-INF/customsignatures.xml", CreatePngWithManifest(CreateManifestStore()));

        OfficeProvenanceRemovalResult result = EpubDocument.RemoveProvenance(package, "publication.epub");

        using var archive = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);
        Assert.False(result.WereInvalidatedSignaturesRemoved);
        Assert.Contains(archive.Entries, entry => entry.FullName.Equals("META-INF/customsignatures.xml", StringComparison.OrdinalIgnoreCase));
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void SignatureRemovalPreviewHonorsDisabledEmbeddedAssets() {
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
                WriteEntry(archive, "META-INF/customsignatures.xml", "<signatures/>", CompressionLevel.Optimal);
                WriteEntry(archive, "META-INF/content_credential.c2pa", CreateManifestStore(), CompressionLevel.Optimal);
                WriteEntry(archive, "media/provenance.png", CreatePngWithManifest(CreateManifestStore()), CompressionLevel.Optimal);
            }
            package = output.ToArray();
        }
        var options = new OfficeProvenanceRemovalOptions {
            ProcessEmbeddedAssets = false,
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };

        OfficeProvenanceRemovalResult result = OdfDocument.RemoveProvenance(package, "document.odt", options);

        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
        Assert.True(result.WereInvalidatedSignaturesRemoved);
    }

    [Fact]
    public void SignatureRemovalPreviewHonorsNestedEmbeddedAssetLimit() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
                WriteEntry(archive, "META-INF/customsignatures.xml", "<signatures/>", CompressionLevel.Optimal);
                WriteEntry(archive, "media/first.png", image, CompressionLevel.Optimal);
                WriteEntry(archive, "media/second.png", image, CompressionLevel.Optimal);
            }
            package = output.ToArray();
        }
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };
        options.Limits.MaxEmbeddedAssets = 1;

        Assert.Throws<InvalidDataException>(() => OdfDocument.RemoveProvenance(package, "document.odt", options));
    }

    [Fact]
    public void SignatureStripAdapterRejectsOversizedExpandedPackagePart() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "[Content_Types].xml", new string('A', 128 * 1024), CompressionLevel.Optimal);
                WriteEntry(archive, "_xmlsignatures/orphan.xml", "<signature/>", CompressionLevel.Optimal);
                WriteEntry(archive, "word/media/provenance.png", image, CompressionLevel.Optimal);
            }
            package = output.ToArray();
        }
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };
        options.Limits.MaxAssetBytes = Math.Max(package.LongLength + 1024, 16 * 1024);
        options.Limits.MaxManifestBytes = 8 * 1024;
        options.Limits.MaxExpandedContainerBytes = 512 * 1024;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            WordDocument.RemoveProvenance(package, "document.docx", options));

        Assert.Contains("package part exceeds", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DanglingSignatureOriginRelationshipIsSignatureEvidence() {
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "[Content_Types].xml",
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/></Types>",
                    CompressionLevel.Optimal);
                WriteEntry(archive, "_rels/.rels",
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"sig\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin\" Target=\"_xmlsignatures/missing-origin.sigs\"/></Relationships>",
                    CompressionLevel.Optimal);
            }
            package = output.ToArray();
        }

        OfficePackageSignatureInfo info = OfficePackageSignatureService.Inspect(package);

        Assert.Equal(1, info.OriginRelationshipCount);
        Assert.Equal(0, info.OriginPartCount);
        Assert.True(info.HasSignatures);
    }

    [Fact]
    public void RemovalResultPreservesTheFiveArgumentBinaryConstructor() {
        Type[] signature = {
            typeof(byte[]),
            typeof(OfficeProvenanceReport),
            typeof(OfficeProvenanceReport),
            typeof(IReadOnlyList<OfficeProvenanceChange>),
            typeof(bool)
        };

        Assert.NotNull(typeof(OfficeProvenanceRemovalResult).GetConstructor(signature));
    }

    private const ushort CompressionMethodStored = 0;

    private static void CreateOpenXmlPackage(string path, string extension) {
        switch (extension) {
            case "docx":
                using (WordDocument document = WordDocument.Create(path)) {
                    document.AddParagraph("provenance fixture");
                    document.Save();
                }
                break;
            case "xlsx":
                using (ExcelDocument document = ExcelDocument.Create(path)) {
                    document.AddWorksheet("Data").CellValue(1, 1, "provenance fixture");
                    document.Save();
                }
                break;
            case "pptx":
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) presentation.Save();
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(extension));
        }
    }

    private static OfficeProvenanceReport InspectOpenXml(string path, string extension) => extension switch {
        "docx" => WordDocument.InspectProvenance(path),
        "xlsx" => ExcelDocument.InspectProvenance(path),
        "pptx" => PowerPointPresentation.InspectProvenance(path),
        _ => throw new ArgumentOutOfRangeException(nameof(extension))
    };

    private static OfficeProvenanceRemovalResult RemoveOpenXml(byte[] package, string extension) => extension switch {
        "docx" => WordDocument.RemoveProvenance(package),
        "xlsx" => ExcelDocument.RemoveProvenance(package),
        "pptx" => PowerPointPresentation.RemoveProvenance(package),
        _ => throw new ArgumentOutOfRangeException(nameof(extension))
    };

    private static void AddZipEntry(string path, string entryName, byte[] data) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using Stream output = entry.Open();
        output.Write(data, 0, data.Length);
    }

    private static byte[] CreateZipPackage(string extension, string signaturePath, byte[] image) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteEntry(archive, "mimetype", extension == "odt" ? "application/vnd.oasis.opendocument.text" : "application/epub+zip", CompressionLevel.NoCompression);
            WriteEntry(archive, signaturePath, "<signatures/>", CompressionLevel.Optimal);
            WriteEntry(archive, signaturePath, "<signatures duplicate=\"true\"/>", CompressionLevel.Optimal);
            WriteEntry(archive, "media/provenance.png", image, CompressionLevel.Optimal);
        }
        return output.ToArray();
    }

    private static void WriteEntry(ZipArchive archive, string name, string content, CompressionLevel level) =>
        WriteEntry(archive, name, Encoding.UTF8.GetBytes(content), level);

    private static void WriteEntry(ZipArchive archive, string name, byte[] content, CompressionLevel level) {
        ZipArchiveEntry entry = archive.CreateEntry(name, level);
        using Stream stream = entry.Open();
        stream.Write(content, 0, content.Length);
    }

    private static byte[] ReadZipEntry(byte[] package, string name) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using Stream input = (archive.GetEntry(name) ?? throw new InvalidOperationException("Missing ZIP entry: " + name)).Open();
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static ushort ReadCompressionMethod(byte[] package, string entryName) {
        byte[] name = Encoding.UTF8.GetBytes(entryName);
        int offset = 0;
        while (offset <= package.Length - 30) {
            if (BitConverter.ToUInt32(package, offset) != 0x04034B50) break;
            ushort method = BitConverter.ToUInt16(package, offset + 8);
            ushort nameLength = BitConverter.ToUInt16(package, offset + 26);
            ushort extraLength = BitConverter.ToUInt16(package, offset + 28);
            string currentName = Encoding.UTF8.GetString(package, offset + 30, nameLength);
            if (currentName == entryName) return method;
            uint compressedLength = BitConverter.ToUInt32(package, offset + 18);
            offset += 30 + nameLength + extraLength + checked((int)compressedLength);
        }
        throw new InvalidDataException("ZIP local header was not found: " + entryName);
    }

    private static byte[] CreatePngWithManifest(byte[] manifest) {
        byte[] header = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        return Join(
            header,
            CreatePngChunk("IHDR", new byte[13]),
            CreatePngChunk("caBX", manifest),
            CreatePngChunk("IEND", Array.Empty<byte>()));
    }

    private static byte[] CreateManifestStore() {
        byte[] data = new byte[126];
        WriteBigEndian(data, 0, data.Length);
        Encoding.ASCII.GetBytes("jumb").CopyTo(data, 4);
        WriteBigEndian(data, 8, 30);
        Encoding.ASCII.GetBytes("jumd").CopyTo(data, 12);
        new byte[] { 0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 }.CopyTo(data, 16);
        data[32] = 0x02;
        Encoding.ASCII.GetBytes("c2pa").CopyTo(data, 33);
        WriteBigEndian(data, 38, data.Length - 38);
        Encoding.ASCII.GetBytes("jumb").CopyTo(data, 42);
        WriteBigEndian(data, 46, 27);
        Encoding.ASCII.GetBytes("jumd").CopyTo(data, 50);
        new byte[] { 0x63, 0x32, 0x6D, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 }.CopyTo(data, 54);
        data[70] = 0x02;
        data[71] = (byte)'m';
        WriteBigEndian(data, 73, 53);
        Encoding.ASCII.GetBytes("jumb").CopyTo(data, 77);
        WriteBigEndian(data, 81, 36);
        Encoding.ASCII.GetBytes("jumd").CopyTo(data, 85);
        new byte[] { 0x63, 0x32, 0x63, 0x6C, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 }.CopyTo(data, 89);
        data[105] = 0x02;
        Encoding.ASCII.GetBytes("c2pa.claim").CopyTo(data, 106);
        WriteBigEndian(data, 117, 9);
        Encoding.ASCII.GetBytes("cbor").CopyTo(data, 121);
        return data;
    }

    private static byte[] CreatePngChunk(string type, byte[] payload) {
        byte[] chunk = new byte[payload.Length + 12];
        WriteBigEndian(chunk, 0, payload.Length);
        Encoding.ASCII.GetBytes(type).CopyTo(chunk, 4);
        payload.CopyTo(chunk, 8);
        WriteBigEndian(chunk, chunk.Length - 4, unchecked((int)ComputePngCrc(chunk, 4, payload.Length + 4)));
        return chunk;
    }

    private static uint ComputePngCrc(byte[] data, int offset, int count) {
        uint crc = 0xFFFFFFFF;
        for (int index = offset; index < offset + count; index++) {
            crc ^= data[index];
            for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? 0xEDB88320U ^ (crc >> 1) : crc >> 1;
        }
        return crc ^ 0xFFFFFFFF;
    }

    private static void WriteBigEndian(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private static byte[] Join(params byte[][] arrays) {
        byte[] output = new byte[arrays.Sum(item => item.Length)];
        int offset = 0;
        foreach (byte[] item in arrays) {
            Buffer.BlockCopy(item, 0, output, offset, item.Length);
            offset += item.Length;
        }
        return output;
    }
}
