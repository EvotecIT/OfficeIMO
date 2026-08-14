using System.IO.Compression;
using System.IO.Packaging;
using OfficeIMO;
using OfficeIMO.Epub;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using OfficeIMO.Security;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
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
    public void HtmlFileRemovalEscapesMultipleCharactersOutsideTheLegacyEncoding() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1252 = Encoding.GetEncoding(1252);
        string html = "<!doctype html><html><head><meta charset=\"windows-1252\"><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>&#x2603;middle&#x1F600;tail</body></html>";
        string inputPath = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.html");
        string outputPath = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.html");
        try {
            File.WriteAllBytes(inputPath, windows1252.GetBytes(html));

            HtmlProvenance.RemoveFile(inputPath, outputPath);
            string output = windows1252.GetString(File.ReadAllBytes(outputPath));

            Assert.Contains("&#x2603;middle&#x1F600;tail", output, StringComparison.OrdinalIgnoreCase);
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
    public void HtmlSrcsetPreservesLiteralCommasInsideEmbeddedSvgPayloads() {
        string svg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\"><!--foo,bar.png-->" +
            "<metadata><c2pa:manifest xmlns:c2pa=\"http://c2pa.org/manifest\">" +
            Convert.ToBase64String(CreateManifestStore()) +
            "</c2pa:manifest></metadata></svg>";
        string dataUri = "data:image/svg+xml," + Uri.EscapeDataString(svg).Replace("%2C", ",").Replace("%2c", ",");
        string html = $"<html><head></head><body><picture><source srcset=\"{dataUri} 1x\"></picture></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Contains(result.Before.Evidence, item => item.Location.Contains("[srcset]", StringComparison.Ordinal));
        Assert.Empty(result.After.Evidence);
        Assert.True(result.WasChanged);
    }

    [Fact]
    public void HtmlSanitizesConsecutiveDataUrisWhenTheFirstSrcsetCandidateHasNoDescriptor() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        string first = "data:image/png;base64," + Convert.ToBase64String(image);
        string second = "data:image/png;base64," + Convert.ToBase64String(image);
        string html = $"<html><head></head><body><picture><source srcset=\"{first}, {second} 2x\"></picture></body></html>";

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
    public void HtmlSanitizesNativeManifestInsideIframeSrcdoc() {
        string nested = $"<html><head><script type=\"application/c2pa\">{Convert.ToBase64String(CreateManifestStore())}</script></head><body></body></html>";
        string html = $"<html><head></head><body><iframe srcdoc=\"{System.Net.WebUtility.HtmlEncode(nested)}\"></iframe></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain("application/c2pa", Encoding.UTF8.GetString(result.ToArray()), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlSanitizesSvgFeImageDataUris() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(image);
        string html = $"<html><body><svg><filter><feImage href=\"{dataUri}\"></feImage></filter></svg></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain(dataUri, Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlDomElementsShareTheConfiguredContainerEntryLimit() {
        const string html = "<html><body><div></div><div></div><div></div><div></div></body></html>";
        var inspectionOptions = new OfficeProvenanceOptions { MaxContainerEntries = 5 };
        var removalOptions = new OfficeProvenanceRemovalOptions();
        removalOptions.Limits.MaxContainerEntries = 5;

        Assert.Throws<InvalidDataException>(() => HtmlProvenance.Inspect(html, inspectionOptions));
        Assert.Throws<InvalidDataException>(() => HtmlProvenance.Remove(html, removalOptions));
    }

    [Fact]
    public void HtmlDomPreflightIgnoresTagLikeRawTextAndComments() {
        const string html = "<html><head><script>const sample = '<div><span><img>';</script>" +
            "<style>/* <section><aside><main> */</style></head><body><!-- <article><header><footer> --></body></html>";
        var inspectionOptions = new OfficeProvenanceOptions { MaxContainerEntries = 5 };
        var removalOptions = new OfficeProvenanceRemovalOptions();
        removalOptions.Limits.MaxContainerEntries = 5;

        Assert.Empty(HtmlProvenance.Inspect(html, inspectionOptions).Evidence);
        Assert.False(HtmlProvenance.Remove(html, removalOptions).WasChanged);
    }

    [Fact]
    public void HtmlEmbeddedSvgRewriteDeclaresTheUtf8OutputEncoding() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1252 = Encoding.GetEncoding(1252);
        string svg = "<?xml version=\"1.0\" encoding=\"windows-1252\"?><svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:x=\"adobe:ns:meta/\" xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><title>café</title><metadata><x:xmpmeta><rdf:RDF><rdf:Description " +
            "iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/></rdf:RDF></x:xmpmeta></metadata></svg>";
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
    public void HtmlEmbeddedSvgHonorsTheDataUriCharsetWithoutAnXmlDeclaration() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1252 = Encoding.GetEncoding(1252);
        string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:x=\"adobe:ns:meta/\" xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><title>café</title><metadata><x:xmpmeta><rdf:RDF><rdf:Description " +
            "iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/></rdf:RDF></x:xmpmeta></metadata></svg>";
        string dataUri = "data:image/svg+xml;charset=windows-1252;base64," + Convert.ToBase64String(windows1252.GetBytes(svg));
        string html = $"<html><head></head><body><img src=\"{dataUri}\"></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());
        int start = output.IndexOf("data:image/svg+xml", StringComparison.Ordinal);
        int end = output.IndexOf('"', start);
        Assert.True(HtmlDataUri.TryParse(output.Substring(start, end - start), out HtmlDataUri rewritten));

        Assert.True(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.Empty(result.After.Evidence);
        Assert.Contains("charset=utf-8", rewritten.Metadata, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("café", rewritten.DecodeText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlEmbeddedSvgHonorsItsXmlDeclarationWhenTheDataUriHasNoCharset() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1252 = Encoding.GetEncoding(1252);
        string svg = "<?xml version=\"1.0\" encoding=\"windows-1252\"?><svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:x=\"adobe:ns:meta/\" xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><title>café</title><metadata><x:xmpmeta><rdf:RDF><rdf:Description " +
            "iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/></rdf:RDF></x:xmpmeta></metadata></svg>";
        string dataUri = "data:image/svg+xml;base64," + Convert.ToBase64String(windows1252.GetBytes(svg));
        string html = $"<html><head></head><body><img src=\"{dataUri}\"></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());
        int start = output.IndexOf("data:image/svg+xml", StringComparison.Ordinal);
        int end = output.IndexOf('"', start);
        Assert.True(HtmlDataUri.TryParse(output.Substring(start, end - start), out HtmlDataUri rewritten));

        Assert.True(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.Empty(result.After.Evidence);
        Assert.Contains("charset=utf-8", rewritten.Metadata, StringComparison.OrdinalIgnoreCase);
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
            using (Package packageHandle = Package.Open(path, FileMode.Open, FileAccess.ReadWrite)) {
                PackagePart signature = packageHandle.CreatePart(
                    PackUriHelper.CreatePartUri(new Uri("/_xmlsignatures/orphan.xml", UriKind.Relative)),
                    "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
                    CompressionOption.Normal);
                using Stream target = signature.GetStream();
                byte[] payload = Encoding.UTF8.GetBytes("<signature/>");
                target.Write(payload, 0, payload.Length);
            }
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
    public void OdfIgnoresCaseDistinctSignatureLikeResources() {
        const string resourcePath = "meta-inf/customsignatures.xml";
        byte[] package = CreateZipPackage("odt", resourcePath, CreatePngWithManifest(CreateManifestStore()));

        OfficeProvenanceRemovalResult result = OdfDocument.RemoveProvenance(package, "document.odt");

        using var archive = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);
        Assert.False(result.WereInvalidatedSignaturesRemoved);
        Assert.Contains(archive.Entries, entry => entry.FullName.Equals(resourcePath, StringComparison.Ordinal));
        Assert.Empty(result.After.Evidence);
    }

    [Theory]
    [InlineData("META-INF/customsignatures.xml")]
    [InlineData("META-INF/SIGNATURES.XML")]
    public void EpubIgnoresNonstandardSignatureLikeResourceNames(string resourcePath) {
        byte[] package = CreateZipPackage("epub", resourcePath, CreatePngWithManifest(CreateManifestStore()));

        OfficeProvenanceRemovalResult result = EpubDocument.RemoveProvenance(package, "publication.epub");

        using var archive = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);
        Assert.False(result.WereInvalidatedSignaturesRemoved);
        Assert.Contains(archive.Entries, entry => entry.FullName.Equals(resourcePath, StringComparison.Ordinal));
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void SignatureRemovalPreviewHonorsDisabledEmbeddedAssets() {
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
                WriteEntry(archive, "META-INF/manifest.xml", ValidOdfManifestXml, CompressionLevel.Optimal);
                WriteEntry(archive, "content.xml", "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"/>", CompressionLevel.Optimal);
                WriteEntry(archive, "META-INF/customsignatures.xml", "<signatures/>", CompressionLevel.Optimal);
                WriteEntry(archive, "META-INF/content_credential.c2pa", CreateManifestStore(), CompressionLevel.Optimal);
                WriteEntry(archive, "media/provenance.png", CreatePngWithManifest(CreateManifestStore()), CompressionLevel.Optimal);
            }
            package = RewriteFixtureWithStoredMimetype(output.ToArray());
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
                WriteEntry(archive, "META-INF/manifest.xml", ValidOdfManifestXml, CompressionLevel.Optimal);
                WriteEntry(archive, "META-INF/customsignatures.xml", "<signatures/>", CompressionLevel.Optimal);
                WriteEntry(archive, "media/first.png", image, CompressionLevel.Optimal);
                WriteEntry(archive, "media/second.png", image, CompressionLevel.Optimal);
            }
            package = RewriteFixtureWithStoredMimetype(output.ToArray());
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
    public void VisioSignatureMetadataUsesTheConfiguredAssetLimit() {
        byte[] package = CreateSignedVisioProvenancePackage(16 * 1024 * 1024 + 1);
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };
        options.Limits.MaxAssetBytes = 32L * 1024L * 1024L;
        options.Limits.MaxManifestBytes = 1024L * 1024L;
        options.Limits.MaxExpandedContainerBytes = 64L * 1024L * 1024L;

        OfficeProvenanceRemovalResult result = VisioDocument.RemoveProvenance(package, options: options);

        Assert.True(result.WereInvalidatedSignaturesRemoved);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void VisioSignatureRemovalSkipsExternalSignatureRelationships() {
        byte[] package = CreateSignedVisioProvenancePackage(0, includeExternalSignatureRelationship: true);
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };

        OfficeProvenanceRemovalResult result = VisioDocument.RemoveProvenance(package, options: options);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void VisioSignatureRemovalRejectsAnOriginRelationshipToDocumentContent() {
        byte[] package;
        using (var output = new MemoryStream()) {
            using (Package packageHandle = Package.Open(output, FileMode.Create, FileAccess.ReadWrite)) {
                Uri documentUri = PackUriHelper.CreatePartUri(new Uri("/visio/document.xml", UriKind.Relative));
                PackagePart document = packageHandle.CreatePart(documentUri, "application/vnd.ms-visio.drawing.main+xml", CompressionOption.Maximum);
                using (Stream target = document.GetStream()) {
                    byte[] content = Encoding.UTF8.GetBytes("<document>keep</document>");
                    target.Write(content, 0, content.Length);
                }
                packageHandle.CreateRelationship(
                    documentUri,
                    TargetMode.Internal,
                    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin");
                packageHandle.CreateRelationship(
                    documentUri,
                    TargetMode.Internal,
                    "http://schemas.microsoft.com/visio/2010/relationships/document");
                Uri manifestUri = PackUriHelper.CreatePartUri(new Uri("/META-INF/content_credential.c2pa", UriKind.Relative));
                using Stream manifestTarget = packageHandle.CreatePart(
                    manifestUri,
                    "application/c2pa",
                    CompressionOption.Maximum).GetStream();
                byte[] manifest = CreateManifestStore();
                manifestTarget.Write(manifest, 0, manifest.Length);
            }
            package = output.ToArray();
        }
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            VisioDocument.RemoveProvenance(package, options: options));

        Assert.Contains("signature-origin", exception.Message, StringComparison.OrdinalIgnoreCase);
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

    private static byte[] CreateSignedVisioProvenancePackage(int paddingCharacters, bool includeExternalSignatureRelationship = false) {
        using var output = new MemoryStream();
        using (Package package = Package.Open(output, FileMode.Create, FileAccess.ReadWrite)) {
            Uri documentUri = PackUriHelper.CreatePartUri(new Uri("/visio/document.xml", UriKind.Relative));
            using (Stream document = package.CreatePart(documentUri, "application/vnd.ms-visio.drawing.main+xml", CompressionOption.Maximum).GetStream()) {
                byte[] xml = Encoding.UTF8.GetBytes("<VisioDocument xmlns=\"http://schemas.microsoft.com/office/visio/2012/main\"/>");
                document.Write(xml, 0, xml.Length);
            }
            package.CreateRelationship(documentUri, TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/document");
            Uri manifestUri = PackUriHelper.CreatePartUri(new Uri("/META-INF/content_credential.c2pa", UriKind.Relative));
            using (Stream target = package.CreatePart(manifestUri, "application/c2pa", CompressionOption.Maximum).GetStream()) {
                byte[] manifest = CreateManifestStore();
                target.Write(manifest, 0, manifest.Length);
            }

            Uri originUri = PackUriHelper.CreatePartUri(new Uri("/_xmlsignatures/origin.sigs", UriKind.Relative));
            PackagePart origin = package.CreatePart(
                originUri,
                "application/vnd.openxmlformats-package.digital-signature-origin",
                CompressionOption.Maximum);
            package.CreateRelationship(
                originUri,
                TargetMode.Internal,
                "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin");
            Uri signatureUri = PackUriHelper.CreatePartUri(new Uri("/_xmlsignatures/sig1.xml", UriKind.Relative));
            PackagePart signature = package.CreatePart(
                signatureUri,
                "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
                CompressionOption.Maximum);
            using (Stream target = signature.GetStream()) {
                byte[] xml = Encoding.UTF8.GetBytes("<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"/>");
                target.Write(xml, 0, xml.Length);
            }
            origin.CreateRelationship(
                signatureUri,
                TargetMode.Internal,
                "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature");
            if (includeExternalSignatureRelationship) {
                origin.CreateRelationship(
                    new Uri("https://example.invalid/external-signature.xml", UriKind.Absolute),
                    TargetMode.External,
                    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature");
            }

            Uri appUri = PackUriHelper.CreatePartUri(new Uri("/docProps/app.xml", UriKind.Relative));
            PackagePart app = package.CreatePart(appUri, "application/vnd.openxmlformats-officedocument.extended-properties+xml", CompressionOption.Maximum);
            using (var writer = new StreamWriter(app.GetStream(), new UTF8Encoding(false), 4096, leaveOpen: false)) {
                writer.Write("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">");
                writer.Write(new string(' ', paddingCharacters));
                writer.Write("<DigSig>signature</DigSig></Properties>");
            }
        }
        return output.ToArray();
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
            if (extension == "odt") {
                WriteEntry(archive, "META-INF/manifest.xml", ValidOdfManifestXml, CompressionLevel.Optimal);
                WriteEntry(archive, "content.xml", "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"/>", CompressionLevel.Optimal);
            }
            WriteEntry(archive, signaturePath, "<signatures/>", CompressionLevel.Optimal);
            WriteEntry(archive, signaturePath, "<signatures duplicate=\"true\"/>", CompressionLevel.Optimal);
            WriteEntry(archive, "media/provenance.png", image, CompressionLevel.Optimal);
        }
        return RewriteFixtureWithStoredMimetype(output.ToArray());
    }

    private const string ValidOdfManifestXml =
        "<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\">" +
        "<manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/>" +
        "</manifest:manifest>";

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
        byte[] storeDescription = CreateBox("jumd", Join(C2paUuid("c2pa"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa\0")));
        byte[] manifestDescription = CreateBox("jumd", Join(C2paUuid("c2ma"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("m\0")));
        byte[] assertionStoreDescription = CreateBox("jumd", Join(C2paUuid("c2as"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.assertions\0")));
        byte[] assertionDescription = CreateBox("jumd", Join(C2paUuid("c2ac"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.test\0")));
        byte[] assertionStore = CreateBox("jumb", Join(assertionStoreDescription,
            CreateBox("jumb", Join(assertionDescription, CreateBox("cbor", new byte[] { 0xA0 })))));
        byte[] claimDescription = CreateBox("jumd", Join(C2paUuid("c2cl"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.claim\0")));
        byte[] claim = CreateBox("jumb", Join(claimDescription, CreateBox("cbor", new byte[] { 0xA0 })));
        byte[] signatureDescription = CreateBox("jumd", Join(C2paUuid("c2cs"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.signature\0")));
        byte[] signature = CreateBox("jumb", Join(signatureDescription, CreateBox("cbor", new byte[] { 0xA0 })));
        return CreateBox("jumb", Join(storeDescription, CreateBox("jumb", Join(manifestDescription, assertionStore, claim, signature))));
    }

    private static byte[] C2paUuid(string code) => Join(
        Encoding.ASCII.GetBytes(code),
        new byte[] { 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 });

    private static byte[] CreateBox(string type, byte[] payload) {
        byte[] box = new byte[payload.Length + 8];
        WriteBigEndian(box, 0, box.Length);
        Encoding.ASCII.GetBytes(type).CopyTo(box, 4);
        Buffer.BlockCopy(payload, 0, box, 8, payload.Length);
        return box;
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
