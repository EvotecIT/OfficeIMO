using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlEmbeddedImagesShareTheExpandedByteBudgetAcrossSrcDocDocuments() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(image);
        string nested = $"<html><body><img src=\"{dataUri}\"></body></html>";
        string html = $"<html><body><img src=\"{dataUri}\"><iframe srcdoc='{nested}'></iframe></body></html>";
        long aggregateLimit = image.LongLength * 2L - 1L;
        var inspectionOptions = new OfficeProvenanceOptions { MaxExpandedContainerBytes = aggregateLimit };
        var removalOptions = new OfficeProvenanceRemovalOptions();
        removalOptions.Limits.MaxExpandedContainerBytes = aggregateLimit;

        InvalidDataException inspectionException = Assert.Throws<InvalidDataException>(() =>
            HtmlProvenance.Inspect(html, inspectionOptions));
        InvalidDataException removalException = Assert.Throws<InvalidDataException>(() =>
            HtmlProvenance.Remove(html, removalOptions));

        Assert.Contains("expanded-container", inspectionException.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("expanded-container", removalException.Message, StringComparison.OrdinalIgnoreCase);
    }


}
