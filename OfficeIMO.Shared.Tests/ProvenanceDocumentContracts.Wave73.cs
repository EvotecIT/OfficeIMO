using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void SuccessfulObjectImageSuppressesFallbackImageProvenance() {
        byte[] carrier = CreatePngWithManifest(CreateManifestStore());
        byte[] clean = OfficeProvenanceRemover.Remove(carrier, "object.png").ToArray();
        string html = "<object type='image/png' data='data:image/png;base64," + Convert.ToBase64String(clean) +
            "'><img src='data:image/png;base64," + Convert.ToBase64String(carrier) + "'></object>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void PseudoElementComputedStyleOwnsActiveImageCarrier() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<style>.box{background-image:none}.box::before{content:'';background-image:url('" +
            dataUri + "')}</style><div class='box'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void DisplayNoneCssImageCarrierIsInactive() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<style>.box{display:none;background-image:url('" + dataUri +
            "')}</style><div class='box'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void InvalidImageSetOptionIsInactive() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<style>.box{background-image:image-set(\"" + dataUri +
            "\" 1x 2x)}</style><div class='box'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }
}
