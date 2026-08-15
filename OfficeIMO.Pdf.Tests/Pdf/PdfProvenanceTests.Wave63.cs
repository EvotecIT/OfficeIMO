using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActivePageViewportGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var viewports = new PdfArray();
            viewports.Items.Add(candidate);
            page.Items["VP"] = viewports;
        });
    }

    [Fact]
    public void RemovalValidatesLimitsBeforeDerivingAttachmentBudgets() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxCarriers = 0;

        Assert.Throws<ArgumentOutOfRangeException>(() => PdfProvenance.Remove(pdf, options));
    }
}
