using System.Threading;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointPresentation {
    /// <summary>Adds slides in source order using one layout, with one presentation-list save for the batch.</summary>
    /// <param name="count">The number of slides to add.</param>
    /// <param name="masterIndex">Index of the slide master.</param>
    /// <param name="layoutIndex">Index of the slide layout.</param>
    /// <param name="cancellationToken">Cancels between completed slide additions.</param>
    /// <remarks>Completed slides remain in the presentation if cancellation interrupts the batch. Each invocation observes current Open XML relationships, including changes made through OpenXmlDocument.</remarks>
    public IReadOnlyList<PowerPointSlide> AddSlides(int count, int masterIndex = 0, int layoutIndex = 0,
        CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
        cancellationToken.ThrowIfCancellationRequested();
        if (count == 0) return Array.Empty<PowerPointSlide>();
        SlideLayoutPart layoutPart = GetSlideLayoutPart(masterIndex, layoutIndex);
        uint nextSlideId = GetNextSlideId();
        if ((ulong)nextSlideId + (ulong)count - 1 > 2147483647U) {
            throw new InvalidOperationException("The requested batch exceeds the supported slide ID range.");
        }
        HashSet<string> relationships = GetPresentationRelationships();
        long nextRelationship = 1;
        var added = new List<PowerPointSlide>(count);
        try {
            for (int index = 0; index < count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                string relationship = ReserveSlideRelationshipId(relationships, ref nextRelationship);
                added.Add(AddSlideCore(layoutPart, relationship, nextSlideId++, savePresentation: false));
            }
        } finally {
            PresentationRoot.Save();
        }
        return added.AsReadOnly();
    }
}
