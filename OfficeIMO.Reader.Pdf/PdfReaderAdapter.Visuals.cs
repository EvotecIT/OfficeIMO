using OfficeIMO.Pdf;
using System.Globalization;

namespace OfficeIMO.Reader.Pdf;

internal static partial class PdfReaderAdapter {
    private static IReadOnlyList<ReaderVisual>? BuildVisuals(IReadOnlyList<PdfLogicalPage> pages, int? pageSelectionIndex = null) {
        if (pages.Count == 0) {
            return null;
        }

        var visuals = new List<ReaderVisual>();
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            PdfLogicalPage page = pages[pageIndex];
            int selectionIndex = pageSelectionIndex ?? pageIndex;
            for (int imageIndex = 0; imageIndex < page.Images.Count; imageIndex++) {
                visuals.Add(BuildVisual(page.Images[imageIndex], selectionIndex, imageIndex));
            }
        }

        return visuals.Count == 0 ? null : visuals.AsReadOnly();
    }

    private static ReaderVisual BuildVisual(PdfLogicalImage image, int selectionIndex, int imageIndex) {
        PdfImagePlacement? placement = image.PrimaryPlacement;
        string content = "PDF image " + image.ResourceName +
            " " + image.Width.ToString(CultureInfo.InvariantCulture) +
            "x" + image.Height.ToString(CultureInfo.InvariantCulture) +
            " placements=" + image.PlacementCount.ToString(CultureInfo.InvariantCulture);

        return new ReaderVisual {
            Kind = "image",
            Language = "pdf-image",
            Content = content,
            PayloadHash = ComputeSha256Hex(content),
            Location = new ReaderLocation {
                Page = image.PageNumber,
                SourceBlockKind = "image",
                BlockAnchor = "page-" + image.PageNumber.ToString(CultureInfo.InvariantCulture)
                    + "-selection-" + selectionIndex.ToString("D4", CultureInfo.InvariantCulture)
                    + "-image-" + imageIndex.ToString(CultureInfo.InvariantCulture)
            },
            SourceName = image.ResourceName,
            MimeType = image.MimeType,
            Width = image.Width,
            Height = image.Height,
            X = placement?.X,
            Y = placement?.Y,
            PlacedWidth = placement?.Width,
            PlacedHeight = placement?.Height,
            PlacementCount = image.PlacementCount,
            HasGeometry = placement is not null,
            IsAxisAligned = placement?.IsAxisAligned
        };
    }
}
