using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote;

/// <summary>Single owner for native named-page geometry used by writing and visual rendering.</summary>
internal static class OneNotePageGeometry {
    internal static (double Width, double Height) GetNamedSizeHalfInches(
        OneNotePageSize size,
        OneNotePageOrientation? orientation = null) {
        OfficePageSize physical;
        switch (size) {
            case OneNotePageSize.Statement: physical = OfficePageSizes.Statement; break;
            case OneNotePageSize.Letter: physical = OfficePageSizes.Letter; break;
            case OneNotePageSize.Tabloid: physical = OfficePageSizes.Tabloid; break;
            case OneNotePageSize.Legal: physical = OfficePageSizes.Legal; break;
            case OneNotePageSize.A3: physical = OfficePageSizes.A3; break;
            case OneNotePageSize.A4: physical = OfficePageSizes.A4; break;
            case OneNotePageSize.A5: physical = OfficePageSizes.A5; break;
            case OneNotePageSize.A6: physical = OfficePageSizes.A6; break;
            case OneNotePageSize.B4: physical = OfficePageSizes.B4Jis; break;
            case OneNotePageSize.B5: physical = OfficePageSizes.B5Jis; break;
            case OneNotePageSize.B6: physical = OfficePageSizes.B6Jis; break;
            case OneNotePageSize.JapanesePostcard: physical = OfficePageSizes.JapanesePostcard; break;
            case OneNotePageSize.IndexCard: physical = OfficePageSizes.IndexCard; break;
            case OneNotePageSize.Billfold: physical = OfficePageSizes.Billfold; break;
            default:
                throw new ArgumentOutOfRangeException(nameof(size), "Automatic and custom pages do not have named physical dimensions.");
        }
        double width = physical.ToPointWidth() / OneNotePageRenderer.PointsPerHalfInch;
        double height = physical.ToPointHeight() / OneNotePageRenderer.PointsPerHalfInch;
        if (orientation == OneNotePageOrientation.Landscape && height > width) (width, height) = (height, width);
        return (width, height);
    }

    internal static (double Width, double Height) GetNamedSizePoints(
        OneNotePageSize size,
        OneNotePageOrientation? orientation = null) {
        (double width, double height) = GetNamedSizeHalfInches(size, orientation);
        return (width * OneNotePageRenderer.PointsPerHalfInch, height * OneNotePageRenderer.PointsPerHalfInch);
    }

    internal static void NormalizeForWrite(OneNotePage page) {
        if (!page.PageSize.HasValue || page.PageSize == OneNotePageSize.Automatic) return;
        if (page.PageSize == OneNotePageSize.Custom) {
            if (!page.Width.HasValue || !page.Height.HasValue) {
                throw new OneNoteFormatException(
                    "ONENOTE_WRITE_PAGE_SIZE_DIMENSIONS",
                    "A custom native OneNote page requires both width and height in half-inch units.");
            }
            return;
        }
        (page.Width, page.Height) = GetNamedSizeHalfInches(page.PageSize.Value, page.Orientation);
    }
}
