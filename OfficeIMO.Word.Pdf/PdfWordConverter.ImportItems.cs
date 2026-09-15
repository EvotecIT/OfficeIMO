using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    internal static partial class PdfWordConverter {
        private static double GetImageSortY(PdfCore.PdfLogicalImage image) {
            if (image.PlacedY.HasValue || image.PlacedHeight.HasValue) {
                return image.PlacedY.GetValueOrDefault() + image.PlacedHeight.GetValueOrDefault();
            }

            return 0D;
        }

        private sealed class ImportItem {
            private ImportItem(ImportItemKind kind, double y, int sequence, int? readingOrderIndex) {
                Kind = kind;
                Y = y;
                Sequence = sequence;
                ReadingOrderIndex = readingOrderIndex;
            }

            public ImportItemKind Kind { get; }
            public double Y { get; }
            public int Sequence { get; }
            public int? ReadingOrderIndex { get; }
            public PdfCore.PdfLogicalHeading? Heading { get; private set; }
            public PdfCore.PdfLogicalParagraph? Paragraph { get; private set; }
            public PdfCore.PdfLogicalTextBlock? TextBlock { get; private set; }
            public PdfCore.PdfLogicalListItem? ListItem { get; private set; }
            public PdfCore.PdfLogicalTableExtraction? TableExtraction { get; private set; }
            public PdfCore.PdfLogicalImage? Image { get; private set; }
            public PdfCore.PdfImagePlacement? ImagePlacement { get; private set; }
            public PdfCore.PdfLogicalFormWidget? FormWidget { get; private set; }
            public PdfCore.PdfLogicalLinkAnnotation? Link { get; private set; }
            public string? LinkText { get; private set; }

            public static ImportItem ForHeading(PdfCore.PdfLogicalHeading heading, double y, int sequence, int? readingOrderIndex, PdfCore.PdfLogicalLinkAnnotation? link = null, string? linkText = null) =>
                new ImportItem(ImportItemKind.Heading, y, sequence, readingOrderIndex) { Heading = heading, Link = link, LinkText = linkText };

            public static ImportItem ForParagraph(PdfCore.PdfLogicalParagraph paragraph, double y, int sequence, int? readingOrderIndex, PdfCore.PdfLogicalLinkAnnotation? link = null, string? linkText = null) =>
                new ImportItem(ImportItemKind.Paragraph, y, sequence, readingOrderIndex) { Paragraph = paragraph, Link = link, LinkText = linkText };

            public static ImportItem ForTextBlock(PdfCore.PdfLogicalTextBlock block, double y, int sequence, int? readingOrderIndex, PdfCore.PdfLogicalLinkAnnotation? link = null, string? linkText = null) =>
                new ImportItem(ImportItemKind.TextBlock, y, sequence, readingOrderIndex) { TextBlock = block, Link = link, LinkText = linkText };

            public static ImportItem ForListItem(PdfCore.PdfLogicalListItem listItem, double y, int sequence, int? readingOrderIndex) =>
                new ImportItem(ImportItemKind.ListItem, y, sequence, readingOrderIndex) { ListItem = listItem };

            public static ImportItem ForTable(PdfCore.PdfLogicalTableExtraction table, double y, int sequence, int? readingOrderIndex) =>
                new ImportItem(ImportItemKind.Table, y, sequence, readingOrderIndex) { TableExtraction = table };

            public static ImportItem ForImage(PdfCore.PdfLogicalImage image, PdfCore.PdfImagePlacement? placement, double y, int sequence, int? readingOrderIndex) =>
                new ImportItem(ImportItemKind.Image, y, sequence, readingOrderIndex) { Image = image, ImagePlacement = placement };

            public static ImportItem ForFormWidget(PdfCore.PdfLogicalFormWidget widget, double y, int sequence, int? readingOrderIndex) =>
                new ImportItem(ImportItemKind.FormWidget, y, sequence, readingOrderIndex) { FormWidget = widget };

            public static ImportItem ForLink(PdfCore.PdfLogicalLinkAnnotation link, double y, int sequence, int? readingOrderIndex, string? linkText) =>
                new ImportItem(ImportItemKind.Link, y, sequence, readingOrderIndex) { Link = link, LinkText = linkText };
        }

        private enum ImportItemKind {
            Heading,
            Paragraph,
            TextBlock,
            ListItem,
            Table,
            Image,
            FormWidget,
            Link
        }
    }
}
