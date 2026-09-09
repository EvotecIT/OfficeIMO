using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

internal static partial class PdfWordConverter {
    internal static PdfWordConversionResult ConvertVisualPages(PdfCore.PdfDocument source, PdfToWordOptions options) {
        var token = options.CancellationToken;
        token.ThrowIfCancellationRequested();
        var renderOptions = new PdfCore.PdfPageRenderOptions {
            Dpi = options.Dpi, MaxPages = options.MaxPages, MaxPixelsPerPage = options.MaxPixelsPerPage,
            MaxOutputBytesPerPage = options.MaxOutputBytesPerPage, MaxTotalOutputBytes = options.MaxTotalOutputBytes,
            ContinueOnError = false
        };
        var pages = source.Render.Pages(options.ReadOptions?.PageSelection, renderOptions, token);
        if (pages.Count == 0) throw new InvalidOperationException("Select at least one PDF page for visual Word conversion.");
        WordDocument target = WordDocument.Create();
        try {
            if (options.IncludeMetadata) CopyMetadata(source.Reader.Metadata(), target);
            AddWarning(options, "VisualPagesNotEditable", "Document",
                "PDF pages are embedded as images. Text, links, and forms are not editable Word objects.",
                PdfCore.PdfConversionWarningSeverity.Warning);
            for (int index = 0; index < pages.Count; index++) {
                token.ThrowIfCancellationRequested();
                PdfCore.PdfPageRenderResult page = pages[index];
                byte[] bytes = page.Bytes ?? throw new InvalidOperationException("A selected PDF page could not be rendered.");
                OfficeDrawing drawing = source.Render.Drawing(page.PageNumber);
                double width = drawing.Width, height = drawing.Height;
                // Word's supported physical page size is at most 22 inches in each dimension.
                if (width <= 0 || height <= 0 || width > 1584 || height > 1584)
                    throw new NotSupportedException("The selected PDF page exceeds Word's supported physical page size.");
                WordSection section = index == 0 ? target.Sections[0] : target.AddSection(WordSectionBreakType.NextPage);
                section.PageSettings.Orientation = width > height ? OfficePageOrientation.Landscape : OfficePageOrientation.Portrait;
                section.PageSettings.Width = (uint)Math.Round(width * 20D);
                section.PageSettings.Height = (uint)Math.Round(height * 20D);
                section.Margins.Left = section.Margins.Right = 0;
                section.Margins.Top = section.Margins.Bottom = 0;
                section.Margins.HeaderDistance = section.Margins.FooterDistance = 0;
                WordParagraph paragraph = target.AddParagraph();
                paragraph.LineSpacingBefore = paragraph.LineSpacingAfter = 0;
                paragraph.LineSpacingPoints = 1;
                using var stream = new MemoryStream(bytes, writable: false);
                WordImage image = paragraph.InsertImage(stream, "page-" + page.PageNumber + ".png", width * 96D / 72D,
                    height * 96D / 72D, WordImageTextWrapping.InFrontOfText, "PDF page " + page.PageNumber);
                image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
                image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
                image.HorizontalPositionOffset = image.VerticalPositionOffset = 0;
                foreach (string diagnostic in page.Diagnostics)
                    AddWarning(options, "VisualPageRendering", "Page " + page.PageNumber, diagnostic, PdfCore.PdfConversionWarningSeverity.Warning);
            }
            token.ThrowIfCancellationRequested();
            return new PdfWordConversionResult(target, options.Report);
        } catch { target.Dispose(); throw; }
    }
}
