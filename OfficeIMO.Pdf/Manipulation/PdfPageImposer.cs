using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Vector page imposition over the existing PDF page-overlay engine.</summary>
internal static class PdfPageImposer {
    internal static PdfImpositionResult ImposeNUp(byte[] pdf, PdfNUpOptions options, PdfPageSelection? selection, PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(options, nameof(options));
        _ = options.Validate();
        PdfDocumentInfo info = PdfInspector.Inspect(pdf, readOptions);
        if (info.PageCount == 0) throw new ArgumentException("PDF does not contain readable pages.", nameof(pdf));
        if (selection is null && info.PageCount > options.MaxSourcePages) throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPages, options.MaxSourcePages, info.PageCount);
        int[] pageNumbers = selection?.ToPageNumbers(info.PageCount, nameof(selection)) ?? Enumerable.Range(1, info.PageCount).ToArray();
        if (pageNumbers.Length == 0) throw new ArgumentException("N-up requires at least one selected source page.", nameof(selection));
        if (pageNumbers.Length > options.MaxSourcePages) throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPages, options.MaxSourcePages, pageNumbers.Length);
        return ImposePlan(pdf, options, pageNumbers.Select(static page => (int?)page).ToArray(), readOptions, info);
    }

    internal static PdfImpositionResult ImposeBooklet(byte[] pdf, PdfBookletOptions options, PdfPageSelection? selection, PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(options, nameof(options));
        PdfNUpOptions layout = options.ToNUpOptions();
        _ = layout.Validate();
        PdfDocumentInfo info = PdfInspector.Inspect(pdf, readOptions);
        if (info.PageCount == 0) throw new ArgumentException("PDF does not contain readable pages.", nameof(pdf));
        if (selection is null && info.PageCount > layout.MaxSourcePages) throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPages, layout.MaxSourcePages, info.PageCount);
        int[] pageNumbers = selection?.ToPageNumbers(info.PageCount, nameof(selection)) ?? Enumerable.Range(1, info.PageCount).ToArray();
        if (pageNumbers.Length == 0) throw new ArgumentException("Booklet requires at least one selected source page.", nameof(selection));
        if (pageNumbers.Length > layout.MaxSourcePages) throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPages, layout.MaxSourcePages, pageNumbers.Length);
        int physicalSheets = checked((int)((pageNumbers.LongLength + 3L) / 4L));
        if (physicalSheets > options.MaxPhysicalSheets) throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPages, options.MaxPhysicalSheets, physicalSheets);
        int padded = checked(physicalSheets * 4);
        var plan = new int?[padded];
        for (int sheet = 0; sheet < physicalSheets; sheet++) {
            int frontLeft = padded - 2 * sheet;
            int frontRight = 1 + 2 * sheet;
            int backLeft = 2 + 2 * sheet;
            int backRight = padded - 1 - 2 * sheet;
            int offset = sheet * 4;
            plan[offset] = PageAt(frontLeft);
            plan[offset + 1] = PageAt(frontRight);
            plan[offset + 2] = PageAt(backLeft);
            plan[offset + 3] = PageAt(backRight);
            if (options.RightToLeft) {
                (plan[offset], plan[offset + 1]) = (plan[offset + 1], plan[offset]);
                (plan[offset + 2], plan[offset + 3]) = (plan[offset + 3], plan[offset + 2]);
            }
        }
        return ImposePlan(pdf, layout, plan, readOptions, info);

        int? PageAt(int oneBasedPosition) => oneBasedPosition <= pageNumbers.Length ? pageNumbers[oneBasedPosition - 1] : null;
    }

    private static PdfImpositionResult ImposePlan(byte[] pdf, PdfNUpOptions options, int?[] pageNumbers, PdfLoadOptions? readOptions, PdfDocumentInfo info) {
        (double cellWidth, double cellHeight) = options.Validate();
        PdfImpositionSourceFeatureLoss sourceFeatureLoss = PdfImpositionSourceFeatureLoss.None;
        if (info.HasAnnotations) sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.Annotations;
        if (info.HasForms) sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.Forms;
        if (info.HasTaggedContent) sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.StructureTags;
        if (info.HasSignatures) sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.Signatures;
        if (info.HasOptionalContent) {
            throw new NotSupportedException("Layered source PDFs cannot be imposed because the sheet output cannot preserve optional-content visibility.");
        }
        byte[] sourcePdf = pdf;
        PdfLoadOptions? sourceReadOptions = readOptions;
        int removedSignatureCount = 0;
        if (info.HasSignatures) {
            if (options.SignaturePolicy != PdfImpositionSignaturePolicy.CreateUnsignedDerivative) {
                throw new NotSupportedException("Signed source PDFs require SignaturePolicy = CreateUnsignedDerivative before imposition.");
            }
            PdfUnsignedDerivativeResult derivative = PdfRedactionApplier.CreateUnsignedDerivative(pdf, readOptions);
            sourcePdf = derivative.Pdf;
            sourceReadOptions = null;
            removedSignatureCount = derivative.RemovedSignatureCount;
            info = PdfInspector.Inspect(sourcePdf);
        }
        if (!options.AllowSourceFeatureLoss && (info.HasAnnotations || info.HasForms || info.HasTaggedContent)) {
            throw new NotSupportedException("Imposed output cannot retain source annotations, forms, or structure tags. Set AllowSourceFeatureLoss to accept this loss.");
        }
        int cellsPerSheet = checked(options.Columns * options.Rows);
        int sheetCount = checked((int)((pageNumbers.LongLength + cellsPerSheet - 1L) / cellsPerSheet));
        if (sheetCount > options.MaxSheets) throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPages, options.MaxSheets, sheetCount);

        var blank = PdfDocument.Create(new PdfOptions { PageSize = options.SheetSize });
        for (int index = 0; index < sheetCount; index++) {
            if (index != 0) blank.PageBreak();
            blank.Canvas(static canvas => {
                OfficeShape marker = OfficeShape.Rectangle(1D, 1D);
                marker.FillColor = OfficeColor.White;
                marker.StrokeWidth = 0D;
                canvas.Shape(marker, 0D, 0D);
            });
        }
        byte[] blankSheets = blank.ToBytes();
        var overlays = new List<PdfPageOverlayOptions>(pageNumbers.Length);
        var placements = new List<PdfImpositionPlacement>(pageNumbers.Length);
        for (int index = 0; index < pageNumbers.Length; index++) {
            if (pageNumbers[index] is not int sourcePage) continue;
            int sheet = index / cellsPerSheet + 1;
            int slot = index % cellsPerSheet;
            int row = slot / options.Columns;
            int column = slot % options.Columns;
            double x = options.Margin + column * (cellWidth + options.HorizontalGutter);
            double y = options.SheetSize.Height - options.Margin - (row + 1) * cellHeight - row * options.VerticalGutter;
            var cell = new PdfPageRectangle(x, y, x + cellWidth, y + cellHeight);
            placements.Add(new PdfImpositionPlacement(sourcePage, sheet, row, column, cell));
            overlays.Add(new PdfPageOverlayOptions {
                SourcePageNumber = sourcePage,
                SourceReadOptions = sourceReadOptions,
                TargetPages = PdfPageSelector.Parse(sheet.ToString(CultureInfo.InvariantCulture)),
                X = x, Y = y, Width = cellWidth, Height = cellHeight,
                Fit = PdfPageOverlayFit.Contain
            });
        }
        byte[] output = PdfStamper.StampPages(blankSheets, sourcePdf, overlays);
        PdfDocumentInfo resultInfo = PdfInspector.Inspect(output);
        if (resultInfo.PageCount != sheetCount) throw new InvalidOperationException("Imposed output page count did not match its placement plan.");
        return new PdfImpositionResult(output, placements, removedSignatureCount, sourceFeatureLoss);
    }
}
