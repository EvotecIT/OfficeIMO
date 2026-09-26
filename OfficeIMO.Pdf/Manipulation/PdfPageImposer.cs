using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Vector page imposition over the existing PDF page-overlay engine.</summary>
internal static class PdfPageImposer {
    private static readonly string[] UnpreservedCatalogFeatureNames = {
        "Collection", "Extensions", "Requirements", "Legal", "Threads", "PieceInfo", "SpiderInfo"
    };

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
        int[] selectedPages = pageNumbers.OfType<int>().Distinct().ToArray();
        PdfReadDocument rawSource = PdfReadDocument.Open(pdf, readOptions);
        bool selectedAnnotations = selectedPages.Any(page => HasRawAnnotations(rawSource.Pages[page - 1], rawSource.Objects));
        bool selectedPageFeatures = selectedPages.Any(page => HasRawPageFeatures(rawSource.Pages[page - 1], rawSource.Objects));
        // XFA and fields without a placed widget have no page ownership to filter by selection.
        bool documentLevelForms = info.HasForms && (info.HasAcroFormXfa || info.FormFields.Count == 0 ||
            info.FormFields.Any(static field => field.Widgets.Count == 0 ||
                field.Widgets.Any(static widget => !widget.PageNumber.HasValue)));
        bool selectedForms = documentLevelForms || selectedPages.Any(page => info.Pages[page - 1].FormWidgets.Count > 0);
        PdfImpositionSourceFeatureLoss sourceFeatureLoss = PdfImpositionSourceFeatureLoss.None;
        if (selectedAnnotations) sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.Annotations;
        if (selectedForms) sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.Forms;
        if (info.HasTaggedContent) sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.StructureTags;
        if (info.HasSignatures) sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.Signatures;
        if (info.HasEmbeddedFiles) sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.EmbeddedFiles;
        if (info.HasOutputIntents) sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.OutputIntents;
        if (info.HasOutlines) sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.Outlines;
        if (info.HasPageLabels) sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.PageLabels;
        if (info.HasNamedDestinations) sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.NamedDestinations;
        if (info.HasCatalogViewSettings || info.HasOpenActions || info.HasViewerPreferences ||
            info.HasCatalogNameTrees || info.HasCatalogUri || info.HasCatalogActions ||
            HasRawCatalogActiveContent(rawSource) ||
            HasRawCatalogUnpreservedFeatures(rawSource) ||
            !string.IsNullOrWhiteSpace(info.CatalogLanguage)) {
            sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.CatalogFeatures;
        }
        if (info.Metadata.HasContent || HasRawCatalogXmpMetadata(rawSource) ||
            HasUnpreservedInfoMetadata(rawSource)) {
            sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.DocumentMetadata;
        }
        if (selectedPageFeatures) {
            sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.PageFeatures;
        }
        if (info.Security.HasEncryption) {
            // Preserve the source extraction-permission error before reporting a derivative-loss policy error.
            rawSource.DemandContentExtraction("page imposition");
            sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.Encryption;
        }
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
            // The signature policy already accounts for removed signature fields and widgets.
            // Retain only interactive features that remain on the unsigned derivative.
            PdfReadDocument derivativeSource = PdfReadDocument.Open(sourcePdf);
            selectedAnnotations = selectedPages.Any(page =>
                HasRawAnnotations(derivativeSource.Pages[page - 1], derivativeSource.Objects));
            bool hasResidualFields = info.FormFields.Count > 0 || HasRawFormFields(derivativeSource);
            documentLevelForms = info.HasAcroFormXfa || hasResidualFields &&
                (info.FormFields.Count == 0 || info.FormFields.Any(static field => field.Widgets.Count == 0 ||
                    field.Widgets.Any(static widget => !widget.PageNumber.HasValue)));
            selectedForms = documentLevelForms || selectedPages.Any(page => info.Pages[page - 1].FormWidgets.Count > 0);
            sourceFeatureLoss &= ~(PdfImpositionSourceFeatureLoss.Annotations | PdfImpositionSourceFeatureLoss.Forms);
            if (selectedAnnotations) sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.Annotations;
            if (selectedForms) sourceFeatureLoss |= PdfImpositionSourceFeatureLoss.Forms;
        }
        if (!options.AllowSourceFeatureLoss &&
            (sourceFeatureLoss & ~PdfImpositionSourceFeatureLoss.Signatures) != PdfImpositionSourceFeatureLoss.None) {
            throw new NotSupportedException("Imposed output cannot retain one or more source features: " +
                (sourceFeatureLoss & ~PdfImpositionSourceFeatureLoss.Signatures) +
                ". Set AllowSourceFeatureLoss to accept this loss.");
        }
        int cellsPerSheet = checked(options.Columns * options.Rows);
        int sheetCount = checked((int)((pageNumbers.LongLength + cellsPerSheet - 1L) / cellsPerSheet));
        if (sheetCount > options.MaxSheets) throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPages, options.MaxSheets, sheetCount);

        var blank = PdfDocument.Create(new PdfOptions { PageSize = options.SheetSize,
            FileVersion = PdfFileAssembler.ParseHeaderVersionOrDefault(info.EffectiveVersion),
            MaxGeneratedOutputBytes = options.MaxOutputBytes });
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
        byte[] output = PdfStamper.StampPages(blankSheets, sourcePdf, overlays, options.MaxOutputBytes);
        PdfDocumentInfo resultInfo = PdfInspector.Inspect(output);
        if (resultInfo.PageCount != sheetCount) throw new InvalidOperationException("Imposed output page count did not match its placement plan.");
        return new PdfImpositionResult(output, placements, removedSignatureCount, sourceFeatureLoss);
    }

    private static bool HasRawAnnotations(PdfReadPage page, Dictionary<int, PdfIndirectObject> objects) {
        if (!page.PageDictionary.Items.TryGetValue("Annots", out PdfObject? value)) return false;
        PdfObject? resolved = PdfObjectLookup.ResolveChain(objects, value);
        return resolved is not PdfNull && (resolved is not PdfArray annotations || annotations.Items.Count > 0);
    }

    private static bool HasRawPageFeatures(PdfReadPage page, Dictionary<int, PdfIndirectObject> objects) {
        foreach (string key in new[] { "AA", "Metadata", "PieceInfo", "Tabs", "Dur", "Trans", "PZ", "TrimBox", "BleedBox", "ArtBox", "Thumb", "VP", "SeparationInfo", "BoxColorInfo", "PresSteps" }) {
            if (page.PageDictionary.Items.TryGetValue(key, out PdfObject? value) &&
                PdfObjectLookup.ResolveChain(objects, value) is not PdfNull) return true;
        }
        return false;
    }

    private static bool HasUnpreservedInfoMetadata(PdfReadDocument source) {
        if (!PdfSyntax.TryGetTrailerReference(source.TrailerRaw, "Info", source.ReadOptions.Limits, out PdfReference reference) ||
            !PdfObjectLookup.TryGet(source.Objects, reference, out PdfIndirectObject info) ||
            info.Value is not PdfDictionary dictionary) return false;
        foreach (KeyValuePair<string, PdfObject> entry in dictionary.Items) {
            PdfObject? value = PdfObjectLookup.ResolveChain(source.Objects, entry.Value);
            if (value is PdfNull) continue;
            if (entry.Key == "Producer" && value is PdfStringObj { Value: "OfficeIMO.Pdf" }) continue;
            return true;
        }
        return false;
    }

    private static bool HasRawCatalogXmpMetadata(PdfReadDocument source) =>
        source.CatalogDictionary?.Items.TryGetValue("Metadata", out PdfObject? value) == true &&
        PdfObjectLookup.ResolveChain(source.Objects, value) is not PdfNull;

    private static bool HasRawCatalogActiveContent(PdfReadDocument source) =>
        source.CatalogDictionary is PdfDictionary catalog &&
        PdfActiveContentPolicy.MarkerNames.Any(name => catalog.Items.TryGetValue(name, out PdfObject? value) &&
            PdfObjectLookup.ResolveChain(source.Objects, value) is not null and not PdfNull);

    private static bool HasRawCatalogUnpreservedFeatures(PdfReadDocument source) =>
        source.CatalogDictionary is PdfDictionary catalog &&
        (UnpreservedCatalogFeatureNames
            .Any(name => catalog.Items.TryGetValue(name, out PdfObject? value) &&
                HasCatalogFeature(name, PdfObjectLookup.ResolveChain(source.Objects, value))) ||
         catalog.Items.TryGetValue("NeedsRendering", out PdfObject? needsRendering) &&
         PdfObjectLookup.ResolveChain(source.Objects, needsRendering) is not null and not PdfNull and not PdfBoolean { Value: false });

    private static bool HasCatalogFeature(string name, PdfObject? value) => value switch {
        null or PdfNull => false,
        PdfArray array => array.Items.Count > 0,
        PdfDictionary dictionary => name == "Collection" || dictionary.Items.Count > 0,
        _ => true
    };

    private static bool HasRawFormFields(PdfReadDocument source) {
        PdfDictionary? catalog = source.CatalogDictionary;
        if (catalog == null ||
            PdfObjectLookup.ResolveChain(source.Objects, catalog.Items.TryGetValue("AcroForm", out PdfObject? form) ? form : null)
                is not PdfDictionary acroForm) return false;
        PdfObject? fields = PdfObjectLookup.ResolveChain(source.Objects,
            acroForm.Items.TryGetValue("Fields", out PdfObject? value) ? value : null);
        return fields is PdfArray array ? array.Items.Count > 0 : fields is not null and not PdfNull;
    }
}
