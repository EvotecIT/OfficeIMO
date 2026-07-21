using OfficeIMO.Pdf.Filters;
using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
    /// <summary>Imports one source PDF page as a Form XObject above selected target pages.</summary>
    public static byte[] OverlayPage(byte[] targetPdf, byte[] sourcePdf, PdfPageOverlayOptions? options = null, PdfReadOptions? targetReadOptions = null) {
        return StampPageCore(targetPdf, sourcePdf, (options ?? new PdfPageOverlayOptions()).Clone(behindContent: false), targetReadOptions);
    }

    /// <summary>Imports one source PDF page as a Form XObject below selected target pages.</summary>
    public static byte[] UnderlayPage(byte[] targetPdf, byte[] sourcePdf, PdfPageOverlayOptions? options = null, PdfReadOptions? targetReadOptions = null) {
        return StampPageCore(targetPdf, sourcePdf, (options ?? new PdfPageOverlayOptions()).Clone(behindContent: true), targetReadOptions);
    }

    /// <summary>Imports one source PDF page as a Form XObject using the requested content order.</summary>
    public static byte[] StampPage(byte[] targetPdf, byte[] sourcePdf, PdfPageOverlayOptions? options = null, PdfReadOptions? targetReadOptions = null) {
        return StampPageCore(targetPdf, sourcePdf, options?.Clone() ?? new PdfPageOverlayOptions(), targetReadOptions);
    }

    /// <summary>Imports a source PDF page onto target pages read from streams.</summary>
    public static byte[] StampPage(Stream targetPdf, Stream sourcePdf, PdfPageOverlayOptions? options = null) {
        return StampPage(ReadStream(targetPdf, nameof(targetPdf)), ReadStream(sourcePdf, nameof(sourcePdf)), options);
    }

    /// <summary>Imports one source PDF page above selected target pages read from streams.</summary>
    public static byte[] OverlayPage(Stream targetPdf, Stream sourcePdf, PdfPageOverlayOptions? options = null) {
        return OverlayPage(ReadStream(targetPdf, nameof(targetPdf)), ReadStream(sourcePdf, nameof(sourcePdf)), options);
    }

    /// <summary>Imports one source PDF page below selected target pages read from streams.</summary>
    public static byte[] UnderlayPage(Stream targetPdf, Stream sourcePdf, PdfPageOverlayOptions? options = null) {
        return UnderlayPage(ReadStream(targetPdf, nameof(targetPdf)), ReadStream(sourcePdf, nameof(sourcePdf)), options);
    }

    private static byte[] StampPageCore(byte[] targetPdf, byte[] sourcePdf, PdfPageOverlayOptions options, PdfReadOptions? targetReadOptions = null) {
        return StampPageSetCore(
            targetPdf,
            new[] { new PageStampRequest(sourcePdf, options) },
            targetReadOptions);
    }

    private static byte[] StampPageSetCore(
        byte[] targetPdf,
        IReadOnlyList<PageStampRequest> requests,
        PdfReadOptions? targetReadOptions = null) {
        Guard.NotNull(targetPdf, nameof(targetPdf));
        Guard.NotNull(requests, nameof(requests));
        if (requests.Count == 0) {
            return targetPdf;
        }

        _ = PdfMutationPlanner.RequireFullRewrite(targetPdf, PdfMutationOperation.ModifyPageContent, targetReadOptions);

        var (targetObjects, targetTrailer) = PdfSyntax.ParseObjects(targetPdf, targetReadOptions);
        PdfReadDocument target = PdfReadDocument.Open(targetPdf, targetReadOptions);
        if (target.Pages.Count == 0) throw new ArgumentException("Target PDF does not contain any pages.", nameof(targetPdf));
        int nextObjectNumber = targetObjects.Count == 0 ? 1 : targetObjects.Keys.Max() + 1;
        int[] pageObjectNumbers = target.Pages.Select(page => page.ObjectNumber).ToArray();
        var overrides = new Dictionary<int, Dictionary<string, PdfObject>>();
        var reservedFormResourceNames = new HashSet<string>(StringComparer.Ordinal);
        var reservedGraphicsStateResourceNames = new HashSet<string>(StringComparer.Ordinal);
        var stampedPageNumbers = new HashSet<int>();
        PdfFileVersion outputVersion = PdfPageExtractor.GetSourceFileVersion(targetPdf);

        for (int requestIndex = 0; requestIndex < requests.Count; requestIndex++) {
            PageStampRequest request = requests[requestIndex];
            byte[] sourcePdf = request.SourcePdf;
            PdfPageOverlayOptions options = request.Options;
            Guard.NotNull(sourcePdf, nameof(requests));
            Guard.NotNull(options, nameof(requests));
            PdfReadOptions? sourceReadOptions = options.SourceReadOptions;
            _ = PdfMutationPlanner.RequireFullRewrite(sourcePdf, PdfMutationOperation.ExtractPages, sourceReadOptions);

            var (sourceObjects, _) = PdfSyntax.ParseObjects(sourcePdf, sourceReadOptions);
            PdfReadDocument source = PdfReadDocument.Open(sourcePdf, sourceReadOptions);
            if (options.SourcePageNumber > source.Pages.Count) {
                throw new ArgumentOutOfRangeException(nameof(requests), options.SourcePageNumber, "Source page number exceeds the source PDF page count.");
            }

            PdfReadPage sourcePage = source.Pages[options.SourcePageNumber - 1];
            if (!sourceObjects.TryGetValue(sourcePage.ObjectNumber, out PdfIndirectObject? sourcePageObject) || sourcePageObject.Value is not PdfDictionary sourcePageDictionary) {
                throw new InvalidOperationException("Source PDF page dictionary was not found.");
            }

            PdfObject? sourceResources = GetInheritedPageValue(sourceObjects, sourcePageDictionary, "Resources");
            PdfObject? sourceGroup = sourcePageDictionary.Items.TryGetValue("Group", out PdfObject? groupValue)
                ? groupValue
                : null;
            var sourceCollector = new PdfPageExtractor.ObjectCollector(sourceObjects);
            sourceCollector.CollectObjectGraph(sourceResources);
            sourceCollector.CollectObjectGraph(sourceGroup);
            var importedObjectNumbers = new Dictionary<int, int>();
            foreach (int sourceObjectNumber in sourceCollector.ObjectIds) importedObjectNumbers[sourceObjectNumber] = nextObjectNumber++;
            foreach (int sourceObjectNumber in sourceCollector.ObjectIds) {
                PdfIndirectObject sourceObject = sourceObjects[sourceObjectNumber];
                int importedNumber = importedObjectNumbers[sourceObjectNumber];
                targetObjects[importedNumber] = new PdfIndirectObject(importedNumber, 0, CloneImportedObject(sourceObject.Value, importedObjectNumbers));
            }

            (double sourceWidth, double sourceHeight, Matrix2D normalization) = sourcePage.GetImportGeometry();
            byte[] formContent = BuildImportedPageContent(sourceObjects, sourcePageDictionary, sourceWidth, sourceHeight, normalization);
            var formDictionary = new PdfDictionary();
            formDictionary.Items["Type"] = new PdfName("XObject");
            formDictionary.Items["Subtype"] = new PdfName("Form");
            formDictionary.Items["FormType"] = new PdfNumber(1D);
            formDictionary.Items["BBox"] = NumberArray(0D, 0D, sourceWidth, sourceHeight);
            formDictionary.Items["Resources"] = sourceResources == null
                ? new PdfDictionary()
                : CloneImportedObject(sourceResources, importedObjectNumbers);
            if (sourceGroup != null) formDictionary.Items["Group"] = CloneImportedObject(sourceGroup, importedObjectNumbers);
            int formObjectNumber = nextObjectNumber++;
            targetObjects[formObjectNumber] = new PdfIndirectObject(formObjectNumber, 0, new PdfStream(formDictionary, formContent));

            int graphicsStateObjectNumber = 0;
            if (options.Opacity < 1D) {
                var graphicsState = new PdfDictionary();
                graphicsState.Items["Type"] = new PdfName("ExtGState");
                graphicsState.Items["ca"] = new PdfNumber(options.Opacity);
                graphicsState.Items["CA"] = new PdfNumber(options.Opacity);
                graphicsStateObjectNumber = nextObjectNumber++;
                targetObjects[graphicsStateObjectNumber] = new PdfIndirectObject(graphicsStateObjectNumber, 0, graphicsState);
            }

            string formResourceName = GetAvailableXObjectResourceName(targetObjects, pageObjectNumbers, reservedFormResourceNames);
            reservedFormResourceNames.Add(formResourceName);
            string? graphicsStateResourceName = graphicsStateObjectNumber == 0
                ? null
                : GetAvailableGraphicsStateResourceName(targetObjects, pageObjectNumbers, reservedGraphicsStateResourceNames);
            if (graphicsStateResourceName != null) reservedGraphicsStateResourceNames.Add(graphicsStateResourceName);

            IReadOnlyList<int> selectedPages = options.TargetPages?.Resolve(target.Pages.Count) ?? Enumerable.Range(1, target.Pages.Count).ToArray();
            for (int selectedIndex = 0; selectedIndex < selectedPages.Count; selectedIndex++) {
                int pageNumber = selectedPages[selectedIndex];
                if (!stampedPageNumbers.Add(pageNumber)) {
                    throw new InvalidOperationException("A page can only receive one generated canvas stamp in a single rewrite.");
                }

                PdfReadPage targetPage = target.Pages[pageNumber - 1];
                (double targetWidth, double targetHeight, Matrix2D targetUserToVisual) = targetPage.GetImportGeometry();
                Matrix2D targetVisualToUser = Invert(targetUserToVisual);
                PdfStream stamp = BuildImportedPageStampStream(formResourceName, graphicsStateResourceName, sourceWidth, sourceHeight, targetWidth, targetHeight, targetVisualToUser, options);
                int stampObjectNumber = nextObjectNumber++;
                targetObjects[stampObjectNumber] = new PdfIndirectObject(stampObjectNumber, 0, stamp);
                overrides[targetPage.ObjectNumber] = BuildImportedPageOverrides(
                    targetObjects,
                    targetPage.ObjectNumber,
                    formResourceName,
                    formObjectNumber,
                    graphicsStateResourceName,
                    graphicsStateObjectNumber,
                    stampObjectNumber,
                    options.BehindContent);
            }

            PdfFileVersion sourceVersion = PdfPageExtractor.GetSourceFileVersion(sourcePdf);
            if (sourceVersion > outputVersion) outputVersion = sourceVersion;
        }

        return PdfPageExtractor.ExtractPages(
            targetObjects,
            target.Metadata,
            pageObjectNumbers,
            overrides,
            catalogState: PdfPageExtractor.ExtractCatalogRewriteState(targetObjects, targetTrailer),
            fileVersion: outputVersion);
    }

    private sealed class PageStampRequest {
        internal PageStampRequest(byte[] sourcePdf, PdfPageOverlayOptions options) {
            SourcePdf = sourcePdf;
            Options = options;
        }

        internal byte[] SourcePdf { get; }

        internal PdfPageOverlayOptions Options { get; }
    }

    private static byte[] BuildImportedPageContent(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfDictionary page,
        double width,
        double height,
        Matrix2D normalization) {
        var builder = new StringBuilder();
        var content = new ContentStreamBuilder(builder);
        content.SaveState()
            .Rectangle(0D, 0D, width, height).ClipPath().EndPath()
            .TransformMatrix(normalization.A, normalization.B, normalization.C, normalization.D, normalization.E, normalization.F);
        foreach (PdfStream stream in GetPageContentStreams(sourceObjects, page)) {
            byte[] decoded = StreamDecoder.Decode(stream.Dictionary, stream.Data, sourceObjects);
            builder.Append(PdfEncoding.Latin1GetString(decoded)).Append('\n');
        }
        content.RestoreState();
        return PdfEncoding.Latin1GetBytes(builder.ToString());
    }

    private static PdfStream BuildImportedPageStampStream(
        string formResourceName,
        string? graphicsStateResourceName,
        double sourceWidth,
        double sourceHeight,
        double targetPageWidth,
        double targetPageHeight,
        Matrix2D targetVisualToUser,
        PdfPageOverlayOptions options) {
        double frameWidth = options.Width ?? targetPageWidth;
        double frameHeight = options.Height ?? targetPageHeight;
        double frameX = options.X ?? AlignHorizontal(0D, targetPageWidth, frameWidth, options.HorizontalAlignment);
        double frameY = options.Y ?? AlignVertical(0D, targetPageHeight, frameHeight, options.VerticalAlignment);
        double drawWidth = sourceWidth;
        double drawHeight = sourceHeight;
        switch (options.Fit) {
            case PdfPageOverlayFit.None:
                break;
            case PdfPageOverlayFit.Contain: {
                double scale = Math.Min(frameWidth / sourceWidth, frameHeight / sourceHeight);
                drawWidth *= scale;
                drawHeight *= scale;
                break;
            }
            case PdfPageOverlayFit.Cover: {
                double scale = Math.Max(frameWidth / sourceWidth, frameHeight / sourceHeight);
                drawWidth *= scale;
                drawHeight *= scale;
                break;
            }
            case PdfPageOverlayFit.Stretch:
                drawWidth = frameWidth;
                drawHeight = frameHeight;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(options), options.Fit, "Unsupported imported-page fit mode.");
        }

        double drawX = AlignHorizontal(frameX, frameWidth, drawWidth, options.HorizontalAlignment);
        double drawY = AlignVertical(frameY, frameHeight, drawHeight, options.VerticalAlignment);
        var builder = new StringBuilder();
        var content = new ContentStreamBuilder(builder).SaveState();
        if (!IsIdentity(targetVisualToUser)) {
            content.TransformMatrix(targetVisualToUser.A, targetVisualToUser.B, targetVisualToUser.C, targetVisualToUser.D, targetVisualToUser.E, targetVisualToUser.F);
        }
        if (options.Fit == PdfPageOverlayFit.Cover) content.Rectangle(frameX, frameY, frameWidth, frameHeight).ClipPath().EndPath();
        if (graphicsStateResourceName != null) content.GraphicsState(graphicsStateResourceName);
        content.TransformMatrix(drawWidth / sourceWidth, 0D, 0D, drawHeight / sourceHeight, drawX, drawY)
            .XObject(formResourceName)
            .RestoreState();
        return new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes(builder.ToString()));
    }

    private static Matrix2D Invert(Matrix2D matrix) {
        double determinant = (matrix.A * matrix.D) - (matrix.B * matrix.C);
        if (Math.Abs(determinant) < 0.000000000001D) throw new InvalidOperationException("Target page geometry transform cannot be inverted.");
        return new Matrix2D(
            matrix.D / determinant,
            -matrix.B / determinant,
            -matrix.C / determinant,
            matrix.A / determinant,
            ((matrix.C * matrix.F) - (matrix.D * matrix.E)) / determinant,
            ((matrix.B * matrix.E) - (matrix.A * matrix.F)) / determinant);
    }

    private static bool IsIdentity(Matrix2D matrix) =>
        Math.Abs(matrix.A - 1D) < 0.000000000001D && Math.Abs(matrix.B) < 0.000000000001D &&
        Math.Abs(matrix.C) < 0.000000000001D && Math.Abs(matrix.D - 1D) < 0.000000000001D &&
        Math.Abs(matrix.E) < 0.000000000001D && Math.Abs(matrix.F) < 0.000000000001D;

    private static Dictionary<string, PdfObject> BuildImportedPageOverrides(
        Dictionary<int, PdfIndirectObject> objects,
        int pageObjectNumber,
        string formResourceName,
        int formObjectNumber,
        string? graphicsStateResourceName,
        int graphicsStateObjectNumber,
        int stampObjectNumber,
        bool behindContent) {
        PdfDictionary page = (PdfDictionary)objects[pageObjectNumber].Value;
        PdfArray contents = BuildContentsArray(objects, page.Items.TryGetValue("Contents", out PdfObject? existing) ? existing : null, stampObjectNumber, behindContent);
        PdfDictionary resources = CloneDictionary(ResolveDictionary(objects, GetInheritedPageValue(objects, page, "Resources")));
        PdfDictionary xObjects = CloneDictionary(ResolveDictionary(objects, resources.Items.TryGetValue("XObject", out PdfObject? xObject) ? xObject : null));
        xObjects.Items[formResourceName] = new PdfReference(formObjectNumber, 0);
        resources.Items["XObject"] = xObjects;
        if (graphicsStateResourceName != null) {
            PdfDictionary states = CloneDictionary(ResolveDictionary(objects, resources.Items.TryGetValue("ExtGState", out PdfObject? state) ? state : null));
            states.Items[graphicsStateResourceName] = new PdfReference(graphicsStateObjectNumber, 0);
            resources.Items["ExtGState"] = states;
        }
        return new Dictionary<string, PdfObject>(StringComparer.Ordinal) { ["Contents"] = contents, ["Resources"] = resources };
    }

    private static List<PdfStream> GetPageContentStreams(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page) {
        var streams = new List<PdfStream>();
        if (!page.Items.TryGetValue("Contents", out PdfObject? contents)) return streams;
        AppendStream(contents);
        return streams;

        void AppendStream(PdfObject value) {
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, value);
            if (resolved is PdfStream stream) {
                streams.Add(stream);
            } else if (resolved is PdfArray array) {
                foreach (PdfObject item in array.Items) AppendStream(item);
            }
        }
    }

    private static PdfObject CloneImportedObject(PdfObject value, IReadOnlyDictionary<int, int> numberMap) {
        switch (value) {
            case PdfReference reference:
                if (!numberMap.TryGetValue(reference.ObjectNumber, out int mapped)) throw new InvalidOperationException("Imported PDF resource references an object outside the collected resource graph.");
                return new PdfReference(mapped, 0);
            case PdfNumber number: return new PdfNumber(number.Value);
            case PdfBoolean boolean: return new PdfBoolean(boolean.Value);
            case PdfName name: return new PdfName(name.Name);
            case PdfStringObj text: return new PdfStringObj(text.RawBytes, text.UseTextStringEncoding);
            case PdfNull: return PdfNull.Instance;
            case PdfArray array: {
                var clone = new PdfArray();
                foreach (PdfObject item in array.Items) clone.Items.Add(CloneImportedObject(item, numberMap));
                return clone;
            }
            case PdfDictionary dictionary: {
                var clone = new PdfDictionary();
                foreach (KeyValuePair<string, PdfObject> item in dictionary.Items) clone.Items[item.Key] = CloneImportedObject(item.Value, numberMap);
                return clone;
            }
            case PdfStream stream: {
                var dictionary = new PdfDictionary();
                foreach (KeyValuePair<string, PdfObject> item in stream.Dictionary.Items) {
                    if (!string.Equals(item.Key, "Length", StringComparison.Ordinal)) {
                        dictionary.Items[item.Key] = CloneImportedObject(item.Value, numberMap);
                    }
                }
                return new PdfStream(dictionary, (byte[])stream.Data.Clone(), stream.DecodingFailed, stream.DecodingError);
            }
            default:
                throw new NotSupportedException("Unsupported imported PDF resource object type " + value.GetType().Name + ".");
        }
    }

    private static PdfArray NumberArray(params double[] values) {
        var array = new PdfArray();
        foreach (double value in values) array.Items.Add(new PdfNumber(value));
        return array;
    }

    private static double AlignHorizontal(double x, double available, double actual, PdfAlign align) => align switch {
        PdfAlign.Center => x + (available - actual) / 2D,
        PdfAlign.Right => x + available - actual,
        _ => x
    };

    private static double AlignVertical(double y, double available, double actual, PdfVerticalAlign align) => align switch {
        PdfVerticalAlign.Top => y + available - actual,
        PdfVerticalAlign.Middle => y + (available - actual) / 2D,
        _ => y
    };

    private static string GetAvailableGraphicsStateResourceName(
        Dictionary<int, PdfIndirectObject> objects,
        int[] pageObjectNumbers,
        HashSet<string>? additionallyUsed = null) {
        var used = new HashSet<string>(StringComparer.Ordinal);
        foreach (int pageObjectNumber in pageObjectNumbers) {
            if (!objects.TryGetValue(pageObjectNumber, out PdfIndirectObject? indirect) || indirect.Value is not PdfDictionary page) continue;
            PdfDictionary? resources = ResolveDictionary(objects, GetInheritedPageValue(objects, page, "Resources"));
            PdfDictionary? states = ResolveDictionary(objects, resources?.Items.TryGetValue("ExtGState", out PdfObject? state) == true ? state : null);
            if (states != null) foreach (string name in states.Items.Keys) used.Add(name);
        }
        for (int i = 1; i < 1000; i++) {
            string candidate = "OIMOStampGS" + i.ToString(CultureInfo.InvariantCulture);
            if (!used.Contains(candidate) && (additionallyUsed is null || !additionallyUsed.Contains(candidate))) return candidate;
        }
        throw new InvalidOperationException("Unable to find an available PDF graphics-state resource name for the imported page.");
    }
}
