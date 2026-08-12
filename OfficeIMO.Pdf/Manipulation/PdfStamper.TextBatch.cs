using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
    internal static byte[] StampTextBatch(
        byte[] pdf,
        IReadOnlyList<TextStampRequest> requests,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(requests, nameof(requests));
        if (requests.Count == 0) return pdf;

        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyPageContent, readOptions);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        if (document.Pages.Count == 0) throw new ArgumentException("PDF does not contain any pages.", nameof(pdf));

        for (int index = 0; index < requests.Count; index++) {
            TextStampRequest request = requests[index];
            if (request.PageNumber < 1 || request.PageNumber > document.Pages.Count) {
                throw new ArgumentOutOfRangeException(nameof(requests), request.PageNumber, "Text stamp page number exceeds the PDF page count.");
            }
            if (request.Text.Length == 0) throw new ArgumentException("Text stamp requests cannot contain empty text.", nameof(requests));
            if (request.FontSize <= 0D || !IsFinite(request.FontSize) || !IsFinite(request.X) || !IsFinite(request.Y) || !IsFinite(request.RotationDegrees)) {
                throw new ArgumentOutOfRangeException(nameof(requests), "Text stamp geometry must contain finite coordinates and a positive font size.");
            }
        }

        int[] pageObjectNumbers = document.Pages.Select(static page => page.ObjectNumber).ToArray();
        PdfStandardFont[] fonts = requests.Select(static request => request.Font).Distinct().ToArray();
        string[] resourceNames = GetAvailableBatchFontResourceNames(objects, pageObjectNumbers, fonts.Length);
        var fontResources = new Dictionary<PdfStandardFont, BatchFontResource>();
        var additionalObjects = new List<PdfPageExtractor.AdditionalObject>();
        int nextPseudoObjectNumber = -200000;
        for (int index = 0; index < fonts.Length; index++) {
            int pseudoObjectNumber = nextPseudoObjectNumber--;
            fontResources.Add(fonts[index], new BatchFontResource(resourceNames[index], pseudoObjectNumber));
            additionalObjects.Add(new PdfPageExtractor.AdditionalObject(pseudoObjectNumber, BuildFontObject(fonts[index])));
        }

        var overrides = new Dictionary<int, Dictionary<string, PdfObject>>();
        foreach (IGrouping<int, TextStampRequest> pageRequests in requests.GroupBy(static request => request.PageNumber)) {
            int pageNumber = pageRequests.Key;
            PdfReadPage page = document.Pages[pageNumber - 1];
            int stampPseudoObjectNumber = nextPseudoObjectNumber--;
            PdfStream stampStream = BuildBatchTextStampStream(pageRequests.OrderBy(static request => request.PaintOrder), fontResources);
            additionalObjects.Add(new PdfPageExtractor.AdditionalObject(stampPseudoObjectNumber, stampStream));
            overrides[page.ObjectNumber] = BuildBatchTextPageOverrides(
                objects,
                page.ObjectNumber,
                fontResources.Values,
                stampPseudoObjectNumber);
        }

        return PdfPageExtractor.ExtractPages(
            objects,
            document.UncheckedMetadata,
            pageObjectNumbers,
            overrides,
            additionalObjects,
            PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw),
            PdfPageExtractor.GetSourceFileVersion(pdf));
    }

    private static PdfStream BuildBatchTextStampStream(
        IEnumerable<TextStampRequest> requests,
        IReadOnlyDictionary<PdfStandardFont, BatchFontResource> fontResources) {
        var builder = new StringBuilder();
        foreach (TextStampRequest request in requests) {
            BatchFontResource font = fontResources[request.Font];
            double radians = request.RotationDegrees * Math.PI / 180D;
            double cos = Math.Cos(radians);
            double sin = Math.Sin(radians);
            new ContentStreamBuilder(builder)
                .SaveState()
                .FillColor(request.Color)
                .BeginText()
                .Font(font.Name, request.FontSize)
                .TextMatrix(cos, sin, -sin, cos, request.X, request.Y)
                .ShowHexText(EncodeWinAnsiHex(request.Text))
                .EndText()
                .RestoreState();
        }
        return new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes(builder.ToString()));
    }

    private static Dictionary<string, PdfObject> BuildBatchTextPageOverrides(
        Dictionary<int, PdfIndirectObject> objects,
        int pageObjectNumber,
        IEnumerable<BatchFontResource> fontResources,
        int stampPseudoObjectNumber) {
        if (!objects.TryGetValue(pageObjectNumber, out PdfIndirectObject? indirect) || indirect.Value is not PdfDictionary pageDictionary) {
            throw new InvalidOperationException("PDF page object " + pageObjectNumber.ToString(CultureInfo.InvariantCulture) + " was not found.");
        }

        PdfArray contents = BuildContentsArray(
            objects,
            pageDictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject) ? contentsObject : null,
            stampPseudoObjectNumber,
            behindContent: false);
        PdfDictionary resources = CloneDictionary(ResolveDictionary(objects, GetInheritedPageValue(objects, pageDictionary, "Resources")));
        PdfDictionary fonts = CloneDictionary(ResolveDictionary(objects, resources.Items.TryGetValue("Font", out PdfObject? fontObject) ? fontObject : null));
        foreach (BatchFontResource font in fontResources) fonts.Items[font.Name] = new PdfReference(font.PseudoObjectNumber, 0);
        resources.Items["Font"] = fonts;
        return new Dictionary<string, PdfObject>(StringComparer.Ordinal) {
            ["Contents"] = contents,
            ["Resources"] = resources
        };
    }

    private static string[] GetAvailableBatchFontResourceNames(
        Dictionary<int, PdfIndirectObject> objects,
        int[] pageObjectNumbers,
        int count) {
        var usedNames = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 0; index < pageObjectNumbers.Length; index++) {
            if (!objects.TryGetValue(pageObjectNumbers[index], out PdfIndirectObject? indirect) || indirect.Value is not PdfDictionary pageDictionary) continue;
            PdfDictionary? resources = ResolveDictionary(objects, GetInheritedPageValue(objects, pageDictionary, "Resources"));
            PdfDictionary? fonts = ResolveDictionary(objects, resources?.Items.TryGetValue("Font", out PdfObject? fontObject) == true ? fontObject : null);
            if (fonts == null) continue;
            foreach (string name in fonts.Items.Keys) usedNames.Add(name);
        }

        var result = new string[count];
        int candidateNumber = 1;
        for (int index = 0; index < count; index++) {
            string candidate;
            do {
                candidate = "OIMOEditF" + (candidateNumber++).ToString(CultureInfo.InvariantCulture);
            } while (!usedNames.Add(candidate));
            result[index] = candidate;
        }
        return result;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    internal readonly struct TextStampRequest {
        internal TextStampRequest(int pageNumber, string text, double x, double y, PdfStandardFont font, double fontSize, PdfColor color, double rotationDegrees, double paintOrder = double.MaxValue) {
            PageNumber = pageNumber;
            Text = text;
            X = x;
            Y = y;
            Font = font;
            FontSize = fontSize;
            Color = color;
            RotationDegrees = rotationDegrees;
            PaintOrder = paintOrder;
        }

        internal int PageNumber { get; }
        internal string Text { get; }
        internal double X { get; }
        internal double Y { get; }
        internal PdfStandardFont Font { get; }
        internal double FontSize { get; }
        internal PdfColor Color { get; }
        internal double RotationDegrees { get; }
        internal double PaintOrder { get; }
    }

    private readonly struct BatchFontResource {
        internal BatchFontResource(string name, int pseudoObjectNumber) {
            Name = name;
            PseudoObjectNumber = pseudoObjectNumber;
        }

        internal string Name { get; }
        internal int PseudoObjectNumber { get; }
    }
}
