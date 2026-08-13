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

        PdfMutationPlanner.RequireCatalogPreservingPageContentRewrite(pdf, readOptions);
        PdfReadLimits limits = readOptions?.Limits ?? new PdfReadLimits();
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(pdf, readOptions).Map;
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
        return PdfDocumentObjectGraphRewriter.Rewrite(pdf, readOptions, null, (rewrittenObjects, security) => {
            int nextObjectNumber = rewrittenObjects.Count == 0 ? 1 : rewrittenObjects.Keys.Max() + 1;
            var fontResources = new Dictionary<PdfStandardFont, BatchFontResource>();
            for (int index = 0; index < fonts.Length; index++) {
                int fontObjectNumber = nextObjectNumber++;
                fontResources.Add(fonts[index], new BatchFontResource(resourceNames[index], fontObjectNumber));
                rewrittenObjects[fontObjectNumber] = new PdfIndirectObject(fontObjectNumber, 0, BuildFontObject(fonts[index]));
            }

            foreach (IGrouping<int, TextStampRequest> pageRequests in requests.GroupBy(static request => request.PageNumber)) {
                TextStampRequest[] orderedPageRequests = pageRequests.OrderBy(static request => request.PaintOrder).ToArray();
                EnsureBatchTextStampStreamWithinLimit(orderedPageRequests, fontResources, limits.MaxDecodedStreamBytes);
                int pageObjectNumber = pageObjectNumbers[pageRequests.Key - 1];
                int saveStateObjectNumber = nextObjectNumber++;
                int restoreStateObjectNumber = nextObjectNumber++;
                int stampObjectNumber = nextObjectNumber++;
                rewrittenObjects[saveStateObjectNumber] = new PdfIndirectObject(saveStateObjectNumber, 0, new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes("q\n")));
                rewrittenObjects[restoreStateObjectNumber] = new PdfIndirectObject(restoreStateObjectNumber, 0, new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes("Q\n")));
                rewrittenObjects[stampObjectNumber] = new PdfIndirectObject(
                    stampObjectNumber,
                    0,
                    BuildBatchTextStampStream(orderedPageRequests, fontResources));
                Dictionary<string, PdfObject> overrides = BuildBatchTextPageOverrides(
                    rewrittenObjects,
                    pageObjectNumber,
                    fontResources.Values,
                    saveStateObjectNumber,
                    restoreStateObjectNumber,
                    stampObjectNumber);
                PdfDictionary pageDictionary = (PdfDictionary)rewrittenObjects[pageObjectNumber].Value;
                foreach (KeyValuePair<string, PdfObject> item in overrides) pageDictionary.Items[item.Key] = item.Value;
            }

            return security.InfoObjectNumber.HasValue && rewrittenObjects.ContainsKey(security.InfoObjectNumber.Value)
                ? security.InfoObjectNumber
                : null;
        });
    }

    private static PdfStream BuildBatchTextStampStream(
        IEnumerable<TextStampRequest> requests,
        IReadOnlyDictionary<PdfStandardFont, BatchFontResource> fontResources) {
        var builder = new StringBuilder();
        var encodedText = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (TextStampRequest request in requests) {
            BatchFontResource font = fontResources[request.Font];
            if (!encodedText.TryGetValue(request.Text, out string? hexText)) {
                hexText = EncodeWinAnsiHex(request.Text);
                encodedText.Add(request.Text, hexText);
            }
            AppendBatchTextStampRequest(builder, request, font, hexText);
        }
        return new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes(builder.ToString()));
    }

    private static void EnsureBatchTextStampStreamWithinLimit(
        TextStampRequest[] requests,
        IReadOnlyDictionary<PdfStandardFont, BatchFontResource> fontResources,
        int maximumDecodedStreamBytes) {
        var encodedLengths = new Dictionary<string, int>(StringComparer.Ordinal);
        long totalBytes = 0L;
        for (int index = 0; index < requests.Length; index++) {
            TextStampRequest request = requests[index];
            if (!encodedLengths.TryGetValue(request.Text, out int encodedByteLength)) {
                encodedByteLength = PdfWinAnsiEncoding.Encode(request.Text).Length;
                encodedLengths.Add(request.Text, encodedByteLength);
            }

            var fixedContent = new StringBuilder();
            AppendBatchTextStampRequest(fixedContent, request, fontResources[request.Font], string.Empty);
            totalBytes += fixedContent.Length + (encodedByteLength * 2L);
            if (totalBytes > maximumDecodedStreamBytes) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, maximumDecodedStreamBytes, totalBytes);
            }
        }
    }

    private static void AppendBatchTextStampRequest(
        StringBuilder builder,
        TextStampRequest request,
        BatchFontResource font,
        string hexText) {
        double radians = request.RotationDegrees * Math.PI / 180D;
        double cos = Math.Cos(radians);
        double sin = Math.Sin(radians);
        new ContentStreamBuilder(builder)
            .SaveState()
            .FillColor(request.Color)
            .BeginText()
            .Font(font.Name, request.FontSize)
            .TextMatrix(cos, sin, -sin, cos, request.X, request.Y)
            .ShowHexText(hexText)
            .EndText()
            .RestoreState();
    }

    private static Dictionary<string, PdfObject> BuildBatchTextPageOverrides(
        Dictionary<int, PdfIndirectObject> objects,
        int pageObjectNumber,
        IEnumerable<BatchFontResource> fontResources,
        int saveStateObjectNumber,
        int restoreStateObjectNumber,
        int stampPseudoObjectNumber) {
        if (!objects.TryGetValue(pageObjectNumber, out PdfIndirectObject? indirect) || indirect.Value is not PdfDictionary pageDictionary) {
            throw new InvalidOperationException("PDF page object " + pageObjectNumber.ToString(CultureInfo.InvariantCulture) + " was not found.");
        }

        var contents = new PdfArray();
        contents.Items.Add(new PdfReference(saveStateObjectNumber, 0));
        AppendContentEntries(
            objects,
            contents,
            pageDictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject) ? contentsObject : null);
        contents.Items.Add(new PdfReference(restoreStateObjectNumber, 0));
        contents.Items.Add(new PdfReference(stampPseudoObjectNumber, 0));
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
