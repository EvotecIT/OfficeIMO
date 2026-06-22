using OfficeIMO.Pdf.Filters;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

/// <summary>
/// Applies rectangle-based redactions by removing matched text objects and annotations, then painting redaction marks.
/// </summary>
public static class PdfRedactionApplier {
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);
    private static readonly Regex TextObjectRegex = new Regex(@"\bBT\b[\s\S]*?\bET\b", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex HexStringRegex = new Regex(@"<([0-9A-Fa-f\s]+)>", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex LiteralStringRegex = new Regex(@"\((?:\\.|[^\\()])*\)", RegexOptions.Compiled, RegexTimeout);

    /// <summary>
    /// Applies rectangle-based redactions to a PDF byte array and returns rewritten PDF bytes.
    /// </summary>
    public static byte[] Apply(
        byte[] pdf,
        IEnumerable<PdfRedactionArea> areas,
        PdfRedactionApplyOptions? applyOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(areas, nameof(areas));
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);

        PdfRedactionArea[] areaArray = areas.ToArray();
        if (areaArray.Length == 0) {
            throw new ArgumentException("At least one redaction area is required.", nameof(areas));
        }

        PdfRedactionApplyOptions effectiveOptions = applyOptions ?? new PdfRedactionApplyOptions();
        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(pdf, areaArray, layoutOptions, readOptions);
        if (!plan.Preflight.CanReadLogicalObjects) {
            throw new InvalidOperationException("PDF redaction cannot be applied because logical content cannot be read. " + string.Join(" ", plan.Preflight.GetCapabilityDiagnostics(PdfPreflightCapability.ReadLogicalObjects)));
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        int catalogObjectNumber = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalogObjectNumber == 0) {
            throw new ArgumentException("PDF does not contain a readable catalog.", nameof(pdf));
        }

        PdfReadDocument document = PdfReadDocument.Load(pdf, readOptions);
        ValidateRedactionAreas(areaArray, document.Pages.Count);
        RedactionMutation mutation = ApplyToObjects(objects, document, plan, areaArray, effectiveOptions);
        if (!mutation.HasChanges) {
            return pdf.ToArray();
        }

        return RewriteAllObjects(objects, catalogObjectNumber, document.Metadata, pdf);
    }

    /// <summary>
    /// Applies rectangle-based redactions from the current position of a readable stream.
    /// </summary>
    public static byte[] Apply(
        Stream stream,
        IEnumerable<PdfRedactionArea> areas,
        PdfRedactionApplyOptions? applyOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        return Apply(ReadStream(stream, nameof(stream)), areas, applyOptions, layoutOptions, readOptions);
    }

    /// <summary>
    /// Applies rectangle-based redactions to a PDF and writes the rewritten bytes to a stream.
    /// </summary>
    public static void Apply(
        byte[] pdf,
        Stream outputStream,
        IEnumerable<PdfRedactionArea> areas,
        PdfRedactionApplyOptions? applyOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        WriteOutput(outputStream, Apply(pdf, areas, applyOptions, layoutOptions, readOptions));
    }

    /// <summary>
    /// Applies rectangle-based redactions to a PDF file and writes a new PDF file.
    /// </summary>
    public static void Apply(
        string inputPath,
        string outputPath,
        IEnumerable<PdfRedactionArea> areas,
        PdfRedactionApplyOptions? applyOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        byte[] redacted = Apply(File.ReadAllBytes(inputPath), areas, applyOptions, layoutOptions, readOptions);
        WriteOutput(fullOutputPath, redacted);
    }

    /// <summary>
    /// Applies rectangle-based redactions to a PDF file and returns rewritten PDF bytes.
    /// </summary>
    public static byte[] ApplyToBytes(
        string inputPath,
        IEnumerable<PdfRedactionArea> areas,
        PdfRedactionApplyOptions? applyOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return Apply(File.ReadAllBytes(inputPath), areas, applyOptions, layoutOptions, readOptions);
    }

    private static RedactionMutation ApplyToObjects(
        Dictionary<int, PdfIndirectObject> objects,
        PdfReadDocument document,
        PdfRedactionPlan plan,
        PdfRedactionArea[] areas,
        PdfRedactionApplyOptions options) {
        var matchesByPage = plan.Matches
            .GroupBy(match => match.PageNumber)
            .ToDictionary(group => group.Key, group => group.ToArray());
        var areasByPage = areas
            .GroupBy(area => area.PageNumber)
            .ToDictionary(group => group.Key, group => group.ToArray());
        bool changed = false;
        int nextObjectNumber = objects.Keys.Count == 0 ? 1 : objects.Keys.Max() + 1;
        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            int pageNumber = pageIndex + 1;
            PdfReadPage readPage = document.Pages[pageIndex];
            if (!objects.TryGetValue(readPage.ObjectNumber, out PdfIndirectObject? pageObject) ||
                pageObject.Value is not PdfDictionary pageDictionary) {
                continue;
            }

            matchesByPage.TryGetValue(pageNumber, out PdfRedactionMatch[]? pageMatches);
            areasByPage.TryGetValue(pageNumber, out PdfRedactionArea[]? pageAreas);
            if ((pageMatches is null || pageMatches.Length == 0) &&
                (pageAreas is null || pageAreas.Length == 0)) {
                continue;
            }

            bool pageChanged = RemoveMatchedTextObjects(objects, pageDictionary, pageMatches ?? Array.Empty<PdfRedactionMatch>(), ref nextObjectNumber);
            pageChanged = RemoveMatchedAnnotations(objects, pageDictionary, pageMatches ?? Array.Empty<PdfRedactionMatch>()) || pageChanged;

            PdfRedactionArea[] paintAreas = SelectPaintAreas(pageAreas ?? Array.Empty<PdfRedactionArea>(), pageMatches ?? Array.Empty<PdfRedactionMatch>(), options);
            if (paintAreas.Length > 0) {
                int contentObjectNumber = nextObjectNumber++;
                objects[contentObjectNumber] = new PdfIndirectObject(contentObjectNumber, 0, BuildRedactionContentStream(paintAreas, options.FillColor));
                AppendPageContent(objects, pageDictionary, contentObjectNumber);
                pageChanged = true;
            }

            changed = pageChanged || changed;
        }

        return new RedactionMutation(changed);
    }

    private static void ValidateRedactionAreas(PdfRedactionArea[] areas, int pageCount) {
        for (int i = 0; i < areas.Length; i++) {
            if (areas[i].PageNumber > pageCount) {
                throw new ArgumentOutOfRangeException(nameof(areas), "Redaction area page number " + areas[i].PageNumber.ToString(CultureInfo.InvariantCulture) + " is outside the document page count " + pageCount.ToString(CultureInfo.InvariantCulture) + ".");
            }
        }
    }

    private static bool RemoveMatchedTextObjects(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        IReadOnlyList<PdfRedactionMatch> matches,
        ref int nextObjectNumber) {
        TextMatchTarget[] textMatches = matches
            .Where(match => match.Kind == PdfRedactionMatchKind.TextBlock && !string.IsNullOrWhiteSpace(match.Text))
            .Select(match => new TextMatchTarget(NormalizeText(match.Text!), match.X, match.Y, match.Width, match.Height))
            .Where(target => target.Text.Length > 0)
            .ToArray();
        if (textMatches.Length == 0 ||
            !pageDictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject)) {
            return false;
        }

        bool changed = false;
        foreach (PdfReference reference in EnumerateContentReferences(objects, contentsObject)) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                indirect.Value is not PdfStream stream ||
                stream.DecodingFailed) {
                continue;
            }

            byte[] contentBytes = StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
            string content = PdfEncoding.Latin1GetString(contentBytes);
            string scrubbed = ScrubTextObjects(content, textMatches);
            if (string.Equals(content, scrubbed, StringComparison.Ordinal)) {
                continue;
            }

            PdfStream scrubbedStream = new PdfStream(CleanStreamDictionary(stream.Dictionary), PdfEncoding.Latin1GetBytes(scrubbed));
            if (CountPageContentReferences(objects, reference) <= 1) {
                objects[reference.ObjectNumber] = new PdfIndirectObject(reference.ObjectNumber, reference.Generation, scrubbedStream);
            } else {
                int clonedObjectNumber = nextObjectNumber++;
                objects[clonedObjectNumber] = new PdfIndirectObject(clonedObjectNumber, 0, scrubbedStream);
                ReplacePageContentReference(objects, pageDictionary, reference, new PdfReference(clonedObjectNumber, 0));
            }

            changed = true;
        }

        changed = ScrubFormXObjects(objects, pageDictionary, textMatches, new HashSet<int>(), ref nextObjectNumber) || changed;
        return changed;
    }

    private static string ScrubTextObjects(string content, TextMatchTarget[] normalizedTextMatches, bool requireIntersection = true) {
        return TextObjectRegex.Replace(content, match => {
            string shownText = NormalizeText(ExtractTextFromTextObject(match.Value));
            bool hasShownText = shownText.Length > 0;
            string contextualTextObject = BuildContextualTextObject(content, match);

            for (int i = 0; i < normalizedTextMatches.Length; i++) {
                bool intersects = TextObjectIntersects(contextualTextObject, normalizedTextMatches[i]);
                if (hasShownText &&
                    ContainsOrdinal(shownText, normalizedTextMatches[i].Text) &&
                    (!requireIntersection || intersects)) {
                    return string.Empty;
                }

                if (intersects) {
                    return string.Empty;
                }
            }

            return match.Value;
        });
    }

    private static bool ScrubFormXObjects(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary ownerDictionary,
        TextMatchTarget[] textMatches,
        HashSet<int> visitedFormObjects,
        ref int nextObjectNumber) {
        if (textMatches.Length == 0 ||
            !TryCloneXObjectResourceDictionary(objects, ownerDictionary, out PdfDictionary xObjects)) {
            return false;
        }

        bool changed = false;
        foreach (KeyValuePair<string, PdfObject> entry in xObjects.Items.ToArray()) {
            PdfObject xObject = entry.Value;
            if (xObject is not PdfReference reference ||
                !visitedFormObjects.Add(reference.ObjectNumber) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                indirect.Value is not PdfStream stream ||
                stream.DecodingFailed ||
                !IsFormXObject(stream.Dictionary)) {
                continue;
            }

            byte[] contentBytes = StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
            string content = PdfEncoding.Latin1GetString(contentBytes);
            string scrubbed = ScrubTextObjects(content, textMatches, requireIntersection: false);
            PdfDictionary clonedFormDictionary = CloneDictionary(stream.Dictionary);
            bool nestedChanged = ScrubFormXObjects(objects, clonedFormDictionary, textMatches, visitedFormObjects, ref nextObjectNumber);
            if (string.Equals(content, scrubbed, StringComparison.Ordinal) && !nestedChanged) {
                continue;
            }

            PdfDictionary streamDictionary = CleanStreamDictionary(clonedFormDictionary);
            if (CountXObjectReferences(objects, reference.ObjectNumber) <= 1) {
                objects[reference.ObjectNumber] = new PdfIndirectObject(reference.ObjectNumber, reference.Generation, new PdfStream(streamDictionary, PdfEncoding.Latin1GetBytes(scrubbed)));
                xObjects.Items[entry.Key] = reference;
            } else {
                int clonedObjectNumber = nextObjectNumber++;
                objects[clonedObjectNumber] = new PdfIndirectObject(clonedObjectNumber, 0, new PdfStream(streamDictionary, PdfEncoding.Latin1GetBytes(scrubbed)));
                xObjects.Items[entry.Key] = new PdfReference(clonedObjectNumber, 0);
            }

            changed = true;
        }

        return changed;
    }

    private static int CountXObjectReferences(Dictionary<int, PdfIndirectObject> objects, int objectNumber) {
        int count = 0;
        foreach (PdfIndirectObject indirect in objects.Values) {
            PdfDictionary? dictionary = indirect.Value switch {
                PdfDictionary value => value,
                PdfStream stream => stream.Dictionary,
                _ => null
            };
            if (dictionary is null ||
                ResolveDictionary(objects, dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject) ? resourcesObject : null) is not PdfDictionary resources ||
                ResolveDictionary(objects, resources.Items.TryGetValue("XObject", out PdfObject? xObjectObject) ? xObjectObject : null) is not PdfDictionary xObjects) {
                continue;
            }

            count += xObjects.Items.Values.Count(value => value is PdfReference reference && reference.ObjectNumber == objectNumber);
        }

        return count;
    }

    private static int CountPageContentReferences(Dictionary<int, PdfIndirectObject> objects, PdfReference contentReference) {
        int count = 0;
        foreach (PdfIndirectObject indirect in objects.Values) {
            if (indirect.Value is not PdfDictionary dictionary ||
                !dictionary.Items.ContainsKey("Contents")) {
                continue;
            }

            count += EnumerateContentReferences(objects, dictionary.Items["Contents"])
                .Count(reference => reference.ObjectNumber == contentReference.ObjectNumber && reference.Generation == contentReference.Generation);
        }

        return count;
    }

    private static bool TryCloneXObjectResourceDictionary(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary ownerDictionary,
        out PdfDictionary xObjects) {
        xObjects = null!;
        if (!TryGetInheritedValue(objects, ownerDictionary, "Resources", out PdfObject? resourcesObject) ||
            ResolveDictionary(objects, resourcesObject) is not PdfDictionary resources ||
            ResolveDictionary(objects, resources.Items.TryGetValue("XObject", out PdfObject? xObjectObject) ? xObjectObject : null) is not PdfDictionary sourceXObjects) {
            return false;
        }

        PdfDictionary clonedResources = CloneDictionary(resources);
        xObjects = CloneDictionary(sourceXObjects);
        clonedResources.Items["XObject"] = xObjects;
        ownerDictionary.Items["Resources"] = clonedResources;
        return true;
    }

    private static PdfDictionary CloneDictionary(PdfDictionary source) {
        var clone = new PdfDictionary();
        foreach (KeyValuePair<string, PdfObject> item in source.Items) {
            clone.Items[item.Key] = item.Value;
        }

        return clone;
    }

    private static bool TryGetInheritedValue(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        string key,
        out PdfObject? value) {
        var visitedParents = new HashSet<int>();
        PdfDictionary? current = dictionary;
        while (current is not null) {
            if (current.Items.TryGetValue(key, out value)) {
                return true;
            }

            if (!current.Items.TryGetValue("Parent", out PdfObject? parentObject)) {
                break;
            }

            if (parentObject is PdfReference parentReference) {
                if (!visitedParents.Add(parentReference.ObjectNumber) ||
                    !PdfObjectLookup.TryGet(objects, parentReference, out PdfIndirectObject? parentIndirect) ||
                    parentIndirect.Value is not PdfDictionary parentDictionary) {
                    break;
                }

                current = parentDictionary;
                continue;
            }

            current = ResolveDictionary(objects, parentObject);
        }

        value = null;
        return false;
    }

    private static void ReplacePageContentReference(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        PdfReference oldReference,
        PdfReference newReference) {
        if (!pageDictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject)) {
            return;
        }

        if (ReferenceEquals(oldReference, contentsObject) || PdfReferenceEquals(contentsObject, oldReference)) {
            pageDictionary.Items["Contents"] = newReference;
            return;
        }

        PdfArray? sourceArray = PdfObjectLookup.Resolve(objects, contentsObject) as PdfArray;
        if (sourceArray is null) {
            return;
        }

        var clonedArray = new PdfArray();
        foreach (PdfObject item in sourceArray.Items) {
            clonedArray.Items.Add(PdfReferenceEquals(item, oldReference) ? newReference : item);
        }

        pageDictionary.Items["Contents"] = clonedArray;
    }

    private static bool PdfReferenceEquals(PdfObject value, PdfReference reference) {
        return value is PdfReference other &&
            other.ObjectNumber == reference.ObjectNumber &&
            other.Generation == reference.Generation;
    }

    private static bool IsFormXObject(PdfDictionary dictionary) =>
        dictionary.Get<PdfName>("Type")?.Name == "XObject" &&
        dictionary.Get<PdfName>("Subtype")?.Name == "Form";

    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) =>
        PdfObjectLookup.Resolve(objects, value) as PdfDictionary;

    private static string BuildContextualTextObject(string content, Match textObjectMatch) {
        string prefix = content.Substring(0, textObjectMatch.Index);
        return TextObjectRegex.Replace(prefix, string.Empty) + textObjectMatch.Value;
    }

    private static bool TextObjectIntersects(string textObject, TextMatchTarget target) {
        List<PdfTextSpan> spans = TextContentParser.Parse(
            textObject,
            static (_, bytes) => PdfWinAnsiEncoding.Decode(bytes),
            static (_, bytes) => Math.Max(0D, bytes.Length * 500D));
        if (spans.Count == 0) {
            return false;
        }

        for (int i = 0; i < spans.Count; i++) {
            PdfTextSpan span = spans[i];
            double width = Math.Max(span.Advance, Math.Max(1, span.Text.Length) * Math.Max(1D, span.FontSize) * 0.5D);
            double height = Math.Max(1D, span.FontSize) * 1.5D;
            double x = Math.Min(span.X, span.X + width);
            double y = span.Y - Math.Max(1D, span.FontSize);
            if (Intersects(target.X, target.Y, target.Width, target.Height, x, y, Math.Abs(width), height)) {
                return true;
            }
        }

        return false;
    }

    private static bool Intersects(double ax, double ay, double aw, double ah, double bx, double by, double bw, double bh) {
        return ax < bx + bw &&
            ax + aw > bx &&
            ay < by + bh &&
            ay + ah > by;
    }

    private static string ExtractTextFromTextObject(string textObject) {
        var builder = new StringBuilder();
        foreach (Match match in HexStringRegex.Matches(textObject)) {
            builder.Append(DecodeHexString(match.Groups[1].Value));
        }

        foreach (Match match in LiteralStringRegex.Matches(textObject)) {
            builder.Append(DecodeLiteralString(match.Value));
        }

        return builder.ToString();
    }

    private static string DecodeHexString(string value) {
        string hex = RemoveWhitespace(value);
        if (hex.Length == 0) {
            return string.Empty;
        }

        if ((hex.Length & 1) == 1) {
            hex += "0";
        }

        var bytes = new byte[hex.Length / 2];
        for (int i = 0; i < bytes.Length; i++) {
            bytes[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
        }

        return PdfWinAnsiEncoding.Decode(bytes);
    }

    private static string DecodeLiteralString(string value) {
        if (value.Length < 2) {
            return string.Empty;
        }

        string body = value.Substring(1, value.Length - 2);
        var builder = new StringBuilder(body.Length);
        for (int i = 0; i < body.Length; i++) {
            char ch = body[i];
            if (ch != '\\' || i + 1 >= body.Length) {
                builder.Append(ch);
                continue;
            }

            char next = body[++i];
            builder.Append(next switch {
                'n' => '\n',
                'r' => '\r',
                't' => '\t',
                'b' => '\b',
                'f' => '\f',
                _ => next
            });
        }

        return builder.ToString();
    }

    private static bool RemoveMatchedAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        IReadOnlyList<PdfRedactionMatch> matches) {
        AnnotationMatchTarget[] annotationMatches = matches
            .Where(match => match.Kind == PdfRedactionMatchKind.Annotation)
            .Select(static match => new AnnotationMatchTarget(match.ObjectNumber, match.X, match.Y, match.Width, match.Height))
            .ToArray();
        if (annotationMatches.Length == 0 ||
            !pageDictionary.Items.TryGetValue("Annots", out PdfObject? annotsObject) ||
            PdfObjectLookup.Resolve(objects, annotsObject) is not PdfArray annotations) {
            return false;
        }

        var removedObjectNumbers = new HashSet<int>();
        bool changed = false;
        for (int i = annotations.Items.Count - 1; i >= 0; i--) {
            PdfObject item = annotations.Items[i];
            int? objectNumber = item is PdfReference reference ? reference.ObjectNumber : null;
            PdfDictionary? annotation = PdfObjectLookup.Resolve(objects, item) as PdfDictionary;
            if (annotation is null ||
                !MatchesAnnotationRedaction(objects, annotation, objectNumber, annotationMatches)) {
                continue;
            }

            annotations.Items.RemoveAt(i);
            if (annotation.Items.TryGetValue("Popup", out PdfObject? popupObject) &&
                popupObject is PdfReference popupReference) {
                removedObjectNumbers.Add(popupReference.ObjectNumber);
            }

            if (objectNumber.HasValue) {
                removedObjectNumbers.Add(objectNumber.Value);
            }

            changed = true;
        }

        if (removedObjectNumbers.Count > 0) {
            changed = RemoveLinkedPopupAnnotations(objects, annotations, removedObjectNumbers) || changed;
            foreach (int objectNumber in removedObjectNumbers) {
                objects.Remove(objectNumber);
            }
        }

        if (!changed) {
            return false;
        }

        if (annotations.Items.Count == 0) {
            pageDictionary.Items.Remove("Annots");
        }

        return true;
    }

    private static bool MatchesAnnotationRedaction(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary annotation,
        int? objectNumber,
        AnnotationMatchTarget[] annotationMatches) {
        for (int i = 0; i < annotationMatches.Length; i++) {
            AnnotationMatchTarget target = annotationMatches[i];
            if (objectNumber.HasValue &&
                target.ObjectNumber.HasValue &&
                objectNumber.Value == target.ObjectNumber.Value) {
                return true;
            }

            if (!objectNumber.HasValue &&
                TryReadRectangle(objects, annotation, "Rect", out double x, out double y, out double width, out double height) &&
                Intersects(target.X, target.Y, target.Width, target.Height, x, y, width, height)) {
                return true;
            }
        }

        return false;
    }

    private static bool RemoveLinkedPopupAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfArray annotations,
        HashSet<int> removedObjectNumbers) {
        bool changed = false;
        for (int i = annotations.Items.Count - 1; i >= 0; i--) {
            if (annotations.Items[i] is not PdfReference reference ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                indirect.Value is not PdfDictionary annotation) {
                continue;
            }

            if (removedObjectNumbers.Contains(reference.ObjectNumber) ||
                IsPopupForRemovedAnnotation(annotation, removedObjectNumbers)) {
                annotations.Items.RemoveAt(i);
                removedObjectNumbers.Add(reference.ObjectNumber);
                changed = true;
            }
        }

        return changed;
    }

    private static bool IsPopupForRemovedAnnotation(PdfDictionary annotation, HashSet<int> removedObjectNumbers) {
        return annotation.Get<PdfName>("Subtype")?.Name == "Popup" &&
            annotation.Items.TryGetValue("Parent", out PdfObject? parentObject) &&
            parentObject is PdfReference parentReference &&
            removedObjectNumbers.Contains(parentReference.ObjectNumber);
    }

    private static bool TryReadRectangle(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        string key,
        out double x,
        out double y,
        out double width,
        out double height) {
        x = 0D;
        y = 0D;
        width = 0D;
        height = 0D;
        if (!dictionary.Items.TryGetValue(key, out PdfObject? rectObject) ||
            PdfObjectLookup.Resolve(objects, rectObject) is not PdfArray rect ||
            rect.Items.Count < 4 ||
            PdfObjectLookup.Resolve(objects, rect.Items[0]) is not PdfNumber x1 ||
            PdfObjectLookup.Resolve(objects, rect.Items[1]) is not PdfNumber y1 ||
            PdfObjectLookup.Resolve(objects, rect.Items[2]) is not PdfNumber x2 ||
            PdfObjectLookup.Resolve(objects, rect.Items[3]) is not PdfNumber y2) {
            return false;
        }

        x = Math.Min(x1.Value, x2.Value);
        y = Math.Min(y1.Value, y2.Value);
        width = Math.Abs(x2.Value - x1.Value);
        height = Math.Abs(y2.Value - y1.Value);
        return width > 0D && height > 0D;
    }

    private static PdfRedactionArea[] SelectPaintAreas(PdfRedactionArea[] areas, PdfRedactionMatch[] matches, PdfRedactionApplyOptions options) {
        if (options.PaintUnmatchedAreas) {
            return areas;
        }

        return areas
            .Where(area => matches.Any(match => ReferenceEquals(match.Area, area) || SameArea(match.Area, area)))
            .ToArray();
    }

    private static bool SameArea(PdfRedactionArea left, PdfRedactionArea right) {
        return left.PageNumber == right.PageNumber &&
            AreClose(left.X, right.X) &&
            AreClose(left.Y, right.Y) &&
            AreClose(left.Width, right.Width) &&
            AreClose(left.Height, right.Height);
    }

    private static bool AreClose(double left, double right) {
        return Math.Abs(left - right) < 0.0001D;
    }

    private static PdfStream BuildRedactionContentStream(PdfRedactionArea[] areas, PdfColor fillColor) {
        var builder = new StringBuilder();
        var content = new ContentStreamBuilder(builder)
            .SaveState()
            .FillColor(fillColor);
        for (int i = 0; i < areas.Length; i++) {
            content.Rectangle(areas[i].X, areas[i].Y, areas[i].Width, areas[i].Height)
                .FillPath();
        }

        content.RestoreState();
        return new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes(builder.ToString()));
    }

    private static void AppendPageContent(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page, int contentObjectNumber) {
        var newReference = new PdfReference(contentObjectNumber, 0);
        if (!page.Items.TryGetValue("Contents", out PdfObject? contents)) {
            page.Items["Contents"] = newReference;
            return;
        }

        if (contents is PdfArray contentsArray) {
            contentsArray.Items.Add(newReference);
            return;
        }

        var array = new PdfArray();
        foreach (PdfObject item in EnumerateContentObjects(objects, contents)) {
            array.Items.Add(item);
        }

        array.Items.Add(newReference);
        page.Items["Contents"] = array;
    }

    private static IEnumerable<PdfReference> EnumerateContentReferences(Dictionary<int, PdfIndirectObject> objects, PdfObject contents) {
        foreach (PdfObject item in EnumerateContentObjects(objects, contents)) {
            if (item is PdfReference reference) {
                yield return reference;
            }
        }
    }

    private static IEnumerable<PdfObject> EnumerateContentObjects(Dictionary<int, PdfIndirectObject> objects, PdfObject contents) {
        if (contents is PdfArray directArray) {
            foreach (PdfObject item in directArray.Items) {
                yield return item;
            }

            yield break;
        }

        if (contents is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
            indirect.Value is PdfArray referencedArray) {
            foreach (PdfObject item in referencedArray.Items) {
                yield return item;
            }

            yield break;
        }

        yield return contents;
    }

    private static PdfDictionary CleanStreamDictionary(PdfDictionary source) {
        var dictionary = new PdfDictionary();
        foreach (KeyValuePair<string, PdfObject> entry in source.Items) {
            if (string.Equals(entry.Key, "Length", StringComparison.Ordinal) ||
                string.Equals(entry.Key, "Filter", StringComparison.Ordinal) ||
                string.Equals(entry.Key, "DecodeParms", StringComparison.Ordinal)) {
                continue;
            }

            dictionary.Items[entry.Key] = entry.Value;
        }

        return dictionary;
    }

    private static byte[] RewriteAllObjects(Dictionary<int, PdfIndirectObject> objects, int catalogObjectNumber, PdfMetadata metadata, byte[] sourcePdf) {
        int[] sourceIds = objects.Keys.OrderBy(id => id).ToArray();
        var numberMap = new Dictionary<int, int>(sourceIds.Length);
        for (int i = 0; i < sourceIds.Length; i++) {
            numberMap[sourceIds[i]] = i + 1;
        }

        var context = new PdfPageExtractor.SerializationContext(numberMap, pagesObjectId: 0, new Dictionary<int, Dictionary<string, PdfObject>>(), objects);
        var rewritten = new List<byte[]>(sourceIds.Length + 1);
        foreach (int sourceId in sourceIds) {
            rewritten.Add(PdfPageExtractor.WrapObject(numberMap[sourceId], PdfPageExtractor.SerializeObject(objects[sourceId].Value, context)));
        }

        int infoId = rewritten.Count + 1;
        rewritten.Add(PdfPageExtractor.WrapObject(infoId, PdfEncoding.Latin1GetBytes(PdfPageExtractor.BuildInfoDictionary(metadata))));

        PdfFileVersion fileVersion = PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(sourcePdf));
        return PdfPageExtractor.Assemble(rewritten, numberMap[catalogObjectNumber], infoId, fileVersion);
    }

    private static int FindCatalogObjectNumber(Dictionary<int, PdfIndirectObject> objects, string? trailerRaw) {
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects, trailerRaw);
        if (catalog is null) {
            return 0;
        }

        foreach (KeyValuePair<int, PdfIndirectObject> entry in objects) {
            if (ReferenceEquals(entry.Value.Value, catalog)) {
                return entry.Key;
            }
        }

        return 0;
    }

    private static string NormalizeText(string value) {
        return string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }

    private static bool ContainsOrdinal(string value, string match) {
#if NET472 || NETSTANDARD2_0
        return value.IndexOf(match, StringComparison.Ordinal) >= 0;
#else
        return value.Contains(match, StringComparison.Ordinal);
#endif
    }

    private static string RemoveWhitespace(string value) {
        var builder = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++) {
            if (!char.IsWhiteSpace(value[i])) {
                builder.Append(value[i]);
            }
        }

        return builder.ToString();
    }

    private static byte[] ReadStream(Stream stream, string paramName) {
        Guard.NotNull(stream, paramName);
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", paramName);
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static void WriteOutput(Stream outputStream, byte[] bytes) {
        Guard.NotNull(outputStream, nameof(outputStream));
        if (!outputStream.CanWrite) {
            throw new ArgumentException("Stream must be writable.", nameof(outputStream));
        }

        outputStream.Write(bytes, 0, bytes.Length);
    }

    private static void WriteOutput(string outputPath, byte[] bytes) {
        string fullPath = ValidateOutputPath(outputPath);
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllBytes(fullPath, bytes);
    }

    private static string ValidateOutputPath(string outputPath) {
        Guard.NotNull(outputPath, nameof(outputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path cannot be empty or whitespace.", nameof(outputPath));
        }

        string fullPath;
        try {
            fullPath = Path.GetFullPath(outputPath);
        } catch (Exception ex) {
            throw new ArgumentException("Output path is invalid.", nameof(outputPath), ex);
        }

        if (Directory.Exists(fullPath) && (File.GetAttributes(fullPath) & FileAttributes.Directory) == FileAttributes.Directory) {
            throw new ArgumentException("Output path refers to a directory; a file path is required.", nameof(outputPath));
        }

        string fileName = Path.GetFileName(fullPath);
        if (string.IsNullOrEmpty(fileName)) {
            throw new ArgumentException("Output path must include a file name.", nameof(outputPath));
        }

        if (fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {
            throw new ArgumentException("Output path contains invalid file name characters.", nameof(outputPath));
        }

        return fullPath;
    }

    private readonly struct RedactionMutation {
        public RedactionMutation(bool hasChanges) {
            HasChanges = hasChanges;
        }

        public bool HasChanges { get; }
    }

    private readonly struct TextMatchTarget {
        public TextMatchTarget(string text, double x, double y, double width, double height) {
            Text = text;
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        public string Text { get; }
        public double X { get; }
        public double Y { get; }
        public double Width { get; }
        public double Height { get; }
    }

    private readonly struct AnnotationMatchTarget {
        public AnnotationMatchTarget(int? objectNumber, double x, double y, double width, double height) {
            ObjectNumber = objectNumber;
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        public int? ObjectNumber { get; }
        public double X { get; }
        public double Y { get; }
        public double Width { get; }
        public double Height { get; }
    }
}
