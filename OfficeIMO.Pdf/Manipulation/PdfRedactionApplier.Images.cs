using OfficeIMO.Pdf.Filters;
using System.Globalization;

namespace OfficeIMO.Pdf;

public static partial class PdfRedactionApplier {
    private const double ImageRedactionTolerance = 0.01D;

    private static ImageRedactionMutation RemoveMatchedImageObjects(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        IReadOnlyList<PdfRedactionMatch> matches,
        PdfColor fillColor,
        ref int nextObjectNumber) {
        ImageRedactionTarget[] wholeImageTargets = BuildWholeImageTargets(matches);
        ImageRedactionTarget[] pixelTargets = BuildPixelImageTargets(matches);
        if ((wholeImageTargets.Length == 0 && pixelTargets.Length == 0) ||
            !pageDictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject)) {
            return ImageRedactionMutation.None;
        }

        bool changed = false;
        var removedMatches = new List<PdfRedactionMatch>();
        var removedResourceNames = new HashSet<string>(StringComparer.Ordinal);
        Dictionary<int, int> referenceCounts = CountIndirectReferenceUsage(objects);
        PdfObject currentContentsObject = contentsObject;
        bool passChanged;
        do {
            passChanged = false;
            PdfReference[] contentReferences = EnumerateContentReferences(objects, currentContentsObject).ToArray();
            foreach (PdfReference reference in contentReferences) {
                if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                    indirect.Value is not PdfStream stream ||
                    stream.DecodingFailed) {
                    continue;
                }

                byte[] contentBytes = StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
                string content = PdfEncoding.Latin1GetString(contentBytes);
                string scrubbed = RemoveImageInvocations(content, wholeImageTargets, out IReadOnlyList<ImageRedactionTarget> removedTargets);
                if (string.Equals(content, scrubbed, StringComparison.Ordinal)) {
                    continue;
                }

                PdfReference targetReference = reference;
                if (IsSharedReference(referenceCounts, reference)) {
                    targetReference = CloneIndirectObject(objects, reference, indirect, ref nextObjectNumber);
                    ReplacePageContentReference(objects, pageDictionary, currentContentsObject, reference, targetReference);
                    currentContentsObject = pageDictionary.Items.TryGetValue("Contents", out PdfObject? updatedContentsObject)
                        ? updatedContentsObject
                        : currentContentsObject;
                }

                objects[targetReference.ObjectNumber] = new PdfIndirectObject(targetReference.ObjectNumber, targetReference.Generation, new PdfStream(CleanStreamDictionary(stream.Dictionary), PdfEncoding.Latin1GetBytes(scrubbed)));
                AddRemovedImageTargets(removedTargets, removedMatches, removedResourceNames);
                changed = true;
                passChanged = true;
            }
        } while (passChanged);

        if (removedResourceNames.Count > 0) {
            RemoveUnusedPageImageResources(objects, pageDictionary, removedResourceNames);
        }

        changed = ScrubMatchedImageFormXObjects(objects, pageDictionary, currentContentsObject, wholeImageTargets, referenceCounts, removedMatches, ref nextObjectNumber) || changed;
        changed = RewriteMatchedImagePixels(objects, pageDictionary, currentContentsObject, pixelTargets, fillColor, referenceCounts, removedMatches, ref nextObjectNumber) || changed;
        return new ImageRedactionMutation(changed, removedMatches.AsReadOnly());
    }

    private static ImageRedactionTarget[] BuildWholeImageTargets(IReadOnlyList<PdfRedactionMatch> matches) {
        return matches
            .Where(match => match.Kind == PdfRedactionMatchKind.ImagePlacement &&
                !string.IsNullOrEmpty(match.ResourceName) &&
                RedactionAreaCoversMatch(match.Area, match))
            .Select(match => new ImageRedactionTarget(match, match.ResourceName!, match.X, match.Y, match.Width, match.Height))
            .ToArray();
    }

    private static ImageRedactionTarget[] BuildPixelImageTargets(IReadOnlyList<PdfRedactionMatch> matches) {
        return matches
            .Where(match => match.Kind == PdfRedactionMatchKind.ImagePlacement &&
                !string.IsNullOrEmpty(match.ResourceName) &&
                RedactionAreaIntersectsMatch(match.Area, match))
            .Select(match => new ImageRedactionTarget(match, match.ResourceName!, match.X, match.Y, match.Width, match.Height))
            .ToArray();
    }

    private static bool RedactionAreaCoversMatch(PdfRedactionArea area, PdfRedactionMatch match) {
        return area.PageNumber == match.PageNumber &&
            area.X <= match.X + ImageRedactionTolerance &&
            area.Y <= match.Y + ImageRedactionTolerance &&
            area.X + area.Width >= match.X + match.Width - ImageRedactionTolerance &&
            area.Y + area.Height >= match.Y + match.Height - ImageRedactionTolerance;
    }

    private static bool RedactionAreaIntersectsMatch(PdfRedactionArea area, PdfRedactionMatch match) {
        return area.PageNumber == match.PageNumber &&
            area.X < match.X + match.Width - ImageRedactionTolerance &&
            area.X + area.Width > match.X + ImageRedactionTolerance &&
            area.Y < match.Y + match.Height - ImageRedactionTolerance &&
            area.Y + area.Height > match.Y + ImageRedactionTolerance;
    }

    private static bool ScrubMatchedImageFormXObjects(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        PdfObject contentsObject,
        ImageRedactionTarget[] targets,
        IReadOnlyDictionary<int, int> referenceCounts,
        List<PdfRedactionMatch> removedMatches,
        ref int nextObjectNumber) {
        PdfDictionary? resources = GetInheritedDictionary(objects, pageDictionary, "Resources");
        if (resources is null ||
            !resources.Items.ContainsKey("XObject")) {
            return false;
        }

        PdfDictionary xObjects = PdfPageResourceHelper.EnsurePageXObjects(objects, pageDictionary, "redaction nested image cleanup");
        resources = ResolveDictionary(objects, pageDictionary.Items.TryGetValue("Resources", out PdfObject? pageResources) ? pageResources : null) ?? resources;
        bool changed = false;
        foreach (PdfReference reference in EnumerateContentReferences(objects, contentsObject)) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                indirect.Value is not PdfStream stream ||
                stream.DecodingFailed) {
                continue;
            }

            string content = PdfEncoding.Latin1GetString(StreamDecoder.Decode(stream.Dictionary, stream.Data, objects));
            changed = ScrubImageFormInvocations(objects, resources, xObjects, content, targets, Matrix2D.Identity, referenceCounts, new HashSet<int>(), removedMatches, ref nextObjectNumber) || changed;
        }

        return changed;
    }

    private static bool ScrubImageFormInvocations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary resources,
        PdfDictionary xObjects,
        string content,
        ImageRedactionTarget[] targets,
        Matrix2D baseTransform,
        IReadOnlyDictionary<int, int> referenceCounts,
        HashSet<int> activeForms,
        List<PdfRedactionMatch> removedMatches,
        ref int nextObjectNumber) {
        bool changed = false;
        foreach (TextContentParser.FormInvocation invocation in TextContentParser.ExtractFormInvocations(content)) {
            if (!TryGetFormXObject(objects, xObjects, invocation.Name, out PdfReference reference, out PdfStream formStream) ||
                formStream.DecodingFailed ||
                !activeForms.Add(reference.ObjectNumber) ||
                CountResourceInvocations(content, invocation.Name) != 1) {
                continue;
            }

            int activeObjectNumber = reference.ObjectNumber;
            try {
                if (IsSharedReference(referenceCounts, reference) &&
                    PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? sourceIndirect)) {
                    reference = CloneIndirectObject(objects, reference, sourceIndirect, ref nextObjectNumber);
                    xObjects.Items[invocation.Name] = reference;
                    formStream = (PdfStream)objects[reference.ObjectNumber].Value;
                    changed = true;
                }

                PdfDictionary formResources = ResolveDictionary(objects, formStream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject) ? resourcesObject : null) ?? resources;
                PdfDictionary formXObjects = ResolveDictionary(objects, formResources.Items.TryGetValue("XObject", out PdfObject? formXObjectObject) ? formXObjectObject : null) ?? new PdfDictionary();
                Matrix2D formTransform = ApplyFormMatrix(Matrix2D.Multiply(baseTransform, invocation.Transform), formStream.Dictionary);
                string formContent = PdfEncoding.Latin1GetString(StreamDecoder.Decode(formStream.Dictionary, formStream.Data, objects));
                string scrubbed = RemoveImageInvocations(formContent, targets, formTransform, out IReadOnlyList<ImageRedactionTarget> removedTargets);
                if (!string.Equals(formContent, scrubbed, StringComparison.Ordinal)) {
                    objects[reference.ObjectNumber] = new PdfIndirectObject(reference.ObjectNumber, reference.Generation, new PdfStream(CleanStreamDictionary(formStream.Dictionary), PdfEncoding.Latin1GetBytes(scrubbed)));
                    formContent = scrubbed;
                    AddRemovedImageTargets(removedTargets, removedMatches, null);
                    RemoveUnusedImageResourcesFromXObjects(objects, formXObjects, formContent, removedTargets);
                    changed = true;
                }

                changed = ScrubImageFormInvocations(objects, formResources, formXObjects, formContent, targets, formTransform, referenceCounts, activeForms, removedMatches, ref nextObjectNumber) || changed;
            } finally {
                activeForms.Remove(activeObjectNumber);
            }
        }

        return changed;
    }

    private static string RemoveImageInvocations(string content, ImageRedactionTarget[] targets, out IReadOnlyList<ImageRedactionTarget> removedTargets) {
        return RemoveImageInvocations(content, targets, Matrix2D.Identity, out removedTargets);
    }

    private static string RemoveImageInvocations(string content, ImageRedactionTarget[] targets, Matrix2D baseTransform, out IReadOnlyList<ImageRedactionTarget> removedTargets) {
        var ranges = new List<RemovalRange>();
        var removed = new List<ImageRedactionTarget>();
        Matrix2D ctm = baseTransform;
        var stack = new Stack<Matrix2D>();
        var args = new List<ImageContentOperand>(8);
        int index = 0;
        int length = content.Length;

        while (index < length) {
            SkipWhiteSpace(content, ref index);
            if (index >= length) {
                break;
            }

            char current = content[index];
            if (current == '%') {
                SkipComment(content, ref index);
                continue;
            }

            if (current == '/') {
                args.Add(ReadNameOperand(content, ref index));
                continue;
            }

            if (current == '(') {
                SkipLiteralString(content, ref index);
                continue;
            }

            if (current == '<') {
                if (index + 1 < length && content[index + 1] == '<') {
                    SkipDictionary(content, ref index);
                } else {
                    SkipHexString(content, ref index);
                }

                continue;
            }

            if (current == '[') {
                SkipArray(content, ref index);
                continue;
            }

            if (current == ']' || current == '>') {
                index++;
                continue;
            }

            if (IsNumberStart(current)) {
                args.Add(ReadNumberOperand(content, ref index));
                continue;
            }

            int operatorStart = index;
            string op = ReadOperator(content, ref index);
            int operatorEnd = index;
            if (op.Length == 0) {
                index++;
                continue;
            }

            switch (op) {
                case "q":
                    stack.Push(ctm);
                    args.Clear();
                    break;
                case "Q":
                    ctm = stack.Count > 0 ? stack.Pop() : baseTransform;
                    args.Clear();
                    break;
                case "cm":
                    if (args.Count >= 6) {
                        ctm = Matrix2D.Multiply(ctm, new Matrix2D(
                            args[args.Count - 6].Number,
                            args[args.Count - 5].Number,
                            args[args.Count - 4].Number,
                            args[args.Count - 3].Number,
                            args[args.Count - 2].Number,
                            args[args.Count - 1].Number));
                    }

                    args.Clear();
                    break;
                case "Do":
                    if (TryGetImageTarget(args, ctm, targets, out ImageRedactionTarget target, out ImageContentOperand operand)) {
                        ranges.Add(new RemovalRange(operand.Start, operatorEnd));
                        removed.Add(target);
                    }

                    args.Clear();
                    break;
                default:
                    args.Clear();
                    break;
            }
        }

        removedTargets = removed.AsReadOnly();
        return RemoveRanges(content, ranges);
    }

    private static bool TryGetImageTarget(IReadOnlyList<ImageContentOperand> args, Matrix2D ctm, ImageRedactionTarget[] targets, out ImageRedactionTarget target, out ImageContentOperand operand) {
        target = default;
        operand = default;
        if (args.Count == 0 || string.IsNullOrEmpty(args[args.Count - 1].Name)) {
            return false;
        }

        operand = args[args.Count - 1];
        GetUnitRectangleBounds(ctm, out double x, out double y, out double width, out double height);
        for (int i = 0; i < targets.Length; i++) {
            if (string.Equals(targets[i].ResourceName, operand.Name, StringComparison.Ordinal) &&
                AreCloseImageCoordinate(targets[i].X, x) &&
                AreCloseImageCoordinate(targets[i].Y, y) &&
                AreCloseImageCoordinate(targets[i].Width, width) &&
                AreCloseImageCoordinate(targets[i].Height, height)) {
                target = targets[i];
                return true;
            }
        }

        for (int i = 0; i < targets.Length; i++) {
            if (string.Equals(targets[i].ResourceName, operand.Name, StringComparison.Ordinal) &&
                RedactionAreaCoversRectangle(targets[i].Match.Area, x, y, width, height)) {
                target = targets[i];
                return true;
            }
        }

        return false;
    }

    private static bool RedactionAreaCoversRectangle(PdfRedactionArea area, double x, double y, double width, double height) =>
        area.X <= x + ImageRedactionTolerance &&
        area.Y <= y + ImageRedactionTolerance &&
        area.X + area.Width >= x + width - ImageRedactionTolerance &&
        area.Y + area.Height >= y + height - ImageRedactionTolerance;

    private static void GetUnitRectangleBounds(Matrix2D transform, out double x, out double y, out double width, out double height) {
        var p0 = transform.Transform(0D, 0D);
        var p1 = transform.Transform(1D, 0D);
        var p2 = transform.Transform(0D, 1D);
        var p3 = transform.Transform(1D, 1D);
        double left = Math.Min(Math.Min(p0.X, p1.X), Math.Min(p2.X, p3.X));
        double right = Math.Max(Math.Max(p0.X, p1.X), Math.Max(p2.X, p3.X));
        double bottom = Math.Min(Math.Min(p0.Y, p1.Y), Math.Min(p2.Y, p3.Y));
        double top = Math.Max(Math.Max(p0.Y, p1.Y), Math.Max(p2.Y, p3.Y));
        x = left;
        y = bottom;
        width = Math.Max(0D, right - left);
        height = Math.Max(0D, top - bottom);
    }

    private static void AddRemovedImageTargets(IReadOnlyList<ImageRedactionTarget> targets, List<PdfRedactionMatch> removedMatches, HashSet<string>? removedResourceNames) {
        for (int i = 0; i < targets.Count; i++) {
            if (!removedMatches.Contains(targets[i].Match)) {
                removedMatches.Add(targets[i].Match);
            }

            removedResourceNames?.Add(targets[i].ResourceName);
        }
    }

    private static void RemoveUnusedImageResourcesFromXObjects(Dictionary<int, PdfIndirectObject> objects, PdfDictionary xObjects, string content, IReadOnlyList<ImageRedactionTarget> removedTargets) {
        for (int i = 0; i < removedTargets.Count; i++) {
            string resourceName = removedTargets[i].ResourceName;
            if (ContentInvokesResource(content, resourceName) ||
                !xObjects.Items.TryGetValue(resourceName, out PdfObject? resourceObject) ||
                PdfObjectLookup.Resolve(objects, resourceObject) is not PdfStream stream ||
                !string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Image", StringComparison.Ordinal)) {
                continue;
            }

            xObjects.Items.Remove(resourceName);
        }
    }

    private static void RemoveUnusedPageImageResources(Dictionary<int, PdfIndirectObject> objects, PdfDictionary pageDictionary, HashSet<string> resourceNames) {
        PdfDictionary xObjects = PdfPageResourceHelper.EnsurePageXObjects(objects, pageDictionary, "redaction image cleanup");
        string remainingContent = GetPageContent(objects, pageDictionary);
        foreach (string resourceName in resourceNames) {
            if (ContentInvokesResource(remainingContent, resourceName) ||
                !xObjects.Items.TryGetValue(resourceName, out PdfObject? resourceObject) ||
                PdfObjectLookup.Resolve(objects, resourceObject) is not PdfStream stream ||
                !string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Image", StringComparison.Ordinal)) {
                continue;
            }

            xObjects.Items.Remove(resourceName);
        }
    }

    private static string GetPageContent(Dictionary<int, PdfIndirectObject> objects, PdfDictionary pageDictionary) {
        if (!pageDictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject)) {
            return string.Empty;
        }

        var builder = new System.Text.StringBuilder();
        foreach (PdfReference reference in EnumerateContentReferences(objects, contentsObject)) {
            if (PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
                indirect.Value is PdfStream stream &&
                !stream.DecodingFailed) {
                builder.Append(PdfEncoding.Latin1GetString(StreamDecoder.Decode(stream.Dictionary, stream.Data, objects)));
                builder.Append('\n');
            }
        }

        return builder.ToString();
    }

    private static bool ContentInvokesResource(string content, string resourceName) {
        foreach (TextContentParser.FormInvocation invocation in TextContentParser.ExtractFormInvocations(content)) {
            if (string.Equals(invocation.Name, resourceName, StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    private static int CountResourceInvocations(string content, string resourceName) {
        int count = 0;
        foreach (TextContentParser.FormInvocation invocation in TextContentParser.ExtractFormInvocations(content)) {
            if (string.Equals(invocation.Name, resourceName, StringComparison.Ordinal)) {
                count++;
            }
        }

        return count;
    }

    private static string RemoveRanges(string content, List<RemovalRange> ranges) {
        if (ranges.Count == 0) {
            return content;
        }

        ranges.Sort((left, right) => right.Start.CompareTo(left.Start));
        var builder = new System.Text.StringBuilder(content);
        for (int i = 0; i < ranges.Count; i++) {
            builder.Remove(ranges[i].Start, ranges[i].End - ranges[i].Start);
        }

        return builder.ToString();
    }

    private static void SkipWhiteSpace(string content, ref int index) {
        while (index < content.Length && char.IsWhiteSpace(content[index])) {
            index++;
        }
    }

    private static void SkipComment(string content, ref int index) {
        while (index < content.Length && content[index] != '\n' && content[index] != '\r') {
            index++;
        }
    }

    private static ImageContentOperand ReadNameOperand(string content, ref int index) {
        int start = index;
        index++;
        int nameStart = index;
        while (index < content.Length) {
            char value = content[index];
            if (char.IsWhiteSpace(value) || value == '%' || value == '/' || value == '[' || value == ']' || value == '(' || value == ')' || value == '<' || value == '>') {
                break;
            }

            index++;
        }

        return ImageContentOperand.ForName(PdfSyntax.DecodeName(content.Substring(nameStart, index - nameStart)), start, index);
    }

    private static ImageContentOperand ReadNumberOperand(string content, ref int index) {
        int start = index;
        index++;
        while (index < content.Length) {
            char value = content[index];
            if (!(IsDigit(value) || value == '.' || value == 'E' || value == 'e' || value == '-' || value == '+')) {
                break;
            }

            index++;
        }

        string numberText = content.Substring(start, index - start);
        if (!double.TryParse(numberText, NumberStyles.Any, CultureInfo.InvariantCulture, out double number)) {
            number = 0D;
        }

        return ImageContentOperand.ForNumber(number, start, index);
    }

    private static string ReadOperator(string content, ref int index) {
        int start = index;
        while (index < content.Length && !char.IsWhiteSpace(content[index]) && content[index] != '/' && content[index] != '[' && content[index] != ']' && content[index] != '(' && content[index] != ')' && content[index] != '<' && content[index] != '>') {
            index++;
        }

        return content.Substring(start, index - start);
    }

    private static void SkipLiteralString(string content, ref int index) {
        index++;
        int depth = 1;
        bool escaped = false;
        while (index < content.Length && depth > 0) {
            char value = content[index++];
            if (escaped) {
                escaped = false;
            } else if (value == '\\') {
                escaped = true;
            } else if (value == '(') {
                depth++;
            } else if (value == ')') {
                depth--;
            }
        }
    }

    private static void SkipHexString(string content, ref int index) {
        index++;
        while (index < content.Length && content[index] != '>') {
            index++;
        }

        if (index < content.Length) {
            index++;
        }
    }

    private static void SkipArray(string content, ref int index) {
        index++;
        int depth = 1;
        while (index < content.Length && depth > 0) {
            char value = content[index];
            if (value == '(') {
                SkipLiteralString(content, ref index);
            } else if (value == '<') {
                if (index + 1 < content.Length && content[index + 1] == '<') {
                    SkipDictionary(content, ref index);
                } else {
                    SkipHexString(content, ref index);
                }
            } else {
                if (value == '[') {
                    depth++;
                } else if (value == ']') {
                    depth--;
                }

                index++;
            }
        }
    }

    private static void SkipDictionary(string content, ref int index) {
        index += 2;
        int depth = 1;
        while (index < content.Length && depth > 0) {
            char value = content[index];
            if (value == '(') {
                SkipLiteralString(content, ref index);
            } else if (value == '<' && index + 1 < content.Length && content[index + 1] == '<') {
                index += 2;
                depth++;
            } else if (value == '>' && index + 1 < content.Length && content[index + 1] == '>') {
                index += 2;
                depth--;
            } else if (value == '<') {
                SkipHexString(content, ref index);
            } else {
                index++;
            }
        }
    }

    private static bool IsDigit(char value) => value >= '0' && value <= '9';

    private static bool IsNumberStart(char value) => value == '-' || value == '+' || value == '.' || IsDigit(value);

    private static bool AreCloseImageCoordinate(double left, double right) => Math.Abs(left - right) <= ImageRedactionTolerance;

    private readonly struct ImageRedactionTarget {
        public ImageRedactionTarget(PdfRedactionMatch match, string resourceName, double x, double y, double width, double height) {
            Match = match;
            ResourceName = resourceName;
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        public PdfRedactionMatch Match { get; }

        public string ResourceName { get; }

        public double X { get; }

        public double Y { get; }

        public double Width { get; }

        public double Height { get; }
    }

    private readonly struct ImageContentOperand {
        private ImageContentOperand(string? name, double number, int start, int end) {
            Name = name;
            Number = number;
            Start = start;
            End = end;
        }

        public string? Name { get; }

        public double Number { get; }

        public int Start { get; }

        public int End { get; }

        public static ImageContentOperand ForName(string name, int start, int end) => new ImageContentOperand(name, 0D, start, end);

        public static ImageContentOperand ForNumber(double number, int start, int end) => new ImageContentOperand(null, number, start, end);
    }

    private readonly struct RemovalRange {
        public RemovalRange(int start, int end) {
            Start = start;
            End = end;
        }

        public int Start { get; }

        public int End { get; }
    }

    private readonly struct ImageRedactionMutation {
        public static readonly ImageRedactionMutation None = new ImageRedactionMutation(false, Array.Empty<PdfRedactionMatch>());

        public ImageRedactionMutation(bool hasChanges, IReadOnlyList<PdfRedactionMatch> removedMatches) {
            HasChanges = hasChanges;
            RemovedMatches = removedMatches;
        }

        public bool HasChanges { get; }

        public IReadOnlyList<PdfRedactionMatch> RemovedMatches { get; }
    }
}
