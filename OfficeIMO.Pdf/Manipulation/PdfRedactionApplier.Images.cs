using OfficeIMO.Pdf.Filters;
using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private const double ImageRedactionTolerance = 0.01D;

    private static ImageRedactionMutation RemoveMatchedImageObjects(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        IReadOnlyList<PdfRedactionMatch> matches,
        PdfRedactionApplyOptions options,
        PdfReadLimits limits,
        ref int nextObjectNumber) {
        int maximumDecodedStreamBytes = limits.MaxDecodedStreamBytes;
        bool removeWholeIntersectingImages = options.UnsupportedImagePolicy == PdfRedactionUnsupportedImagePolicy.RemoveWholePlacement;
        ImageRedactionTarget[] wholeImageTargets = BuildWholeImageTargets(matches, removeWholeIntersectingImages);
        ImageRedactionTarget[] pixelTargets = BuildPixelImageTargets(matches);
        if ((wholeImageTargets.Length == 0 && pixelTargets.Length == 0) ||
            !pageDictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject)) {
            return ImageRedactionMutation.None;
        }

        bool changed = false;
        var removedMatches = new List<PdfRedactionMatch>();
        var removedResourceNames = new HashSet<string>(StringComparer.Ordinal);
        Dictionary<int, int> referenceCounts = CountIndirectReferenceUsage(objects);
        PdfDictionary? pageResources = GetInheritedDictionary(objects, pageDictionary, "Resources");
        PdfDictionary? pageXObjects = pageResources != null && pageResources.Items.TryGetValue("XObject", out PdfObject? pageXObjectValue)
            ? ResolveDictionary(objects, pageXObjectValue)
            : null;
        PdfObject currentContentsObject = contentsObject;
        bool passChanged;
        do {
            passChanged = false;
            var contentState = new ImageContentGraphicsState(Matrix2D.Identity);
            PdfReference[] contentReferences = EnumerateContentReferences(objects, currentContentsObject).ToArray();
            foreach (PdfReference reference in contentReferences) {
                if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                    indirect.Value is not PdfStream stream ||
                    stream.DecodingFailed) {
                    continue;
                }

                byte[] contentBytes = StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects, maximumDecodedStreamBytes);
                string content = PdfEncoding.Latin1GetString(contentBytes);
                string scrubbed = RemoveImageInvocations(content, wholeImageTargets, pageXObjects, objects, contentState, out IReadOnlyList<ImageRedactionTarget> removedTargets);
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
            RemoveUnusedPageImageResources(objects, pageDictionary, removedResourceNames, limits);
        }

        changed = ScrubMatchedImageFormXObjects(objects, pageDictionary, currentContentsObject, wholeImageTargets, referenceCounts, removedMatches, limits, ref nextObjectNumber) || changed;
        changed = RewriteMatchedImagePixels(objects, pageDictionary, currentContentsObject, pixelTargets, options, referenceCounts, removedMatches, limits, ref nextObjectNumber) || changed;
        return new ImageRedactionMutation(changed, removedMatches.AsReadOnly());
    }

    private static ImageRedactionTarget[] BuildWholeImageTargets(IReadOnlyList<PdfRedactionMatch> matches, bool removeWholeIntersectingImages) {
        return matches
            .Where(match => match.Kind == PdfRedactionMatchKind.ImagePlacement &&
                !string.IsNullOrEmpty(match.ResourceName) &&
                (removeWholeIntersectingImages ? RedactionAreaIntersectsMatch(match.Area, match) : RedactionAreaCoversMatch(match.Area, match)))
            .Select(match => new ImageRedactionTarget(match))
            .ToArray();
    }

    private static ImageRedactionTarget[] BuildPixelImageTargets(IReadOnlyList<PdfRedactionMatch> matches) {
        return matches
            .Where(match => match.Kind == PdfRedactionMatchKind.ImagePlacement &&
                !string.IsNullOrEmpty(match.ResourceName) &&
                RedactionAreaIntersectsMatch(match.Area, match))
            .Select(match => new ImageRedactionTarget(match))
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
        PdfReadLimits limits,
        ref int nextObjectNumber) {
        int maximumDecodedStreamBytes = limits.MaxDecodedStreamBytes;
        PdfDictionary? resources = GetInheritedDictionary(objects, pageDictionary, "Resources");
        if (resources is null ||
            !resources.Items.ContainsKey("XObject")) {
            return false;
        }

        PdfDictionary xObjects = PdfPageResourceHelper.EnsurePageXObjects(objects, pageDictionary, "redaction nested image cleanup");
        resources = ResolveDictionary(objects, pageDictionary.Items.TryGetValue("Resources", out PdfObject? pageResources) ? pageResources : null) ?? resources;
        bool changed = false;
        PdfObject currentContentsObject = contentsObject;
        var contentState = new ImageContentGraphicsState(Matrix2D.Identity);
        foreach (PdfReference reference in EnumerateContentReferences(objects, currentContentsObject)) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                indirect.Value is not PdfStream stream ||
                stream.DecodingFailed) {
                continue;
            }

            string content = PdfEncoding.Latin1GetString(StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects, maximumDecodedStreamBytes));
            ImagePixelRewriteContentResult result = ScrubImageFormInvocations(objects, resources, xObjects, content, targets, contentState, referenceCounts, new HashSet<int>(), removedMatches, limits, ref nextObjectNumber);
            if (!string.Equals(result.Content, content, StringComparison.Ordinal)) {
                PdfReference targetReference = reference;
                if (IsSharedReference(referenceCounts, reference)) {
                    targetReference = CloneIndirectObject(objects, reference, indirect, ref nextObjectNumber);
                    ReplacePageContentReference(objects, pageDictionary, currentContentsObject, reference, targetReference);
                    currentContentsObject = pageDictionary.Items.TryGetValue("Contents", out PdfObject? updatedContentsObject)
                        ? updatedContentsObject
                        : currentContentsObject;
                }

                objects[targetReference.ObjectNumber] = new PdfIndirectObject(targetReference.ObjectNumber, targetReference.Generation, new PdfStream(CleanStreamDictionary(stream.Dictionary), PdfEncoding.Latin1GetBytes(result.Content)));
            }

            changed = result.HasChanges || changed;
        }

        return changed;
    }

    private static ImagePixelRewriteContentResult ScrubImageFormInvocations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary resources,
        PdfDictionary xObjects,
        string content,
        ImageRedactionTarget[] targets,
        ImageContentGraphicsState graphicsState,
        IReadOnlyDictionary<int, int> referenceCounts,
        HashSet<int> activeForms,
        List<PdfRedactionMatch> removedMatches,
        PdfReadLimits limits,
        ref int nextObjectNumber) {
        int maximumDecodedStreamBytes = limits.MaxDecodedStreamBytes;
        bool changed = false;
        string rewrittenContent = content;
        ImageResourceInvocation[] invocations = ExtractImageResourceInvocations(content, graphicsState);
        for (int invocationIndex = invocations.Length - 1; invocationIndex >= 0; invocationIndex--) {
            ImageResourceInvocation invocation = invocations[invocationIndex];
            if (!TryGetFormXObject(objects, xObjects, invocation.Name, out PdfReference reference, out PdfStream formStream) ||
                formStream.DecodingFailed ||
                activeForms.Contains(reference.ObjectNumber)) {
                continue;
            }

            Matrix2D invocationTransform = invocation.Transform;
            bool repeatedInvocation = CountResourceInvocations(content, invocation.Name, limits) != 1;
            if (repeatedInvocation &&
                PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? repeatedSourceIndirect)) {
                string resourceName = CreateUniqueResourceName(xObjects, invocation.Name);
                reference = CloneIndirectObject(objects, reference, repeatedSourceIndirect, ref nextObjectNumber);
                xObjects.Items[resourceName] = reference;
                formStream = (PdfStream)objects[reference.ObjectNumber].Value;

                if (!activeForms.Add(reference.ObjectNumber)) {
                    continue;
                }

                int repeatedObjectNumber = reference.ObjectNumber;
                try {
                    if (ScrubImageForm(objects, resources, reference, formStream, targets, invocationTransform, referenceCounts, activeForms, removedMatches, limits, ref nextObjectNumber).HasChanges) {
                        rewrittenContent = ReplaceInvocationResourceName(rewrittenContent, invocation, resourceName);
                        changed = true;
                    }
                } finally {
                    activeForms.Remove(repeatedObjectNumber);
                }

                continue;
            }

            if (!activeForms.Add(reference.ObjectNumber)) {
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

                changed = ScrubImageForm(objects, resources, reference, formStream, targets, invocationTransform, referenceCounts, activeForms, removedMatches, limits, ref nextObjectNumber).HasChanges || changed;
            } finally {
                activeForms.Remove(activeObjectNumber);
            }
        }

        return new ImagePixelRewriteContentResult(changed, rewrittenContent);
    }

    private static ImagePixelRewriteContentResult ScrubImageForm(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary inheritedResources,
        PdfReference formReference,
        PdfStream formStream,
        ImageRedactionTarget[] targets,
        Matrix2D invocationTransform,
        IReadOnlyDictionary<int, int> referenceCounts,
        HashSet<int> activeForms,
        List<PdfRedactionMatch> removedMatches,
        PdfReadLimits limits,
        ref int nextObjectNumber) {
        int maximumDecodedStreamBytes = limits.MaxDecodedStreamBytes;
        PdfDictionary formResources;
        if (formStream.Dictionary.Items.ContainsKey("Resources")) {
            formResources = EnsureFormResources(objects, formStream);
        } else {
            formResources = CloneDictionary(inheritedResources);
            formStream.Dictionary.Items["Resources"] = formResources;
        }
        PdfDictionary formXObjects = EnsureResourceXObjects(objects, formResources);
        Matrix2D formTransform = ApplyFormMatrix(invocationTransform, formStream.Dictionary);
        string formContent = PdfEncoding.Latin1GetString(StreamDecoder.DecodeRequired(formStream.Dictionary, formStream.Data, objects, maximumDecodedStreamBytes));
        string scrubbed = RemoveImageInvocations(formContent, targets, formXObjects, objects, new ImageContentGraphicsState(formTransform), out IReadOnlyList<ImageRedactionTarget> removedTargets);
        bool changed = false;

        if (!string.Equals(formContent, scrubbed, StringComparison.Ordinal)) {
            formContent = scrubbed;
            AddRemovedImageTargets(removedTargets, removedMatches, null);
            RemoveUnusedImageResourcesFromXObjects(objects, formXObjects, formContent, removedTargets, limits);
            changed = true;
        }

        ImagePixelRewriteContentResult nestedResult = ScrubImageFormInvocations(objects, formResources, formXObjects, formContent, targets, new ImageContentGraphicsState(formTransform), referenceCounts, activeForms, removedMatches, limits, ref nextObjectNumber);
        if (!string.Equals(nestedResult.Content, formContent, StringComparison.Ordinal)) {
            formContent = nestedResult.Content;
            changed = true;
        }

        if (changed || nestedResult.HasChanges) {
            objects[formReference.ObjectNumber] = new PdfIndirectObject(formReference.ObjectNumber, formReference.Generation, new PdfStream(CleanStreamDictionary(formStream.Dictionary), PdfEncoding.Latin1GetBytes(formContent)));
        }

        return new ImagePixelRewriteContentResult(changed || nestedResult.HasChanges, formContent);
    }

    private static string RemoveImageInvocations(
        string content,
        ImageRedactionTarget[] targets,
        PdfDictionary? xObjects,
        Dictionary<int, PdfIndirectObject> objects,
        ImageContentGraphicsState graphicsState,
        out IReadOnlyList<ImageRedactionTarget> removedTargets) {
        var ranges = new List<RemovalRange>();
        var removed = new List<ImageRedactionTarget>();
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
                case "BI":
                    SkipInlineImage(content, ref index);
                    args.Clear();
                    break;
                case "q":
                    graphicsState.Stack.Push(graphicsState.Transform);
                    args.Clear();
                    break;
                case "Q":
                    graphicsState.Transform = graphicsState.Stack.Count > 0 ? graphicsState.Stack.Pop() : graphicsState.BaseTransform;
                    args.Clear();
                    break;
                case "cm":
                    if (args.Count >= 6) {
                        graphicsState.Transform = Matrix2D.Multiply(graphicsState.Transform, new Matrix2D(
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
                    if (TryGetImageTarget(args, graphicsState.Transform, targets, xObjects, objects, out ImageRedactionTarget target, out ImageContentOperand operand)) {
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

    private static bool TryGetImageTarget(
        IReadOnlyList<ImageContentOperand> args,
        Matrix2D ctm,
        ImageRedactionTarget[] targets,
        PdfDictionary? xObjects,
        Dictionary<int, PdfIndirectObject> objects,
        out ImageRedactionTarget target,
        out ImageContentOperand operand) {
        target = default;
        operand = default;
        if (args.Count == 0 || string.IsNullOrEmpty(args[args.Count - 1].Name)) {
            return false;
        }

        operand = args[args.Count - 1];
        (int ObjectNumber, int DirectStreamIdentity) invocationIdentity = ResolveImageInvocationIdentity(operand.Name!, xObjects, objects);
        GetUnitRectangleBounds(ctm, out double x, out double y, out double width, out double height);
        for (int i = 0; i < targets.Length; i++) {
            if (string.Equals(targets[i].ResourceName, operand.Name, StringComparison.Ordinal) &&
                targets[i].MatchesIdentity(invocationIdentity.ObjectNumber, invocationIdentity.DirectStreamIdentity) &&
                targets[i].MatchesTransform(ctm) &&
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
                targets[i].MatchesIdentity(invocationIdentity.ObjectNumber, invocationIdentity.DirectStreamIdentity) &&
                targets[i].MatchesTransform(ctm) &&
                RedactionAreaCoversRectangle(targets[i].Match.Area, x, y, width, height)) {
                target = targets[i];
                return true;
            }
        }

        return false;
    }

    private static (int ObjectNumber, int DirectStreamIdentity) ResolveImageInvocationIdentity(string resourceName, PdfDictionary? xObjects, Dictionary<int, PdfIndirectObject> objects) {
        if (xObjects is null || !xObjects.Items.TryGetValue(resourceName, out PdfObject? value)) return (0, 0);
        if (value is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
            indirect.Value is PdfStream stream &&
            string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Image", StringComparison.Ordinal)) {
            return (reference.ObjectNumber, 0);
        }
        return value is PdfStream directStream &&
            string.Equals(directStream.Dictionary.Get<PdfName>("Subtype")?.Name, "Image", StringComparison.Ordinal)
                ? (0, PdfDirectStreamIdentity.Compute(directStream))
                : (int.MinValue, 0);
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

    private static void RemoveUnusedImageResourcesFromXObjects(Dictionary<int, PdfIndirectObject> objects, PdfDictionary xObjects, string content, IReadOnlyList<ImageRedactionTarget> removedTargets, PdfReadLimits limits) {
        for (int i = 0; i < removedTargets.Count; i++) {
            string resourceName = removedTargets[i].ResourceName;
            if (ContentOrInheritedFormsInvokeResource(objects, xObjects, content, resourceName, limits, new HashSet<int>()) ||
                !xObjects.Items.TryGetValue(resourceName, out PdfObject? resourceObject) ||
                PdfObjectLookup.Resolve(objects, resourceObject) is not PdfStream stream ||
                !string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Image", StringComparison.Ordinal)) {
                continue;
            }

            xObjects.Items.Remove(resourceName);
        }
    }

    private static void RemoveUnusedPageImageResources(Dictionary<int, PdfIndirectObject> objects, PdfDictionary pageDictionary, HashSet<string> resourceNames, PdfReadLimits limits) {
        PdfDictionary xObjects = PdfPageResourceHelper.EnsurePageXObjects(objects, pageDictionary, "redaction image cleanup");
        string remainingContent = GetPageContent(objects, pageDictionary, limits.MaxDecodedStreamBytes);
        foreach (string resourceName in resourceNames) {
            if (ContentOrInheritedFormsInvokeResource(objects, xObjects, remainingContent, resourceName, limits, new HashSet<int>()) ||
                !xObjects.Items.TryGetValue(resourceName, out PdfObject? resourceObject) ||
                PdfObjectLookup.Resolve(objects, resourceObject) is not PdfStream stream ||
                !string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Image", StringComparison.Ordinal)) {
                continue;
            }

            xObjects.Items.Remove(resourceName);
        }
    }

    private static string GetPageContent(Dictionary<int, PdfIndirectObject> objects, PdfDictionary pageDictionary, int maximumDecodedStreamBytes) {
        if (!pageDictionary.Items.TryGetValue("Contents", out PdfObject? contentsObject)) {
            return string.Empty;
        }

        var builder = new System.Text.StringBuilder();
        foreach (PdfReference reference in EnumerateContentReferences(objects, contentsObject)) {
            if (PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
                indirect.Value is PdfStream stream &&
                !stream.DecodingFailed) {
                builder.Append(PdfEncoding.Latin1GetString(StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects, maximumDecodedStreamBytes)));
                builder.Append('\n');
            }
        }

        return builder.ToString();
    }

    private static bool ContentInvokesResource(string content, string resourceName, PdfReadLimits limits) {
        foreach (TextContentParser.FormInvocation invocation in ExtractFormInvocations(content, limits)) {
            if (string.Equals(invocation.Name, resourceName, StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    private static bool ContentOrInheritedFormsInvokeResource(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary xObjects,
        string content,
        string resourceName,
        PdfReadLimits limits,
        HashSet<int> activeForms) {
        if (ContentInvokesResource(content, resourceName, limits)) return true;

        foreach (TextContentParser.FormInvocation invocation in ExtractFormInvocations(content, limits)) {
            if (!xObjects.Items.TryGetValue(invocation.Name, out PdfObject? formObject) ||
                formObject is not PdfReference formReference ||
                !PdfObjectLookup.TryGet(objects, formReference, out PdfIndirectObject? indirect) ||
                indirect.Value is not PdfStream form ||
                !string.Equals(form.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal) ||
                form.Dictionary.Items.ContainsKey("Resources") ||
                !activeForms.Add(formReference.ObjectNumber)) {
                continue;
            }

            try {
                string formContent = PdfEncoding.Latin1GetString(
                    StreamDecoder.DecodeRequired(form.Dictionary, form.Data, objects, limits.MaxDecodedStreamBytes));
                if (ContentOrInheritedFormsInvokeResource(
                        objects,
                        xObjects,
                        formContent,
                        resourceName,
                        limits,
                        activeForms)) {
                    return true;
                }
            } finally {
                activeForms.Remove(formReference.ObjectNumber);
            }
        }

        return false;
    }

    private static int CountResourceInvocations(string content, string resourceName, PdfReadLimits limits) {
        int count = 0;
        foreach (TextContentParser.FormInvocation invocation in ExtractFormInvocations(content, limits)) {
            if (string.Equals(invocation.Name, resourceName, StringComparison.Ordinal)) {
                count++;
            }
        }

        return count;
    }

    private static List<TextContentParser.FormInvocation> ExtractFormInvocations(string content, PdfReadLimits limits) =>
        TextContentParser.ExtractFormInvocations(
            content,
            maxOperations: limits.MaxContentOperations,
            maxNestingDepth: limits.MaxContentNestingDepth,
            maxOperands: limits.MaxContentOperands);

    private static bool RemoveUnusedImageObjectReferences(Dictionary<int, PdfIndirectObject> objects, HashSet<int> targetObjectNumbers, PdfReadLimits limits) {
        var invokedNames = new HashSet<string>(StringComparer.Ordinal);
        var pageContentStreamNumbers = new HashSet<int>();
        var type3CharProcStreams = new HashSet<PdfStream>();
        var scannedStreams = new HashSet<PdfStream>();
        foreach (PdfIndirectObject candidate in objects.Values) {
            if (candidate.Value is not PdfDictionary page ||
                !string.Equals(page.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal) ||
                !page.Items.TryGetValue("Contents", out PdfObject? contents)) continue;
            foreach (PdfReference reference in EnumerateContentStreamReferences(objects, contents)) {
                pageContentStreamNumbers.Add(reference.ObjectNumber);
            }
        }
        foreach (PdfIndirectObject candidate in objects.Values) {
            PdfDictionary? dictionary = candidate.Value switch {
                PdfDictionary value => value,
                PdfStream value => value.Dictionary,
                _ => null
            };
            if (dictionary == null ||
                !string.Equals(dictionary.Get<PdfName>("Subtype")?.Name, "Type3", StringComparison.Ordinal) ||
                !dictionary.Items.TryGetValue("CharProcs", out PdfObject? charProcsObject) ||
                PdfObjectLookup.Resolve(objects, charProcsObject) is not PdfDictionary charProcs) continue;
            foreach (PdfObject charProcObject in charProcs.Items.Values) {
                if (PdfObjectLookup.Resolve(objects, charProcObject) is PdfStream charProc) type3CharProcStreams.Add(charProc);
            }
        }
        foreach (PdfIndirectObject indirect in objects.Values) {
            if (indirect.Value is not PdfStream stream ||
                (!pageContentStreamNumbers.Contains(indirect.ObjectNumber) &&
                 !type3CharProcStreams.Contains(stream) &&
                 !CanOwnContentInvocations(stream.Dictionary)) ||
                stream.DecodingFailed || StreamDecoder.GetUnsupportedFilters(stream.Dictionary, objects).Count != 0) continue;
            ScanContentInvocations(stream);
        }
        foreach (PdfStream charProc in type3CharProcStreams) ScanContentInvocations(charProc);
        bool changed = false;
        foreach (PdfIndirectObject indirect in objects.Values) changed = RemoveUnusedImageEntries(indirect.Value, objects, targetObjectNumbers, invokedNames) || changed;
        return changed;

        void ScanContentInvocations(PdfStream stream) {
            if (!scannedStreams.Add(stream) ||
                stream.DecodingFailed ||
                StreamDecoder.GetUnsupportedFilters(stream.Dictionary, objects).Count != 0) return;
            string content = PdfEncoding.Latin1GetString(StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects, limits.MaxDecodedStreamBytes));
            foreach (TextContentParser.FormInvocation invocation in TextContentParser.ExtractFormInvocations(
                content,
                maxOperations: limits.MaxContentOperations,
                maxNestingDepth: limits.MaxContentNestingDepth,
                maxOperands: limits.MaxContentOperands)) {
                invokedNames.Add(invocation.Name);
            }
        }

        static bool CanOwnContentInvocations(PdfDictionary dictionary) {
            string? subtype = dictionary.Get<PdfName>("Subtype")?.Name;
            string? type = dictionary.Get<PdfName>("Type")?.Name;
            if (string.Equals(subtype, "Image", StringComparison.Ordinal) ||
                string.Equals(type, "EmbeddedFile", StringComparison.Ordinal) ||
                string.Equals(type, "ObjStm", StringComparison.Ordinal) ||
                string.Equals(type, "XRef", StringComparison.Ordinal)) return false;
            return string.Equals(subtype, "Form", StringComparison.Ordinal) ||
                dictionary.Get<PdfNumber>("PatternType")?.Value == 1D;
        }
    }

    private static IEnumerable<PdfReference> EnumerateContentStreamReferences(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject contents) {
        if (contents is PdfReference reference) {
            if (PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) && indirect.Value is PdfArray array) {
                for (int i = 0; i < array.Items.Count; i++) {
                    foreach (PdfReference item in EnumerateContentStreamReferences(objects, array.Items[i])) yield return item;
                }
            } else {
                yield return reference;
            }
        } else if (contents is PdfArray directArray) {
            for (int i = 0; i < directArray.Items.Count; i++) {
                foreach (PdfReference item in EnumerateContentStreamReferences(objects, directArray.Items[i])) yield return item;
            }
        }
    }

    private static bool RemoveUnusedImageEntries(PdfObject value, Dictionary<int, PdfIndirectObject> objects, HashSet<int> targets, HashSet<string> invokedNames) {
        PdfDictionary? dictionary = value is PdfDictionary direct ? direct : value is PdfStream stream ? stream.Dictionary : null; if (dictionary is null) return false;
        bool changed = false;
        if (dictionary.Items.TryGetValue("XObject", out PdfObject? xObjectObject) && ResolveDictionary(objects, xObjectObject) is PdfDictionary xObjects) {
            foreach (string name in xObjects.Items.Keys.ToArray()) if (!invokedNames.Contains(name) && xObjects.Items[name] is PdfReference reference && targets.Contains(reference.ObjectNumber)) { xObjects.Items.Remove(name); changed = true; }
        }
        foreach (PdfObject child in dictionary.Items.Values.ToArray()) if (child is not PdfReference) changed = RemoveUnusedImageEntries(child, objects, targets, invokedNames) || changed;
        return changed;
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

    private static void SkipInlineImage(string content, ref int index) {
        while (index < content.Length) {
            SkipWhiteSpace(content, ref index);
            if (index >= content.Length) return;
            if (content[index] == '%') {
                SkipComment(content, ref index);
                continue;
            }
            if (content[index] == '/') {
                _ = ReadNameOperand(content, ref index);
                continue;
            }
            if (content[index] == '(') {
                SkipLiteralString(content, ref index);
                continue;
            }
            if (content[index] == '<') {
                if (index + 1 < content.Length && content[index + 1] == '<') SkipDictionary(content, ref index);
                else SkipHexString(content, ref index);
                continue;
            }
            if (content[index] == '[') {
                SkipArray(content, ref index);
                continue;
            }

            string token = ReadOperator(content, ref index);
            if (!string.Equals(token, "ID", StringComparison.Ordinal)) continue;
            if (index < content.Length && char.IsWhiteSpace(content[index])) index++;
            int dataLength = PdfInlineImageDataScanner.FindLength(content, index);
            if (dataLength < 0) {
                index = content.Length;
                return;
            }
            index += dataLength;
            SkipWhiteSpace(content, ref index);
            if (PdfInlineImageDataScanner.IsTerminatorAt(content, index)) index += 2;
            return;
        }
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
        public ImageRedactionTarget(PdfRedactionMatch match) {
            Match = match;
            ResourceName = match.ResourceName!;
            X = match.X;
            Y = match.Y;
            Width = match.Width;
            Height = match.Height;
            Placement = match.ImagePlacement;
        }

        public PdfRedactionMatch Match { get; }

        public string ResourceName { get; }

        public double X { get; }

        public double Y { get; }

        public double Width { get; }

        public double Height { get; }

        private PdfImagePlacement? Placement { get; }

        public bool MatchesObjectNumber(int objectNumber) => Placement is null || Placement.ObjectNumber == objectNumber;

        public bool MatchesIdentity(int objectNumber, int directStreamIdentity) => Placement is null ||
            (Placement.ObjectNumber == objectNumber &&
             (objectNumber != 0 || Placement.DirectStreamIdentity == directStreamIdentity));

        public bool MatchesTransform(Matrix2D transform) => Placement is null ||
            (AreCloseImageCoordinate(Placement.A, transform.A) &&
             AreCloseImageCoordinate(Placement.B, transform.B) &&
             AreCloseImageCoordinate(Placement.C, transform.C) &&
             AreCloseImageCoordinate(Placement.D, transform.D) &&
             AreCloseImageCoordinate(Placement.E, transform.E) &&
             AreCloseImageCoordinate(Placement.F, transform.F));
    }

    private sealed class ImageContentGraphicsState {
        public ImageContentGraphicsState(Matrix2D baseTransform) {
            BaseTransform = baseTransform;
            Transform = baseTransform;
        }

        public Matrix2D BaseTransform { get; }
        public Matrix2D Transform { get; set; }
        public Stack<Matrix2D> Stack { get; } = new Stack<Matrix2D>();
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
