namespace OfficeIMO.Pdf;

/// <summary>Marked-content, font and XObject evidence that the accessibility checker needs from one page.</summary>
internal sealed class PdfAccessibilityPageContent {
    /// <summary>Content items (text, painted paths, images, shadings) outside both tagged content and artifacts.</summary>
    internal int UntaggedItems { get; set; }
    /// <summary>Artifact sequences opened inside tagged content.</summary>
    internal int ArtifactsInTaggedContent { get; set; }
    /// <summary>MCID sequences opened inside an artifact.</summary>
    internal int TaggedContentInArtifacts { get; set; }
    /// <summary>Artifact sequences opened inside an MCID that the structure tree does not reference.</summary>
    internal int ArtifactsInUnlinkedMcid { get; set; }
    /// <summary>MCIDs enclosing shown text with no enclosing marked-content Lang, keyed by owning stream (null for page content).</summary>
    internal HashSet<(int? Stream, int Mcid)> TextMcids { get; } = new();
    /// <summary>Text-showing operators outside tagged content and artifacts with no enclosing marked-content Lang.</summary>
    internal int UntaggedTextWithoutLanguage { get; set; }
    /// <summary>Marked-content sequences whose property list carries Alt, ActualText or E without Lang, with their nearest MCID.</summary>
    internal List<(int? Stream, int? Mcid, string Key)> PropertyTextWithoutLang { get; } = new();
    /// <summary>Lang values declared in marked-content property lists.</summary>
    internal List<string> PropertyLanguages { get; } = new();
    /// <summary>Fonts selected by Tf and used by a text-showing operator; the flag is true when used with a visible rendering mode.</summary>
    internal Dictionary<PdfDictionary, bool> Fonts { get; } = new(ReferenceEqualityComparer<PdfDictionary>.Instance);
    /// <summary>Object numbers of fonts in <see cref="Fonts"/>, when indirect.</summary>
    internal Dictionary<PdfDictionary, int?> FontObjectNumbers { get; } = new(ReferenceEqualityComparer<PdfDictionary>.Instance);
    /// <summary>Form XObject invocations on this page, keyed by object number.</summary>
    internal Dictionary<int, int> FormInvocations { get; } = new();
    /// <summary>Form XObjects whose own content declares MCIDs.</summary>
    internal HashSet<int> FormsWithMcids { get; } = new();
    /// <summary>Reference XObjects (Form XObjects with a Ref entry) painted on this page.</summary>
    internal HashSet<int> ReferenceXObjects { get; } = new();

    /// <summary>Adds invocation counts without wrapping, so a large count can never read as a small one.</summary>
    internal static int AddCount(int current, int added) => current > int.MaxValue - added ? int.MaxValue : current + added;

    internal sealed class ReferenceEqualityComparer<T> : IEqualityComparer<T> where T : class {
        internal static ReferenceEqualityComparer<T> Instance { get; } = new();
        public bool Equals(T? x, T? y) => ReferenceEquals(x, y);
        public int GetHashCode(T obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
    }
}

public sealed partial class PdfReadPage {
    private readonly struct AccessibilityScope {
        internal AccessibilityScope(bool tagged, bool artifact, int? mcid, int? mcidStream, bool language) {
            Tagged = tagged; Artifact = artifact; Mcid = mcid; McidStream = mcidStream; Language = language;
        }
        internal bool Language { get; }
        internal bool Tagged { get; }
        internal bool Artifact { get; }
        internal int? Mcid { get; }
        internal int? McidStream { get; }
    }

    /// <summary>
    /// Walks the page content and the Form XObjects it paints, following marked-content nesting the way
    /// ISO 14289-1 7.1 reads it. <paramref name="isLinked"/> says whether an MCID is referenced from the structure tree.
    /// </summary>
    internal PdfAccessibilityPageContent ScanAccessibilityContent(Func<int?, int, bool> isLinked, CancellationToken cancellationToken) =>
        ScanAccessibility(isLinked, null, cancellationToken);

    /// <summary>
    /// Walks one annotation appearance stream and the forms it paints. Appearances are not page content, so only
    /// the font evidence is meant to be used from the result.
    /// </summary>
    internal PdfAccessibilityPageContent ScanAccessibilityAppearance(PdfStream appearance, CancellationToken cancellationToken) =>
        ScanAccessibility(static (_, _) => false, appearance, cancellationToken);

    private PdfAccessibilityPageContent ScanAccessibility(Func<int?, int, bool> isLinked, PdfStream? appearance, CancellationToken cancellationToken) {
        var result = new PdfAccessibilityPageContent();
        var budget = new PageContentBudget(this, cancellationToken);
        PdfDictionary? resources = ResolveDictionary(GetInheritedValue("Resources"));
        PdfContentSource source;
        if (appearance is null) {
            source = GetContentStreamSequence(budget).Source;
        } else {
            source = PdfContentSource.FromBytes(budget.Decode(appearance));
            resources = ResolveDictionary(appearance.Dictionary.Items.TryGetValue("Resources", out PdfObject? appearanceResources) ? appearanceResources : null) ?? resources;
        }
        // A form scanned once under the same inherited state yields the same evidence; on a repeat paint only
        // the invocation counts of the forms it paints in turn are replayed.
        var scannedForms = new Dictionary<(int Form, PdfDictionary? Resources, bool Tagged, bool Artifact, int? Mcid, int? McidStream, bool Language, PdfDictionary? Font, bool Rendered), Dictionary<int, int>>();
        var activeForms = new HashSet<int>();
        // One counter per form scan in progress, so each scan learns which forms it painted without copying page-wide counts.
        var nestedCounters = new List<Dictionary<int, int>>();
        long work = 0;
        long maxWork = Math.Max(1L, (long)_limits.MaxContentOperations) * 4L;
        ScanAccessibilityStream(source, null, resources, new AccessibilityScope(false, false, null, null, false), null, 0, 0);
        return result;

        // A Form XObject inherits the graphics state at Do, so the caller's font and rendering mode carry in.
        void ScanAccessibilityStream(PdfContentSource source, int? streamObject, PdfDictionary? streamResources,
            AccessibilityScope inherited, PdfDictionary? inheritedFont, int inheritedMode, int depth) {
            var scopes = new Stack<AccessibilityScope>();
            var states = new Stack<(PdfDictionary? Font, int Mode)>();
            PdfDictionary? font = inheritedFont;
            int mode = inheritedMode;
            bool hasPath = false;
            AccessibilityScope Current() => scopes.Count == 0 ? inherited : scopes.Peek();

            PdfContentStreamInterpreter.Interpret(source, _limits.MaxContentOperations, operation => {
                if (++work > maxWork) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.ContentOperations, maxWork, work);
                }
                if ((work & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                AccessibilityScope scope = Current();
                switch (operation.Name) {
                    case "BMC":
                    case "BDC": {
                        string? tag = operation.Operands.Count > 0 ? operation.Operands[0] as string : null;
                        PdfContentDictionary? inline = operation.Name == "BDC" && operation.Operands.Count == 2
                            ? operation.Operands[1] as PdfContentDictionary : null;
                        PdfDictionary? named = operation.Name == "BDC" && operation.Operands.Count == 2 && operation.Operands[1] is string propertyName
                            ? ResolvePropertyList(streamResources, propertyName) : null;
                        int? mcid = ReadPropertyMcid(inline, named);
                        if (mcid.HasValue && streamObject is int formWithMcid) result.FormsWithMcids.Add(formWithMcid);
                        bool artifact = string.Equals(tag, "Artifact", StringComparison.Ordinal);
                        if (artifact && scope.Tagged) result.ArtifactsInTaggedContent++;
                        else if (artifact && scope.Mcid.HasValue) result.ArtifactsInUnlinkedMcid++;
                        // Content keeps the tagged state of an enclosing MCID, so an Artifact opened inside tagged
                        // content also holds tagged content (veraPDF reports both 7.1-1 and 7.1-2 for it).
                        if ((mcid.HasValue || scope.Mcid.HasValue) && (artifact || scope.Artifact)) result.TaggedContentInArtifacts++;
                        bool linked = mcid.HasValue && isLinked(streamObject, mcid.Value);
                        string? lang = ReadPropertyText(inline, named, "Lang");
                        if (lang is not null) result.PropertyLanguages.Add(lang);
                        int? nearestMcid = mcid ?? scope.Mcid;
                        int? nearestStream = mcid.HasValue ? streamObject : scope.McidStream;
                        if (lang is null && !scope.Language) {
                            foreach (string key in PropertyTextKeys) {
                                if (ReadPropertyText(inline, named, key) is not null) {
                                    result.PropertyTextWithoutLang.Add((nearestStream, nearestMcid, key));
                                }
                            }
                        }
                        scopes.Push(new AccessibilityScope(
                            scope.Tagged || linked,
                            scope.Artifact || artifact,
                            nearestMcid,
                            nearestStream,
                            scope.Language || !string.IsNullOrEmpty(lang)));
                        break;
                    }
                    case "EMC":
                        if (scopes.Count > 0) scopes.Pop();
                        break;
                    case "q":
                        states.Push((font, mode));
                        break;
                    case "Q":
                        if (states.Count > 0) (font, mode) = states.Pop();
                        break;
                    case "Tf":
                        int? fontObject = null;
                        font = operation.Operands.Count > 0 && operation.Operands[0] is string fontName
                            ? ResolveNamedResource(streamResources, "Font", fontName, out fontObject, out _)
                            : null;
                        if (font is not null) result.FontObjectNumbers[font] = fontObject;
                        break;
                    case "gs":
                        if (operation.Operands.Count > 0 && operation.Operands[0] is string stateName &&
                            ResolveNamedResource(streamResources, "ExtGState", stateName, out _, out _) is PdfDictionary graphicsState &&
                            TryResolveGraphicsStateFont(graphicsState, streamResources, _objects, cancellationToken, out PdfDictionary? stateFont) &&
                            stateFont is not null) {
                            font = stateFont;
                            if (!result.FontObjectNumbers.ContainsKey(stateFont)) result.FontObjectNumbers[stateFont] = GraphicsStateFontObjectNumber(graphicsState, streamResources);
                        }
                        break;
                    case "Tr":
                        if (operation.Operands.Count > 0 && operation.Operands[0] is double rendering) mode = (int)rendering;
                        break;
                    case "Tj":
                    case "TJ":
                    case "'":
                    case "\"":
                        // An empty string or a TJ array of positioning numbers shows no glyphs.
                        if (!ShowsCharacters(operation)) break;
                        if (font is not null) {
                            bool rendered = mode != 3;
                            result.Fonts[font] = result.Fonts.TryGetValue(font, out bool previous) ? previous || rendered : rendered;
                        }
                        if (scope.Tagged && scope.Mcid.HasValue && !scope.Language) result.TextMcids.Add((scope.McidStream, scope.Mcid.Value));
                        else if (!scope.Tagged && !scope.Artifact && !scope.Language) result.UntaggedTextWithoutLanguage++;
                        CountContentItem(scope);
                        break;
                    case "m": case "l": case "c": case "v": case "y": case "h": case "re":
                        hasPath = true;
                        break;
                    case "n":
                        hasPath = false;
                        break;
                    case "S": case "s": case "f": case "F": case "f*": case "B": case "B*": case "b": case "b*":
                        // A painting operator with no current path paints nothing.
                        if (hasPath) CountContentItem(scope);
                        hasPath = false;
                        break;
                    case "sh": case "BI":
                        CountContentItem(scope);
                        break;
                    case "Do":
                        if (operation.Operands.Count > 0 && operation.Operands[0] is string xobjectName) {
                            PaintXObject(xobjectName, streamResources, scope, font, mode, depth);
                        }
                        break;
                }
            },
                inlineImageComponentCount: name => GetDeclaredColorSpaceComponentCount(streamResources, name),
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxGraphicsStateDepth: _limits.MaxGraphicsStateDepth,
                maxOperands: _limits.MaxContentOperands,
                operandBuffer: new PdfContentOperands());
        }

        void CountContentItem(AccessibilityScope scope) {
            if (!scope.Tagged && !scope.Artifact) result.UntaggedItems++;
        }

        void PaintXObject(string name, PdfDictionary? streamResources, AccessibilityScope scope, PdfDictionary? font, int mode, int depth) {
            PdfDictionary? xobjects = streamResources is null ? null
                : ResolveDictionary(streamResources.Items.TryGetValue("XObject", out PdfObject? xobjectsObject) ? xobjectsObject : null);
            if (xobjects is null || !xobjects.Items.TryGetValue(name, out PdfObject? entry)) return;
            int? objectNumber = entry is PdfReference reference ? reference.ObjectNumber : null;
            if (ResolveObject(entry) is not PdfStream stream) return;
            string? subtype = stream.Dictionary.Get<PdfName>("Subtype")?.Name;
            if (string.Equals(subtype, "Image", StringComparison.Ordinal)) {
                CountContentItem(scope);
                return;
            }
            if (!string.Equals(subtype, "Form", StringComparison.Ordinal)) return;
            // A form without its own Resources uses the painting stream's, so those are part of what it shows.
            PdfDictionary? formResources = ResolveDictionary(stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject) ? resourcesObject : null)
                ?? streamResources;
            var key = (objectNumber ?? 0, formResources, scope.Tagged, scope.Artifact, scope.Mcid, scope.McidStream, scope.Language, font, mode != 3);
            if (objectNumber is int formNumber) {
                CountInvocation(formNumber, 1);
                if (stream.Dictionary.Items.ContainsKey("Ref")) result.ReferenceXObjects.Add(formNumber);
                if (activeForms.Contains(formNumber)) return;
                if (scannedForms.TryGetValue(key, out Dictionary<int, int>? painted)) {
                    foreach (KeyValuePair<int, int> invocation in painted) CountInvocation(invocation.Key, invocation.Value);
                    return;
                }
                activeForms.Add(formNumber);
            }
            if (depth + 1 > _limits.MaxContentNestingDepth) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ContentNestingDepth, _limits.MaxContentNestingDepth, depth + 1);
            }
            var nested = new Dictionary<int, int>();
            nestedCounters.Add(nested);
            try {
                byte[] decoded = budget.Decode(stream);
                ScanAccessibilityStream(PdfContentSource.FromBytes(decoded), objectNumber, formResources, scope, font, mode, depth + 1);
                if (objectNumber is not null) scannedForms[key] = nested;
            } finally {
                nestedCounters.RemoveAt(nestedCounters.Count - 1);
                if (objectNumber is int number) activeForms.Remove(number);
            }
        }

        void CountInvocation(int form, int count) {
            result.FormInvocations[form] = PdfAccessibilityPageContent.AddCount(result.FormInvocations.TryGetValue(form, out int seen) ? seen : 0, count);
            foreach (Dictionary<int, int> counter in nestedCounters) {
                counter[form] = PdfAccessibilityPageContent.AddCount(counter.TryGetValue(form, out int nestedSeen) ? nestedSeen : 0, count);
            }
        }
    }

    private static readonly string[] PropertyTextKeys = { "Alt", "ActualText", "E" };

    // Operands that can't be read are treated as showing text, so malformed content is never skipped.
    private static bool ShowsCharacters(PdfContentOperation operation) {
        if (operation.Operands.Count == 0) return true;
        object last = operation.Operands[operation.Operands.Count - 1];
        if (last is byte[] text) return text.Length > 0;
        // The interpreter reads an array of numbers only, such as [120], as double[].
        if (last is double[] && string.Equals(operation.Name, "TJ", StringComparison.Ordinal)) return false;
        if (last is List<object> items && string.Equals(operation.Name, "TJ", StringComparison.Ordinal)) {
            foreach (object item in items) {
                if (item is byte[] part ? part.Length > 0 : item is not double) return true;
            }
            return false;
        }
        return true;
    }

    private int? GraphicsStateFontObjectNumber(PdfDictionary graphicsState, PdfDictionary? resources) {
        if (ResolveObject(graphicsState.Items.TryGetValue("Font", out PdfObject? selection) ? selection : null) is not PdfArray values || values.Items.Count == 0) return null;
        PdfObject first = values.Items[0];
        if (first is PdfReference reference) return reference.ObjectNumber;
        if (ResolveObject(first) is PdfName name) {
            ResolveNamedResource(resources, "Font", name.Name, out int? objectNumber, out _);
            return objectNumber;
        }
        return null;
    }

    private PdfDictionary? ResolvePropertyList(PdfDictionary? resources, string name) {
        PdfDictionary? properties = resources is null ? null
            : ResolveDictionary(resources.Items.TryGetValue("Properties", out PdfObject? propertiesObject) ? propertiesObject : null);
        return properties is not null && properties.Items.TryGetValue(name, out PdfObject? property) ? ResolveDictionary(property) : null;
    }

    private PdfDictionary? ResolveNamedResource(PdfDictionary? resources, string category, string name, out int? objectNumber, out PdfObject? entry) {
        objectNumber = null;
        entry = null;
        PdfDictionary? group = resources is null ? null
            : ResolveDictionary(resources.Items.TryGetValue(category, out PdfObject? groupObject) ? groupObject : null);
        if (group is null || !group.Items.TryGetValue(name, out entry)) return null;
        if (entry is PdfReference reference) objectNumber = reference.ObjectNumber;
        return ResolveDictionary(entry);
    }

    private int? ReadPropertyMcid(PdfContentDictionary? inline, PdfDictionary? named) {
        if (inline is not null) {
            return inline.Items.TryGetValue("MCID", out object? value) && value is double number &&
                number >= 0D && number <= int.MaxValue && Math.Truncate(number) == number ? (int)number : null;
        }
        return named is null ? null : TryReadInteger(named, "MCID");
    }

    private string? ReadPropertyText(PdfContentDictionary? inline, PdfDictionary? named, string key) {
        if (inline is not null) {
            if (!inline.Items.TryGetValue(key, out object? value)) return null;
            return value switch {
                string name => name,
                byte[] bytes => PdfTextString.Decode(bytes),
                _ => string.Empty
            };
        }
        if (named is null || !named.Items.TryGetValue(key, out PdfObject? namedValue)) return null;
        return ResolveObject(namedValue) is PdfStringObj namedText ? namedText.Value : string.Empty;
    }
}
