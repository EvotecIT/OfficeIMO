using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfPrintProductionStructureInspector {
    private static ReachableFontInspection InspectReachableFonts(
        PdfReadDocument document,
        System.Threading.CancellationToken cancellationToken) {
        var collector = new ReachableFontCollector(document, cancellationToken);
        return collector.Inspect();
    }

    private sealed class ReachableFontCollector {
        private readonly PdfReadDocument _document;
        private readonly Dictionary<int, PdfIndirectObject> _objects;
        private readonly PdfReadLimits _limits;
        private readonly System.Threading.CancellationToken _cancellationToken;
        private readonly HashSet<PdfDictionary> _fonts = new HashSet<PdfDictionary>();
        private readonly List<ContentContext> _pending = new List<ContentContext>();
        private readonly List<ContentContext> _visited = new List<ContentContext>();
        private int _uninspectableContextCount;

        internal ReachableFontCollector(
            PdfReadDocument document,
            System.Threading.CancellationToken cancellationToken) {
            _document = document;
            _objects = document.Objects;
            _limits = document.ReadOptions.Limits;
            _cancellationToken = cancellationToken;
        }

        internal ReachableFontInspection Inspect() {
            foreach (PdfReadPage page in _document.Pages) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (!_objects.TryGetValue(page.ObjectNumber, out PdfIndirectObject? pageObject) ||
                    pageObject == null ||
                    pageObject.Value is not PdfDictionary pageDictionary) {
                    _uninspectableContextCount++;
                    continue;
                }

                PdfDictionary? resources = ResolveInheritedResources(pageDictionary);
                if (pageDictionary.Items.TryGetValue("Contents", out PdfObject? contents)) {
                    var pageFontState = new PageFontState();
                    if (!AddContentObject(contents, resources, contentDepth: 0, pageFontState: pageFontState)) {
                        _uninspectableContextCount++;
                    }
                }
                if (pageDictionary.Items.TryGetValue("Annots", out PdfObject? annotations)) {
                    AddAnnotationAppearances(annotations, resources, objectDepth: 0, new HashSet<PdfObject>());
                }
            }

            while (_pending.Count > 0) {
                _cancellationToken.ThrowIfCancellationRequested();
                ContentContext context = _pending[0];
                _pending.RemoveAt(0);
                if (ContainsContext(_visited, context.Stream, context.Resources, context.SelectedFontObject, context.PageFontState)) continue;
                _visited.Add(context);
                InspectContent(context);
            }

            return new ReachableFontInspection(_fonts, _uninspectableContextCount);
        }

        private void InspectContent(ContentContext context) {
            if (context.ContentDepth > _limits.MaxContentNestingDepth) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.ContentNestingDepth,
                    _limits.MaxContentNestingDepth,
                    context.ContentDepth);
            }
            if (!StreamDecoder.TryDecode(
                    context.Stream.Dictionary,
                    context.Stream.Data,
                    _limits.MaxDecodedStreamBytes,
                    out byte[] decoded,
                    _objects)) {
                _uninspectableContextCount++;
                return;
            }

            bool contextWasUninspectable = false;
            PdfObject? activeFontObject = context.PageFontState?.SelectedFontObject ?? context.SelectedFontObject;
            PdfContentStreamInterpreter.Interpret(
                PdfEncoding.Latin1GetString(decoded),
                _limits.MaxContentOperations,
                operation => {
                    _cancellationToken.ThrowIfCancellationRequested();
                    if (operation.HasInvalidOperands) {
                        if (IsFontRelevantOperator(operation.Name)) contextWasUninspectable = true;
                        return;
                    }
                    switch (operation.Name) {
                        case "Tf" when operation.Operands.Count == 2 && operation.Operands[0] is string fontName:
                            if (!TryResolveResource(context.Resources, "Font", fontName, out PdfObject? fontObject) ||
                                ResolveObject(_objects, fontObject, 0, _limits.MaxObjectNestingDepth, out _) is not PdfDictionary font) {
                                contextWasUninspectable = true;
                                break;
                            }
                            activeFontObject = fontObject;
                            AddSelectedFont(font);
                            break;
                        case "Tf":
                            contextWasUninspectable = true;
                            break;
                        case "Do" when operation.Operands.Count == 1 && operation.Operands[0] is string xObjectName:
                            if (!TryResolveResource(context.Resources, "XObject", xObjectName, out PdfObject? xObject) ||
                                ResolveObject(_objects, xObject, 0, _limits.MaxObjectNestingDepth, out _) is not PdfStream form) {
                                contextWasUninspectable = true;
                                break;
                            }
                            if (string.Equals(ResolveName(
                                    form.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtype) ? subtype : null,
                                    _objects,
                                    _limits.MaxObjectNestingDepth),
                                "Form",
                                StringComparison.Ordinal)) {
                                AddStream(
                                    form,
                                    ResolveStreamResources(form, context.Resources),
                                    context.ContentDepth + 1,
                                    activeFontObject);
                            }
                            break;
                        case "Do":
                            contextWasUninspectable = true;
                            break;
                        case "gs" when operation.Operands.Count == 1 && operation.Operands[0] is string graphicsStateName:
                            if (!TryResolveResource(context.Resources, "ExtGState", graphicsStateName, out PdfObject? graphicsStateObject)) break;
                            AddSoftMaskContent(graphicsStateObject!, context.Resources, context.ContentDepth + 1, activeFontObject);
                            break;
                        case "scn":
                        case "SCN":
                            if (operation.Operands.Count > 0 && operation.Operands[operation.Operands.Count - 1] is string patternName) {
                                AddPatternContent(patternName, context.Resources, context.ContentDepth + 1, activeFontObject);
                            }
                            break;
                        case "Tj":
                        case "TJ":
                        case "'":
                        case "\"":
                            if (activeFontObject == null ||
                                !TryAddShownType3CharProcs(activeFontObject, operation, context, context.ContentDepth + 1)) {
                                contextWasUninspectable = true;
                            }
                            break;
                    }
                },
                maxNestingDepth: _limits.MaxContentNestingDepth,
                maxOperands: _limits.MaxContentOperands,
                dispatchInvalidOperations: true);
            if (context.PageFontState != null) context.PageFontState.SelectedFontObject = activeFontObject;
            if (contextWasUninspectable) _uninspectableContextCount++;
        }

        private void AddSelectedFont(PdfDictionary font) {
            _fonts.Add(font);
        }

        private void AddPatternContent(string name, PdfDictionary? resources, int contentDepth, PdfObject? selectedFontObject) {
            if (!TryResolveResource(resources, "Pattern", name, out PdfObject? patternObject) ||
                ResolveObject(_objects, patternObject, 0, _limits.MaxObjectNestingDepth, out _) is not PdfStream pattern) return;
            PdfObject? patternType = ResolveObject(
                _objects,
                pattern.Dictionary.Items.TryGetValue("PatternType", out PdfObject? value) ? value : null,
                0,
                _limits.MaxObjectNestingDepth,
                out _);
            if (patternType is PdfNumber { Value: 1D }) {
                AddStream(pattern, ResolveStreamResources(pattern, resources), contentDepth, selectedFontObject);
            }
        }

        private void AddSoftMaskContent(PdfObject graphicsStateObject, PdfDictionary? resources, int contentDepth, PdfObject? selectedFontObject) {
            if (ResolveObject(_objects, graphicsStateObject, 0, _limits.MaxObjectNestingDepth, out _) is not PdfDictionary graphicsState ||
                !graphicsState.Items.TryGetValue("SMask", out PdfObject? softMaskObject) ||
                ResolveObject(_objects, softMaskObject, 0, _limits.MaxObjectNestingDepth, out _) is not PdfDictionary softMask ||
                !softMask.Items.TryGetValue("G", out PdfObject? groupObject) ||
                ResolveObject(_objects, groupObject, 0, _limits.MaxObjectNestingDepth, out _) is not PdfStream group) return;
            AddStream(group, ResolveStreamResources(group, resources), contentDepth, selectedFontObject);
        }

        private bool TryAddShownType3CharProcs(
            PdfObject fontObject,
            PdfContentOperation operation,
            ContentContext context,
            int contentDepth) {
            if (ResolveObject(_objects, fontObject, 0, _limits.MaxObjectNestingDepth, out int fontDepth) is not PdfDictionary font) return false;
            AddSelectedFont(font);
            string? subtype = ResolveName(
                font.Items.TryGetValue("Subtype", out PdfObject? subtypeObject) ? subtypeObject : null,
                _objects,
                _limits.MaxObjectNestingDepth);
            if (!string.Equals(subtype, "Type3", StringComparison.Ordinal)) return true;
            if (!font.Items.TryGetValue("CharProcs", out PdfObject? charProcsObject) ||
                ResolveObject(_objects, charProcsObject, fontDepth + 1, _limits.MaxObjectNestingDepth, out _) is not PdfDictionary charProcs ||
                !PdfPrintProductionColorInspector.TryGetType3GlyphNames(font, _objects, _limits.MaxObjectNestingDepth, out Dictionary<int, string> glyphNames) ||
                !PdfPrintProductionColorInspector.TryGetShownTextBytes(operation, out List<byte[]> shownText)) return false;

            PdfDictionary? resources = ResolveDirectResources(font) ?? context.Resources;
            var shownGlyphNames = new HashSet<string>(StringComparer.Ordinal);
            foreach (byte[] text in shownText) {
                for (int index = 0; index < text.Length; index++) {
                    if (!glyphNames.TryGetValue(text[index], out string? glyphName)) return false;
                    shownGlyphNames.Add(glyphName);
                }
            }
            foreach (string glyphName in shownGlyphNames) {
                if (!charProcs.Items.TryGetValue(glyphName, out PdfObject? charProc) ||
                    !AddContentObject(charProc, resources, contentDepth)) return false;
            }
            return true;
        }

        private void AddAnnotationAppearances(
            PdfObject value,
            PdfDictionary? pageResources,
            int objectDepth,
            HashSet<PdfObject> visited) {
            _cancellationToken.ThrowIfCancellationRequested();
            PdfObject? resolved = ResolveObject(_objects, value, objectDepth, _limits.MaxObjectNestingDepth, out int resolvedDepth);
            if (resolved == null || !visited.Add(resolved)) return;
            if (resolved is PdfArray array) {
                foreach (PdfObject item in array.Items) AddAnnotationAppearances(item, pageResources, resolvedDepth + 1, visited);
                return;
            }
            if (resolved is PdfStream appearance) {
                AddStream(appearance, ResolveStreamResources(appearance, pageResources), contentDepth: 1);
                return;
            }
            if (resolved is not PdfDictionary dictionary) return;
            if (dictionary.Items.TryGetValue("AP", out PdfObject? appearances)) {
                AddAppearanceObject(appearances, pageResources, resolvedDepth + 1, visited);
            }
        }

        private void AddAppearanceObject(
            PdfObject value,
            PdfDictionary? pageResources,
            int objectDepth,
            HashSet<PdfObject> visited) {
            _cancellationToken.ThrowIfCancellationRequested();
            PdfObject? resolved = ResolveObject(_objects, value, objectDepth, _limits.MaxObjectNestingDepth, out int resolvedDepth);
            if (resolved == null || !visited.Add(resolved)) return;
            if (resolved is PdfStream appearance) {
                AddStream(appearance, ResolveStreamResources(appearance, pageResources), contentDepth: 1);
            } else if (resolved is PdfDictionary dictionary) {
                foreach (PdfObject child in dictionary.Items.Values) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    AddAppearanceObject(child, pageResources, resolvedDepth + 1, visited);
                }
            }
        }

        private bool AddContentObject(
            PdfObject value,
            PdfDictionary? resources,
            int contentDepth,
            PdfObject? selectedFontObject = null,
            PageFontState? pageFontState = null) {
            bool complete = true;
            var pending = new Stack<(PdfObject Value, int Depth)>();
            var visitedArrays = new HashSet<PdfArray>();
            pending.Push((value, 0));
            while (pending.Count > 0) {
                _cancellationToken.ThrowIfCancellationRequested();
                (PdfObject candidate, int depth) = pending.Pop();
                PdfObject? resolved = ResolveObject(_objects, candidate, depth, _limits.MaxObjectNestingDepth, out int resolvedDepth);
                if (resolved is PdfStream stream) {
                    AddStream(stream, resources, contentDepth, selectedFontObject, pageFontState);
                } else if (resolved is PdfArray array && visitedArrays.Add(array)) {
                    for (int index = array.Items.Count - 1; index >= 0; index--) {
                        pending.Push((array.Items[index], resolvedDepth + 1));
                    }
                } else {
                    complete = false;
                }
            }
            return complete;
        }

        private void AddStream(
            PdfStream stream,
            PdfDictionary? resources,
            int contentDepth,
            PdfObject? selectedFontObject = null,
            PageFontState? pageFontState = null) {
            if (ContainsContext(_visited, stream, resources, selectedFontObject, pageFontState) ||
                ContainsContext(_pending, stream, resources, selectedFontObject, pageFontState)) return;
            _pending.Add(new ContentContext(stream, resources, contentDepth, selectedFontObject, pageFontState));
        }

        private PdfDictionary? ResolveInheritedResources(PdfDictionary page) {
            var visited = new HashSet<PdfDictionary>();
            PdfDictionary? current = page;
            int depth = 0;
            while (current != null && visited.Add(current)) {
                if (current.Items.TryGetValue("Resources", out PdfObject? resourcesObject) &&
                    ResolveObject(_objects, resourcesObject, depth + 1, _limits.MaxObjectNestingDepth, out _) is PdfDictionary resources) {
                    return resources;
                }
                if (!current.Items.TryGetValue("Parent", out PdfObject? parentObject) ||
                    ResolveObject(_objects, parentObject, depth + 1, _limits.MaxObjectNestingDepth, out int parentDepth) is not PdfDictionary parent) break;
                current = parent;
                depth = parentDepth;
            }
            return null;
        }

        private PdfDictionary? ResolveStreamResources(PdfStream stream, PdfDictionary? inheritedResources) =>
            ResolveDirectResources(stream.Dictionary) ?? inheritedResources;

        private PdfDictionary? ResolveDirectResources(PdfDictionary owner) =>
            owner.Items.TryGetValue("Resources", out PdfObject? resourcesObject) &&
            ResolveObject(_objects, resourcesObject, 0, _limits.MaxObjectNestingDepth, out _) is PdfDictionary resources
                ? resources
                : null;

        private bool TryResolveResource(
            PdfDictionary? resources,
            string category,
            string name,
            out PdfObject? resource) {
            resource = null;
            if (resources == null ||
                !resources.Items.TryGetValue(category, out PdfObject? categoryObject) ||
                ResolveObject(_objects, categoryObject, 0, _limits.MaxObjectNestingDepth, out _) is not PdfDictionary entries) return false;
            return entries.Items.TryGetValue(name, out resource);
        }

        private static bool ContainsContext(
            IReadOnlyList<ContentContext> contexts,
            PdfStream stream,
            PdfDictionary? resources,
            PdfObject? selectedFontObject,
            PageFontState? pageFontState) {
            for (int index = 0; index < contexts.Count; index++) {
                if (ReferenceEquals(contexts[index].Stream, stream) &&
                    ReferenceEquals(contexts[index].Resources, resources) &&
                    ReferenceEquals(contexts[index].SelectedFontObject, selectedFontObject) &&
                    ReferenceEquals(contexts[index].PageFontState, pageFontState)) return true;
            }
            return false;
        }

        private static bool IsFontRelevantOperator(string name) =>
            string.Equals(name, "Tf", StringComparison.Ordinal) ||
            string.Equals(name, "Do", StringComparison.Ordinal) ||
            string.Equals(name, "gs", StringComparison.Ordinal) ||
            string.Equals(name, "scn", StringComparison.Ordinal) ||
            string.Equals(name, "SCN", StringComparison.Ordinal) ||
            string.Equals(name, "Tj", StringComparison.Ordinal) ||
            string.Equals(name, "TJ", StringComparison.Ordinal) ||
            string.Equals(name, "'", StringComparison.Ordinal) ||
            string.Equals(name, "\"", StringComparison.Ordinal);

    }

    private sealed class PageFontState {
        internal PdfObject? SelectedFontObject { get; set; }
    }

    private sealed record ContentContext(
        PdfStream Stream,
        PdfDictionary? Resources,
        int ContentDepth,
        PdfObject? SelectedFontObject,
        PageFontState? PageFontState);
    private sealed record ReachableFontInspection(HashSet<PdfDictionary> Fonts, int UninspectableContextCount);
}
