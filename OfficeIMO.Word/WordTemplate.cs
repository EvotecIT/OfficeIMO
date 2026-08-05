using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    /// <summary>
    /// Binds ordinary <c>{{Name}}</c> placeholders, conditional blocks, and repeated blocks in a Word document.
    /// </summary>
    public static class WordTemplate {
        private static readonly Regex PlaceholderRegex = new Regex(
            @"\{\{\s*(?<name>[A-Za-z0-9_.-]+)\s*\}\}",
            RegexOptions.Compiled,
            TimeSpan.FromMilliseconds(100));

        private static readonly Regex ConditionalMarkerRegex = new Regex(
            @"^\s*\{\{\s*(?<kind>[#/])\s*(?<name>[A-Za-z0-9_.-]+)\s*\}\}\s*$",
            RegexOptions.Compiled,
            TimeSpan.FromMilliseconds(100));

        private static readonly Regex RepeatingMarkerRegex = new Regex(
            @"^\s*\{\{\s*(?<kind>#each|/each)\s+(?<name>[A-Za-z0-9_.-]+)\s*\}\}\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            TimeSpan.FromMilliseconds(100));

        private static readonly ConcurrentDictionary<Type, IReadOnlyDictionary<string, PropertyInfo>> PropertyCache =
            new ConcurrentDictionary<Type, IReadOnlyDictionary<string, PropertyInfo>>();

        /// <summary>
        /// Applies AOT-safe dictionary entries to a Word document. Nested objects must also be dictionaries; arbitrary objects are not reflected.
        /// </summary>
        public static WordTemplateResult Apply(
            WordDocument document,
            IEnumerable<KeyValuePair<string, object?>> values,
            WordTemplateOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (values == null) throw new ArgumentNullException(nameof(values));

            var model = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, object?> pair in values) {
                if (model.ContainsKey(pair.Key)) {
                    throw new ArgumentException("Template dictionary keys must be unique ignoring case.", nameof(values));
                }
                model.Add(pair.Key, pair.Value);
            }

            return ApplyCore(document, model, allowReflection: false, options);
        }

        /// <summary>
        /// Applies a public-property object model to a Word document. Use the dictionary overload for trimming and NativeAOT-safe binding.
        /// </summary>
        [RequiresUnreferencedCode("Uses reflection over the supplied object graph. For trimming and NativeAOT, pass dictionary entries instead.")]
        public static WordTemplateResult Apply(
            WordDocument document,
            object model,
            WordTemplateOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (model == null) throw new ArgumentNullException(nameof(model));

            return ApplyCore(document, model, allowReflection: true, options);
        }

        private static WordTemplateResult ApplyCore(
            WordDocument document,
            object model,
            bool allowReflection,
            WordTemplateOptions? options) {
            var state = new BindingState(document, options ?? new WordTemplateOptions());
            var scope = new TemplateScope(model, parent: null, allowReflection);

            foreach (OpenXmlCompositeElement root in WordMailMerge.EnumerateTemplateRoots(document)) {
                ProcessContainer(root, scope, state);
            }

            return new WordTemplateResult(
                state.PlaceholderCount,
                state.ReplacedPlaceholderCount,
                state.RepeatedBlockCount,
                state.ConditionalBlockCount,
                state.MissingValueNames);
        }

        private static void ProcessContainer(OpenXmlCompositeElement container, TemplateScope scope, BindingState state) {
            ProcessDirectBlocks(container, scope, state);

            foreach (OpenXmlElement child in container.ChildElements.ToList()) {
                if (child is Paragraph paragraph) {
                    ReplacePlaceholders(paragraph, scope, state);
                    foreach (OpenXmlCompositeElement nested in paragraph.ChildElements.OfType<OpenXmlCompositeElement>()) {
                        ProcessContainer(nested, scope, state);
                    }
                } else if (child is OpenXmlCompositeElement composite) {
                    ProcessContainer(composite, scope, state);
                }
            }
        }

        private static void ProcessDirectBlocks(OpenXmlCompositeElement container, TemplateScope scope, BindingState state) {
            List<OpenXmlElement> elements = container.ChildElements.ToList();
            var stack = new List<BlockStart>();
            var blocks = new List<BlockRange>();

            for (int index = 0; index < elements.Count; index++) {
                if (elements[index] is not Paragraph paragraph || !TryGetBlockMarker(paragraph, out BlockMarker marker)) {
                    continue;
                }

                if (marker.IsStart) {
                    stack.Add(new BlockStart(marker.Kind, marker.Name, index));
                    continue;
                }

                if (stack.Count == 0) {
                    throw new InvalidOperationException("Template block end marker '" + paragraph.InnerText + "' has no matching start marker.");
                }

                BlockStart start = stack[stack.Count - 1];
                stack.RemoveAt(stack.Count - 1);
                if (start.Kind != marker.Kind || !string.Equals(start.Name, marker.Name, StringComparison.OrdinalIgnoreCase)) {
                    throw new InvalidOperationException("Template block end marker '" + paragraph.InnerText + "' does not match the open '" + start.Name + "' block.");
                }

                if (stack.Count == 0) {
                    blocks.Add(new BlockRange(start.Kind, start.Name, start.Index, index));
                }
            }

            if (stack.Count > 0) {
                BlockStart start = stack[stack.Count - 1];
                throw new InvalidOperationException("Template block '" + start.Name + "' has no matching end marker.");
            }

            bool nestedBlocksExposed = false;
            foreach (BlockRange block in blocks.OrderByDescending(static item => item.StartIndex)) {
                if (block.Kind == BlockKind.Repeating) {
                    ExpandRepeatingBlock(elements, block, scope, state);
                } else {
                    nestedBlocksExposed |= EvaluateConditionalBlock(elements, block, scope, state);
                }
            }

            if (nestedBlocksExposed) ProcessDirectBlocks(container, scope, state);
        }

        private static void ExpandRepeatingBlock(
            IReadOnlyList<OpenXmlElement> elements,
            BlockRange block,
            TemplateScope scope,
            BindingState state) {
            if (!scope.TryResolve(block.Name, out object? value)) {
                throw new InvalidOperationException("Repeating block '" + block.Name + "' was not supplied.");
            }
            if (!TryGetEnumerable(value, out IEnumerable? items)) {
                throw new InvalidOperationException("Repeating block '" + block.Name + "' requires an enumerable value.");
            }

            List<OpenXmlElement> templateElements = elements
                .Skip(block.StartIndex + 1)
                .Take(block.EndIndex - block.StartIndex - 1)
                .ToList();
            OpenXmlElement insertionPoint = elements[block.StartIndex];
            state.ReleaseRepeatedTemplateBookmarkNames(templateElements);

            foreach (object? item in items!) {
                var working = new SdtContentBlock();
                foreach (OpenXmlElement templateElement in templateElements) {
                    working.Append(templateElement.CloneNode(true));
                }

                state.NormalizeRepeatedClone(working);
                var itemScope = new TemplateScope(item, scope, scope.AllowReflection);
                ProcessContainer(working, itemScope, state);
                foreach (OpenXmlElement generated in working.ChildElements.ToList()) {
                    generated.Remove();
                    insertionPoint.InsertBeforeSelf(generated);
                }
                state.RepeatedBlockCount++;
            }

            RemoveRange(elements, block.StartIndex, block.EndIndex);
        }

        private static bool EvaluateConditionalBlock(
            IReadOnlyList<OpenXmlElement> elements,
            BlockRange block,
            TemplateScope scope,
            BindingState state) {
            if (!scope.TryResolve(block.Name, out object? value) || value is not bool include) {
                throw new InvalidOperationException("Conditional block '" + block.Name + "' requires a Boolean value.");
            }

            state.ConditionalBlockCount++;
            if (!include) {
                RemoveRange(elements, block.StartIndex, block.EndIndex);
                return false;
            }

            RemoveIfAttached(elements[block.EndIndex]);
            RemoveIfAttached(elements[block.StartIndex]);
            return true;
        }

        private static void ReplacePlaceholders(Paragraph paragraph, TemplateScope scope, BindingState state) {
            List<TextSlice> slices = BuildTextSlices(paragraph);
            if (slices.Count == 0) return;

            string text = string.Concat(slices.Select(static slice => slice.Text.Text));
            MatchCollection matches = PlaceholderRegex.Matches(text);
            state.PlaceholderCount += matches.Count;

            for (int index = matches.Count - 1; index >= 0; index--) {
                Match match = matches[index];
                string name = match.Groups["name"].Value;
                if (!scope.TryResolve(name, out object? value)) {
                    state.MissingValueNames.Add(name);
                    if (state.Options.RemoveMissingPlaceholders) {
                        ReplaceWithText(slices, match.Index, match.Length, string.Empty);
                    }
                    continue;
                }

                if (value is WordTemplateImage image) {
                    ReplaceWithImage(paragraph, slices, match.Index, match.Length, image, state.Document);
                } else if (value is WordTemplateHyperlink hyperlink) {
                    ReplaceWithHyperlink(paragraph, slices, match.Index, match.Length, hyperlink, state.Document);
                } else if (IsScalar(value)) {
                    ReplaceWithText(slices, match.Index, match.Length, FormatScalar(value, state.Options.Culture));
                } else {
                    throw new InvalidOperationException("Placeholder '" + name + "' requires a scalar, WordTemplateImage, or WordTemplateHyperlink value.");
                }

                state.ReplacedPlaceholderCount++;
            }
        }

        private static List<TextSlice> BuildTextSlices(Paragraph paragraph) {
            var slices = new List<TextSlice>();
            int offset = 0;
            foreach (Text text in paragraph.Descendants<Text>()
                         .Where(candidate => ReferenceEquals(candidate.Ancestors<Paragraph>().FirstOrDefault(), paragraph))) {
                string value = text.Text ?? string.Empty;
                if (value.Length == 0) continue;
                slices.Add(new TextSlice(text, offset, offset + value.Length));
                offset += value.Length;
            }
            return slices;
        }

        private static void ReplaceWithText(IReadOnlyList<TextSlice> slices, int start, int length, string replacement) {
            LocateSlices(slices, start, length, out TextSlice first, out TextSlice last);
            int firstOffset = start - first.Start;
            int lastOffset = start + length - last.Start;
            string before = first.Text.Text.Substring(0, firstOffset);
            string after = last.Text.Text.Substring(lastOffset);

            if (ReferenceEquals(first.Text, last.Text)) {
                SetText(first.Text, before + replacement + after);
                return;
            }

            SetText(first.Text, before + replacement);
            bool between = false;
            foreach (TextSlice slice in slices) {
                if (ReferenceEquals(slice.Text, first.Text)) {
                    between = true;
                    continue;
                }
                if (!between) continue;
                if (ReferenceEquals(slice.Text, last.Text)) {
                    SetText(slice.Text, after);
                    break;
                }
                SetText(slice.Text, string.Empty);
            }
        }

        private static void ReplaceWithImage(
            Paragraph paragraph,
            IReadOnlyList<TextSlice> slices,
            int start,
            int length,
            WordTemplateImage image,
            WordDocument document) {
            Run anchor = PrepareRichReplacement(paragraph, slices, start, length, out Run? suffixRun);
            var imageRun = new Run();
            var wrapper = new WordParagraph(document, paragraph, imageRun);
            using (var stream = new MemoryStream(image.GetContentUnsafe(), writable: false)) {
                var wordImage = new WordImage(document, wrapper, stream, image.FileName, image.Width, image.Height, description: image.Description);
                imageRun.Append(wordImage._Image);
            }
            anchor.InsertAfterSelf(imageRun);
            if (suffixRun != null) imageRun.InsertAfterSelf(suffixRun);
        }

        private static void ReplaceWithHyperlink(
            Paragraph paragraph,
            IReadOnlyList<TextSlice> slices,
            int start,
            int length,
            WordTemplateHyperlink hyperlink,
            WordDocument document) {
            Run anchor = PrepareRichReplacement(paragraph, slices, start, length, out Run? suffixRun);
            var wrapper = new WordParagraph(document, paragraph);
            WordHyperLink.AddHyperLink(wrapper, hyperlink.Text, hyperlink.Uri, hyperlink.AddStyle, hyperlink.Tooltip);
            Hyperlink generated = wrapper._hyperlink
                ?? throw new InvalidOperationException("The template hyperlink could not be created.");
            generated.Remove();
            anchor.InsertAfterSelf(generated);
            if (suffixRun != null) generated.InsertAfterSelf(suffixRun);
        }

        private static Run PrepareRichReplacement(
            Paragraph paragraph,
            IReadOnlyList<TextSlice> slices,
            int start,
            int length,
            out Run? suffixRun) {
            LocateSlices(slices, start, length, out TextSlice first, out TextSlice last);
            Run firstRun = first.Text.Ancestors<Run>().FirstOrDefault()
                ?? throw new InvalidOperationException("Template placeholders must be stored in Word text runs.");
            Run lastRun = last.Text.Ancestors<Run>().FirstOrDefault()
                ?? throw new InvalidOperationException("Template placeholders must be stored in Word text runs.");
            if (!ReferenceEquals(firstRun.Parent, paragraph) || !ReferenceEquals(lastRun.Parent, paragraph)) {
                throw new InvalidOperationException("Image and hyperlink placeholders must be stored in direct paragraph text runs.");
            }
            int firstOffset = start - first.Start;
            int lastOffset = start + length - last.Start;
            string before = first.Text.Text.Substring(0, firstOffset);
            string after = last.Text.Text.Substring(lastOffset);
            SetText(first.Text, before);

            suffixRun = null;
            if (ReferenceEquals(first.Text, last.Text)) {
                if (after.Length > 0) {
                    suffixRun = new Run();
                    if (firstRun.RunProperties != null) {
                        suffixRun.RunProperties = (RunProperties)firstRun.RunProperties.CloneNode(true);
                    }
                    suffixRun.Append(new Text(after) { Space = SpaceProcessingModeValues.Preserve });
                }
                return firstRun;
            }

            bool between = false;
            foreach (TextSlice slice in slices) {
                if (ReferenceEquals(slice.Text, first.Text)) {
                    between = true;
                    continue;
                }
                if (!between) continue;
                if (ReferenceEquals(slice.Text, last.Text)) {
                    SetText(slice.Text, after);
                    break;
                }
                SetText(slice.Text, string.Empty);
            }

            if (ReferenceEquals(firstRun, lastRun) && after.Length > 0) {
                suffixRun = new Run();
                if (firstRun.RunProperties != null) {
                    suffixRun.RunProperties = (RunProperties)firstRun.RunProperties.CloneNode(true);
                }
                suffixRun.Append(new Text(after) { Space = SpaceProcessingModeValues.Preserve });
                SetText(last.Text, string.Empty);
            }
            return firstRun;
        }

        private static void LocateSlices(
            IReadOnlyList<TextSlice> slices,
            int start,
            int length,
            out TextSlice first,
            out TextSlice last) {
            int end = start + length;
            first = slices.First(slice => slice.Start <= start && start < slice.End);
            last = slices.First(slice => slice.Start < end && end <= slice.End);
        }

        private static void SetText(Text text, string value) {
            text.Text = value;
            text.Space = SpaceProcessingModeValues.Preserve;
        }

        private static bool TryGetBlockMarker(Paragraph paragraph, out BlockMarker marker) {
            string text = paragraph.InnerText ?? string.Empty;
            Match repeating = RepeatingMarkerRegex.Match(text);
            if (repeating.Success) {
                marker = new BlockMarker(
                    BlockKind.Repeating,
                    repeating.Groups["name"].Value,
                    repeating.Groups["kind"].Value.StartsWith("#", StringComparison.Ordinal));
                return true;
            }

            Match conditional = ConditionalMarkerRegex.Match(text);
            if (conditional.Success) {
                marker = new BlockMarker(
                    BlockKind.Conditional,
                    conditional.Groups["name"].Value,
                    conditional.Groups["kind"].Value == "#");
                return true;
            }

            marker = default;
            return false;
        }

        private static bool TryGetEnumerable(object? value, out IEnumerable? enumerable) {
            if (value is IEnumerable items && value is not string && value is not byte[]) {
                enumerable = items;
                return true;
            }
            enumerable = null;
            return false;
        }

        private static bool IsScalar(object? value) {
            if (value == null) return true;
            Type type = value.GetType();
            return type.IsPrimitive || type.IsEnum || value is string || value is decimal || value is DateTime
                || value is DateTimeOffset || value is TimeSpan || value is Guid || value is Uri;
        }

        private static string FormatScalar(object? value, CultureInfo culture) {
            if (value == null) return string.Empty;
            if (value is IFormattable formattable) return formattable.ToString(null, culture) ?? string.Empty;
            return Convert.ToString(value, culture) ?? string.Empty;
        }

        private static void RemoveRange(IReadOnlyList<OpenXmlElement> elements, int start, int end) {
            for (int index = start; index <= end; index++) RemoveIfAttached(elements[index]);
        }

        private static void RemoveIfAttached(OpenXmlElement element) {
            if (element.Parent != null) element.Remove();
        }

        [UnconditionalSuppressMessage(
            "Trimming",
            "IL2070",
            Justification = "Reflection is enabled only by the RequiresUnreferencedCode POCO entry point; dictionary binding never enters this branch.")]
        private static bool TryGetValue(object? source, string name, bool allowReflection, out object? value) {
            if (source == null) {
                value = null;
                return false;
            }
            if (string.Equals(name, "this", StringComparison.OrdinalIgnoreCase)) {
                value = source;
                return true;
            }
            if (source is IReadOnlyDictionary<string, object?> readOnly && TryGetDictionaryValue(readOnly, name, out value)) return true;
            if (source is IDictionary<string, object?> dictionary && TryGetDictionaryValue(dictionary, name, out value)) return true;
            if (source is IDictionary legacy) {
                foreach (DictionaryEntry entry in legacy) {
                    if (entry.Key is string key && string.Equals(key, name, StringComparison.OrdinalIgnoreCase)) {
                        value = entry.Value;
                        return true;
                    }
                }
            }
            if (allowReflection) {
                IReadOnlyDictionary<string, PropertyInfo> properties = PropertyCache.GetOrAdd(source.GetType(), static type =>
                    new ReadOnlyDictionary<string, PropertyInfo>(type.GetProperties(BindingFlags.Instance | BindingFlags.Public)
                        .Where(static property => property.CanRead && property.GetIndexParameters().Length == 0)
                        .ToDictionary(static property => property.Name, StringComparer.OrdinalIgnoreCase)));
                if (properties.TryGetValue(name, out PropertyInfo? property)) {
                    value = property.GetValue(source);
                    return true;
                }
            }
            value = null;
            return false;
        }

        private static bool TryGetDictionaryValue<TDictionary>(TDictionary dictionary, string name, out object? value)
            where TDictionary : IEnumerable<KeyValuePair<string, object?>> {
            foreach (KeyValuePair<string, object?> pair in dictionary) {
                if (string.Equals(pair.Key, name, StringComparison.OrdinalIgnoreCase)) {
                    value = pair.Value;
                    return true;
                }
            }
            value = null;
            return false;
        }

        private sealed class TemplateScope {
            internal TemplateScope(object? value, TemplateScope? parent, bool allowReflection) {
                Value = value;
                Parent = parent;
                AllowReflection = allowReflection;
            }

            internal object? Value { get; }
            internal TemplateScope? Parent { get; }
            internal bool AllowReflection { get; }

            internal bool TryResolve(string path, out object? value) {
                for (TemplateScope? scope = this; scope != null; scope = scope.Parent) {
                    if (TryResolveWithin(scope.Value, path, scope.AllowReflection, out value)) return true;
                }
                value = null;
                return false;
            }

            private static bool TryResolveWithin(object? source, string path, bool allowReflection, out object? value) {
                if (TryGetValue(source, path, allowReflection, out value)) return true;

                string[] segments = path.Split('.');
                if (segments.Length < 2 || !TryGetValue(source, segments[0], allowReflection, out value)) return false;
                for (int index = 1; index < segments.Length; index++) {
                    if (!TryGetValue(value, segments[index], allowReflection, out value)) return false;
                }
                return true;
            }
        }

        private sealed class BindingState {
            internal BindingState(WordDocument document, WordTemplateOptions options) {
                Document = document;
                Options = options;
                foreach (OpenXmlCompositeElement root in WordMailMerge.EnumerateTemplateRoots(document)) {
                    foreach (BookmarkStart bookmark in root.Descendants<BookmarkStart>()) {
                        string? id = bookmark.Id?.Value;
                        if (int.TryParse(id, NumberStyles.None, CultureInfo.InvariantCulture, out int numericId)) {
                            NextBookmarkId = Math.Max(NextBookmarkId, numericId + 1);
                        }
                        if (!string.IsNullOrWhiteSpace(bookmark.Name?.Value)) {
                            string name = bookmark.Name!.Value!;
                            BookmarkNames.Add(name);
                            BookmarkNameCounts[name] = BookmarkNameCounts.TryGetValue(name, out int count) ? count + 1 : 1;
                        }
                    }
                }
            }

            internal WordDocument Document { get; }
            internal WordTemplateOptions Options { get; }
            internal int PlaceholderCount { get; set; }
            internal int ReplacedPlaceholderCount { get; set; }
            internal int RepeatedBlockCount { get; set; }
            internal int ConditionalBlockCount { get; set; }
            internal HashSet<string> MissingValueNames { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            private HashSet<string> BookmarkNames { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            private Dictionary<string, int> BookmarkNameCounts { get; } = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            private int NextBookmarkId { get; set; }

            internal void ReleaseRepeatedTemplateBookmarkNames(IEnumerable<OpenXmlElement> templateElements) {
                foreach (BookmarkStart bookmark in templateElements.SelectMany(element => element.Descendants<BookmarkStart>())) {
                    string? name = bookmark.Name?.Value;
                    if (string.IsNullOrWhiteSpace(name) || !BookmarkNameCounts.TryGetValue(name!, out int count)) continue;
                    if (count > 1) {
                        BookmarkNameCounts[name!] = count - 1;
                    } else {
                        BookmarkNameCounts.Remove(name!);
                        BookmarkNames.Remove(name!);
                    }
                }
            }

            internal void NormalizeRepeatedClone(OpenXmlElement root) {
                WordDrawingIdAllocator.Reassign(Document, root);
                var bookmarkIds = new Dictionary<string, string>(StringComparer.Ordinal);
                var bookmarkNameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                foreach (BookmarkStart start in root.Descendants<BookmarkStart>()) {
                    string oldId = start.Id?.Value ?? string.Empty;
                    string newId = NextBookmarkId++.ToString(CultureInfo.InvariantCulture);
                    start.Id = newId;
                    if (oldId.Length > 0) bookmarkIds[oldId] = newId;

                    string oldName = start.Name?.Value ?? "Bookmark";
                    string newName = ReserveBookmarkName(oldName);
                    start.Name = newName;
                    if (oldName.Length > 0) bookmarkNameMap[oldName] = newName;
                }

                foreach (BookmarkEnd end in root.Descendants<BookmarkEnd>()) {
                    string oldId = end.Id?.Value ?? string.Empty;
                    if (bookmarkIds.TryGetValue(oldId, out string? newId)) end.Id = newId;
                }

                foreach (Hyperlink hyperlink in root.Descendants<Hyperlink>()) {
                    string oldAnchor = hyperlink.Anchor?.Value ?? string.Empty;
                    if (bookmarkNameMap.TryGetValue(oldAnchor, out string? newAnchor)) hyperlink.Anchor = newAnchor;
                }

                foreach (OpenXmlElement element in root.Descendants()) {
                    ReassignHexIdentifier(element, "paraId");
                    ReassignHexIdentifier(element, "textId");
                }
            }

            private string ReserveBookmarkName(string originalName) {
                string baseName = string.IsNullOrWhiteSpace(originalName) ? "Bookmark" : originalName;
                int suffix = 2;
                string candidate = baseName;
                while (!BookmarkNames.Add(candidate)) {
                    string suffixText = "_" + (suffix++).ToString(CultureInfo.InvariantCulture);
                    candidate = baseName.Substring(0, Math.Min(baseName.Length, 40 - suffixText.Length)) + suffixText;
                }
                BookmarkNameCounts[candidate] = 1;
                return candidate;
            }

            private static void ReassignHexIdentifier(OpenXmlElement element, string localName) {
                const string Word2010Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";
                OpenXmlAttribute attribute = element.GetAttributes().FirstOrDefault(candidate =>
                    candidate.LocalName == localName && candidate.NamespaceUri == Word2010Namespace);
                if (string.IsNullOrEmpty(attribute.Value)) return;
                element.SetAttribute(new OpenXmlAttribute("w14", localName, Word2010Namespace,
                    Guid.NewGuid().ToString("N").Substring(0, 8).ToUpperInvariant()));
            }
        }

        private enum BlockKind { Conditional, Repeating }

        private readonly struct BlockMarker {
            internal BlockMarker(BlockKind kind, string name, bool isStart) {
                Kind = kind;
                Name = name;
                IsStart = isStart;
            }
            internal BlockKind Kind { get; }
            internal string Name { get; }
            internal bool IsStart { get; }
        }

        private readonly struct BlockStart {
            internal BlockStart(BlockKind kind, string name, int index) {
                Kind = kind;
                Name = name;
                Index = index;
            }
            internal BlockKind Kind { get; }
            internal string Name { get; }
            internal int Index { get; }
        }

        private readonly struct BlockRange {
            internal BlockRange(BlockKind kind, string name, int startIndex, int endIndex) {
                Kind = kind;
                Name = name;
                StartIndex = startIndex;
                EndIndex = endIndex;
            }
            internal BlockKind Kind { get; }
            internal string Name { get; }
            internal int StartIndex { get; }
            internal int EndIndex { get; }
        }

        private readonly struct TextSlice {
            internal TextSlice(Text text, int start, int end) {
                Text = text;
                Start = start;
                End = end;
            }
            internal Text Text { get; }
            internal int Start { get; }
            internal int End { get; }
        }
    }
}
