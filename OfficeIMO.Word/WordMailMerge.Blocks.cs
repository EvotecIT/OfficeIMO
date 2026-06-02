using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeIMO.Word {
    public static partial class WordMailMerge {
        private static int ExecuteRepeatingBlocks(OpenXmlCompositeElement container, IDictionary<string, IEnumerable<WordMailMergeBlockData>> regions, bool removeFields) {
            var elements = container.ChildElements.ToList();
            var stack = new List<RepeatingBlockStart>();
            var blocks = new List<RepeatingBlockRange>();

            for (int index = 0; index < elements.Count; index++) {
                if (elements[index] is not Paragraph paragraph || !TryGetRepeatingBlockMarker(paragraph, out string? markerName, out bool isStart)) {
                    continue;
                }

                if (isStart) {
                    stack.Add(new RepeatingBlockStart(markerName!, index));
                    continue;
                }

                if (stack.Count == 0) {
                    throw new InvalidOperationException($"Repeating block end marker '{{{{/each {markerName}}}}}' does not have a matching start marker.");
                }

                var start = stack[stack.Count - 1];
                stack.RemoveAt(stack.Count - 1);
                if (!string.Equals(start.Name, markerName, StringComparison.OrdinalIgnoreCase)) {
                    throw new InvalidOperationException($"Repeating block end marker '{{{{/each {markerName}}}}}' does not match start marker '{{{{#each {start.Name}}}}}'.");
                }

                if (stack.Count == 0) {
                    blocks.Add(new RepeatingBlockRange(start.Name, start.Index, index));
                }
            }

            if (stack.Count > 0) {
                var start = stack[stack.Count - 1];
                throw new InvalidOperationException($"Repeating block start marker '{{{{#each {start.Name}}}}}' does not have a matching end marker.");
            }

            int generated = 0;
            foreach (var block in blocks.OrderByDescending(item => item.StartIndex)) {
                if (!TryGetRegionRows(regions, block.Name, out IEnumerable<WordMailMergeBlockData>? rows)) {
                    throw new InvalidOperationException($"Repeating block '{block.Name}' was not supplied.");
                }

                var templateElements = elements
                    .Skip(block.StartIndex + 1)
                    .Take(block.EndIndex - block.StartIndex - 1)
                    .ToList();
                var rowValues = rows!.ToList();
                OpenXmlElement insertionPoint = elements[block.StartIndex];

                foreach (var row in rowValues) {
                    if (row == null) throw new ArgumentException("Repeating block rows cannot contain null items.", nameof(regions));

                    var workingContainer = new SdtContentBlock();
                    foreach (OpenXmlElement templateElement in templateElements) {
                        var clonedElement = templateElement.CloneNode(true);
                        workingContainer.Append(clonedElement);
                    }

                    generated += ExecuteRepeatingBlocks(workingContainer, row.Regions, removeFields);
                    ReplaceMergeFields(workingContainer, row.Values, removeFields);

                    foreach (OpenXmlElement clonedElement in workingContainer.ChildElements.ToList()) {
                        clonedElement.Remove();
                        insertionPoint.InsertBeforeSelf(clonedElement);
                    }

                    generated++;
                }

                for (int index = block.StartIndex; index <= block.EndIndex; index++) {
                    RemoveIfAttached(elements[index]);
                }
            }

            foreach (var child in container.ChildElements.OfType<OpenXmlCompositeElement>().ToList()) {
                if (!CanContainTemplateBlockMarkers(child)) {
                    continue;
                }

                generated += ExecuteRepeatingBlocks(child, regions, removeFields);
            }

            return generated;
        }

        private static int ExecuteConditionalBlocks(OpenXmlCompositeElement container, IDictionary<string, bool> conditions, bool removeMarkers) {
            var elements = container.ChildElements.ToList();
            var stack = new List<ConditionalBlockStart>();
            var blocks = new List<ConditionalBlockRange>();

            for (int index = 0; index < elements.Count; index++) {
                if (elements[index] is not Paragraph paragraph || !TryGetConditionalBlockMarker(paragraph, out string? markerName, out bool isStart)) {
                    continue;
                }

                if (isStart) {
                    stack.Add(new ConditionalBlockStart(markerName!, index));
                    continue;
                }

                if (stack.Count == 0) {
                    throw new InvalidOperationException($"Conditional block end marker '{{{{/{markerName}}}}}' does not have a matching start marker.");
                }

                var start = stack[stack.Count - 1];
                stack.RemoveAt(stack.Count - 1);
                if (!string.Equals(start.Name, markerName, StringComparison.OrdinalIgnoreCase)) {
                    throw new InvalidOperationException($"Conditional block end marker '{{{{/{markerName}}}}}' does not match start marker '{{{{#{start.Name}}}}}'.");
                }

                blocks.Add(new ConditionalBlockRange(start.Name, start.Index, index));
            }

            if (stack.Count > 0) {
                var start = stack[stack.Count - 1];
                throw new InvalidOperationException($"Conditional block start marker '{{{{#{start.Name}}}}}' does not have a matching end marker.");
            }

            foreach (var block in blocks) {
                if (!conditions.ContainsKey(block.Name)) {
                    throw new InvalidOperationException($"Conditional block '{block.Name}' was not supplied.");
                }
            }

            foreach (var block in blocks.OrderByDescending(item => item.StartIndex)) {
                bool include = conditions[block.Name];

                if (include) {
                    if (removeMarkers) {
                        RemoveIfAttached(elements[block.EndIndex]);
                        RemoveIfAttached(elements[block.StartIndex]);
                    }

                    continue;
                }

                for (int index = block.StartIndex; index <= block.EndIndex; index++) {
                    RemoveIfAttached(elements[index]);
                }
            }

            int processed = blocks.Count;
            foreach (var child in container.ChildElements.OfType<OpenXmlCompositeElement>().ToList()) {
                if (!CanContainTemplateBlockMarkers(child)) {
                    continue;
                }

                processed += ExecuteConditionalBlocks(child, conditions, removeMarkers);
            }

            return processed;
        }

        private static bool TryGetConditionalBlockMarker(Paragraph paragraph, out string? name, out bool isStart) {
            string text = paragraph.InnerText ?? string.Empty;
            var match = ConditionalBlockMarkerRegex.Match(text);
            if (!match.Success) {
                name = null;
                isStart = false;
                return false;
            }

            name = match.Groups["name"].Value;
            isStart = match.Groups["kind"].Value == "#";
            return true;
        }

        private static bool TryGetRepeatingBlockMarker(Paragraph paragraph, out string? name, out bool isStart) {
            string text = paragraph.InnerText ?? string.Empty;
            var match = RepeatingBlockMarkerRegex.Match(text);
            if (!match.Success) {
                name = null;
                isStart = false;
                return false;
            }

            name = match.Groups["name"].Value;
            isStart = match.Groups["kind"].Value.StartsWith("#", StringComparison.Ordinal);
            return true;
        }

        private static void RemoveIfAttached(OpenXmlElement element) {
            if (element.Parent != null) {
                element.Remove();
            }
        }
        private static Dictionary<string, IEnumerable<WordMailMergeBlockData>> ConvertFlatRegions(IDictionary<string, IEnumerable<IDictionary<string, string>>> regions) {
            var converted = new Dictionary<string, IEnumerable<WordMailMergeBlockData>>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in regions) {
                converted[pair.Key] = pair.Value.Select(values => new WordMailMergeBlockData(values));
            }

            return converted;
        }

        private static bool TryGetRegionRows(IDictionary<string, IEnumerable<WordMailMergeBlockData>> regions, string name, out IEnumerable<WordMailMergeBlockData>? rows) {
            if (regions.TryGetValue(name, out rows)) {
                return true;
            }

            foreach (var pair in regions) {
                if (string.Equals(pair.Key, name, StringComparison.OrdinalIgnoreCase)) {
                    rows = pair.Value;
                    return true;
                }
            }

            rows = null;
            return false;
        }

        private static ConditionalBlockInspection InspectConditionalBlocks(OpenXmlCompositeElement? container) {
            if (container == null) {
                return new ConditionalBlockInspection(Array.Empty<string>(), Array.Empty<WordMailMergeTemplateIssue>());
            }

            var names = new List<string>();
            var issues = new List<WordMailMergeTemplateIssue>();
            InspectConditionalBlocks(container, names, issues);
            return new ConditionalBlockInspection(names, issues);
        }

        private static RepeatingBlockInspection InspectRepeatingBlocks(OpenXmlCompositeElement? container) {
            if (container == null) {
                return new RepeatingBlockInspection(Array.Empty<string>(), Array.Empty<WordMailMergeTemplateIssue>());
            }

            var names = new List<string>();
            var issues = new List<WordMailMergeTemplateIssue>();
            InspectRepeatingBlocks(container, names, issues);
            return new RepeatingBlockInspection(names, issues);
        }

        private static void InspectRepeatingBlocks(OpenXmlCompositeElement container, List<string> names, List<WordMailMergeTemplateIssue> issues) {
            var elements = container.ChildElements.ToList();
            var stack = new List<RepeatingBlockStart>();

            for (int index = 0; index < elements.Count; index++) {
                if (elements[index] is not Paragraph paragraph || !TryGetRepeatingBlockMarker(paragraph, out string? markerName, out bool isStart)) {
                    continue;
                }

                names.Add(markerName!);
                if (isStart) {
                    stack.Add(new RepeatingBlockStart(markerName!, index));
                    continue;
                }

                if (stack.Count == 0) {
                    issues.Add(new WordMailMergeTemplateIssue(
                        WordMailMergeTemplateIssueKind.UnmatchedRepeatingBlockEnd,
                        markerName!,
                        $"Repeating block end marker '{{{{/each {markerName}}}}}' does not have a matching start marker."));
                    continue;
                }

                var start = stack[stack.Count - 1];
                stack.RemoveAt(stack.Count - 1);
                if (!string.Equals(start.Name, markerName, StringComparison.OrdinalIgnoreCase)) {
                    issues.Add(new WordMailMergeTemplateIssue(
                        WordMailMergeTemplateIssueKind.MismatchedRepeatingBlockEnd,
                        markerName!,
                        $"Repeating block end marker '{{{{/each {markerName}}}}}' does not match start marker '{{{{#each {start.Name}}}}}'."));
                }
            }

            foreach (var start in stack) {
                issues.Add(new WordMailMergeTemplateIssue(
                    WordMailMergeTemplateIssueKind.UnmatchedRepeatingBlockStart,
                    start.Name,
                    $"Repeating block start marker '{{{{#each {start.Name}}}}}' does not have a matching end marker."));
            }

            foreach (var child in container.ChildElements.OfType<OpenXmlCompositeElement>().ToList()) {
                if (!CanContainTemplateBlockMarkers(child)) {
                    continue;
                }

                InspectRepeatingBlocks(child, names, issues);
            }
        }

        private static void InspectConditionalBlocks(OpenXmlCompositeElement container, List<string> names, List<WordMailMergeTemplateIssue> issues) {
            var elements = container.ChildElements.ToList();
            var stack = new List<ConditionalBlockStart>();

            for (int index = 0; index < elements.Count; index++) {
                if (elements[index] is not Paragraph paragraph || !TryGetConditionalBlockMarker(paragraph, out string? markerName, out bool isStart)) {
                    continue;
                }

                names.Add(markerName!);
                if (isStart) {
                    stack.Add(new ConditionalBlockStart(markerName!, index));
                    continue;
                }

                if (stack.Count == 0) {
                    issues.Add(new WordMailMergeTemplateIssue(
                        WordMailMergeTemplateIssueKind.UnmatchedConditionalEnd,
                        markerName!,
                        $"Conditional block end marker '{{{{/{markerName}}}}}' does not have a matching start marker."));
                    continue;
                }

                var start = stack[stack.Count - 1];
                stack.RemoveAt(stack.Count - 1);
                if (!string.Equals(start.Name, markerName, StringComparison.OrdinalIgnoreCase)) {
                    issues.Add(new WordMailMergeTemplateIssue(
                        WordMailMergeTemplateIssueKind.MismatchedConditionalEnd,
                        markerName!,
                        $"Conditional block end marker '{{{{/{markerName}}}}}' does not match start marker '{{{{#{start.Name}}}}}'."));
                }
            }

            foreach (var start in stack) {
                issues.Add(new WordMailMergeTemplateIssue(
                    WordMailMergeTemplateIssueKind.UnmatchedConditionalStart,
                    start.Name,
                    $"Conditional block start marker '{{{{#{start.Name}}}}}' does not have a matching end marker."));
            }

            foreach (var child in container.ChildElements.OfType<OpenXmlCompositeElement>().ToList()) {
                if (!CanContainTemplateBlockMarkers(child)) {
                    continue;
                }

                InspectConditionalBlocks(child, names, issues);
            }
        }

        private static bool CanContainTemplateBlockMarkers(OpenXmlCompositeElement element) {
            return element is not Paragraph && element is not Run;
        }

        private readonly struct ConditionalBlockStart {
            internal ConditionalBlockStart(string name, int index) {
                Name = name;
                Index = index;
            }

            internal string Name { get; }
            internal int Index { get; }
        }

        private readonly struct ConditionalBlockRange {
            internal ConditionalBlockRange(string name, int startIndex, int endIndex) {
                Name = name;
                StartIndex = startIndex;
                EndIndex = endIndex;
            }

            internal string Name { get; }
            internal int StartIndex { get; }
            internal int EndIndex { get; }
        }

        private readonly struct RepeatingBlockStart {
            internal RepeatingBlockStart(string name, int index) {
                Name = name;
                Index = index;
            }

            internal string Name { get; }
            internal int Index { get; }
        }

        private readonly struct RepeatingBlockRange {
            internal RepeatingBlockRange(string name, int startIndex, int endIndex) {
                Name = name;
                StartIndex = startIndex;
                EndIndex = endIndex;
            }

            internal string Name { get; }
            internal int StartIndex { get; }
            internal int EndIndex { get; }
        }

        private readonly struct ConditionalBlockInspection {
            internal ConditionalBlockInspection(IReadOnlyList<string> names, IReadOnlyList<WordMailMergeTemplateIssue> issues) {
                Names = names;
                Issues = issues;
            }

            internal IReadOnlyList<string> Names { get; }
            internal IReadOnlyList<WordMailMergeTemplateIssue> Issues { get; }
        }

        private readonly struct RepeatingBlockInspection {
            internal RepeatingBlockInspection(IReadOnlyList<string> names, IReadOnlyList<WordMailMergeTemplateIssue> issues) {
                Names = names;
                Issues = issues;
            }

            internal IReadOnlyList<string> Names { get; }
            internal IReadOnlyList<WordMailMergeTemplateIssue> Issues { get; }
        }

    }
}
