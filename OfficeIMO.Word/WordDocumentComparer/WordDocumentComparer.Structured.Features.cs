using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeFields(WordDocument source, WordDocument target, WordComparisonResult result, WordComparisonOptions options) {
            List<FieldSnapshot> sourceFields = GetFieldSnapshots(source, options);
            List<FieldSnapshot> targetFields = GetFieldSnapshots(target, options);
            AddFeatureRangeFindings(
                sourceFields,
                targetFields,
                WordComparisonScope.Field,
                "field",
                "Field",
                result);
        }

        private static void AnalyzeContentControls(WordDocument source, WordDocument target, WordComparisonResult result, WordComparisonOptions options) {
            List<ContentControlSnapshot> sourceControls = GetContentControlSnapshots(source, options);
            List<ContentControlSnapshot> targetControls = GetContentControlSnapshots(target, options);
            AddFeatureRangeFindings(
                sourceControls,
                targetControls,
                WordComparisonScope.ContentControl,
                "content-control",
                "Content control",
                result);
        }

        private static void AddFeatureRangeFindings<TSnapshot>(
            IReadOnlyList<TSnapshot> sourceItems,
            IReadOnlyList<TSnapshot> targetItems,
            WordComparisonScope scope,
            string locationPrefix,
            string label,
            WordComparisonResult result)
            where TSnapshot : IFeatureSnapshot {
            IReadOnlyList<MatchedIndexPair> matches = FindMatchingIndexes(
                sourceItems,
                targetItems,
                FeatureSnapshotKeyComparer<TSnapshot>.Instance);

            int sourceStart = 0;
            int targetStart = 0;
            foreach (MatchedIndexPair match in matches) {
                AddFeatureUnmatchedRangeFindings(sourceItems, targetItems, scope, locationPrefix, label, sourceStart, match.SourceIndex, targetStart, match.TargetIndex, result);
                AddFeatureModifiedFinding(sourceItems[match.SourceIndex], targetItems[match.TargetIndex], scope, locationPrefix, label, result);
                sourceStart = match.SourceIndex + 1;
                targetStart = match.TargetIndex + 1;
            }

            AddFeatureUnmatchedRangeFindings(sourceItems, targetItems, scope, locationPrefix, label, sourceStart, sourceItems.Count, targetStart, targetItems.Count, result);
        }

        private static void AddFeatureUnmatchedRangeFindings<TSnapshot>(
            IReadOnlyList<TSnapshot> sourceItems,
            IReadOnlyList<TSnapshot> targetItems,
            WordComparisonScope scope,
            string locationPrefix,
            string label,
            int sourceStart,
            int sourceEnd,
            int targetStart,
            int targetEnd,
            WordComparisonResult result)
            where TSnapshot : IFeatureSnapshot {
            int sourceIndex = sourceStart;
            int targetIndex = targetStart;

            while (sourceIndex < sourceEnd && targetIndex < targetEnd) {
                TSnapshot sourceItem = sourceItems[sourceIndex];
                TSnapshot targetItem = targetItems[targetIndex];
                if (CanPairUnmatchedFeatureItems(sourceItem, targetItem)) {
                    AddFeatureModifiedFinding(sourceItem, targetItem, scope, locationPrefix, label, result);
                    sourceIndex++;
                    targetIndex++;
                    continue;
                }

                if (sourceItem.DocumentOrder <= targetItem.DocumentOrder) {
                    AddFeatureDeletedFinding(sourceItem, scope, locationPrefix, label, result);
                    sourceIndex++;
                } else {
                    AddFeatureInsertedFinding(targetItem, scope, locationPrefix, label, result);
                    targetIndex++;
                }
            }

            while (targetIndex < targetEnd) {
                AddFeatureInsertedFinding(targetItems[targetIndex], scope, locationPrefix, label, result);
                targetIndex++;
            }

            while (sourceIndex < sourceEnd) {
                AddFeatureDeletedFinding(sourceItems[sourceIndex], scope, locationPrefix, label, result);
                sourceIndex++;
            }
        }

        private static bool CanPairUnmatchedFeatureItems<TSnapshot>(TSnapshot sourceItem, TSnapshot targetItem)
            where TSnapshot : IFeatureSnapshot {
            if (sourceItem is FieldSnapshot sourceField && targetItem is FieldSnapshot targetField) {
                return string.Equals(sourceField.LocationKey, targetField.LocationKey, StringComparison.Ordinal);
            }

            if (sourceItem is ContentControlSnapshot sourceControl && targetItem is ContentControlSnapshot targetControl) {
                return string.Equals(sourceControl.LocationKey, targetControl.LocationKey, StringComparison.Ordinal);
            }

            return string.Equals(GetFeatureLocationScope(sourceItem.DetailedLocation), GetFeatureLocationScope(targetItem.DetailedLocation), StringComparison.Ordinal);
        }

        private static string GetFeatureLocationScope(string detailedLocation) {
            if (string.IsNullOrWhiteSpace(detailedLocation)) {
                return string.Empty;
            }

            int separatorIndex = detailedLocation.IndexOf("//", StringComparison.Ordinal);
            if (separatorIndex >= 0) {
                int nextSeparatorIndex = detailedLocation.IndexOf('/', separatorIndex + 2);
                if (nextSeparatorIndex >= 0) {
                    nextSeparatorIndex = detailedLocation.IndexOf('/', nextSeparatorIndex + 1);
                }

                return nextSeparatorIndex >= 0
                    ? detailedLocation.Substring(0, nextSeparatorIndex)
                    : detailedLocation;
            }

            string[] parts = detailedLocation.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            return parts.Length >= 2
                ? parts[0] + "/" + parts[1]
                : detailedLocation;
        }

        private static void AddFeatureInsertedFinding<TSnapshot>(
            TSnapshot targetItem,
            WordComparisonScope scope,
            string locationPrefix,
            string label,
            WordComparisonResult result)
            where TSnapshot : IFeatureSnapshot {
            result.Add(new WordComparisonFinding(
                scope,
                WordComparisonChangeKind.Inserted,
                FeatureLocation(locationPrefix, targetItem.Index),
                null,
                targetItem.Index,
                null,
                targetItem.DisplayText,
                label + " inserted.",
                targetItem.DetailedLocation),
                targetItem.DocumentOrder);
        }

        private static void AddFeatureDeletedFinding<TSnapshot>(
            TSnapshot sourceItem,
            WordComparisonScope scope,
            string locationPrefix,
            string label,
            WordComparisonResult result)
            where TSnapshot : IFeatureSnapshot {
            result.Add(new WordComparisonFinding(
                scope,
                WordComparisonChangeKind.Deleted,
                FeatureLocation(locationPrefix, sourceItem.Index),
                sourceItem.Index,
                null,
                sourceItem.DisplayText,
                null,
                label + " deleted.",
                sourceItem.DetailedLocation),
                sourceItem.DocumentOrder);
        }

        private static void AddFeatureModifiedFinding<TSnapshot>(
            TSnapshot sourceItem,
            TSnapshot targetItem,
            WordComparisonScope scope,
            string locationPrefix,
            string label,
            WordComparisonResult result)
            where TSnapshot : IFeatureSnapshot {
            if (string.Equals(sourceItem.Signature, targetItem.Signature, StringComparison.Ordinal)) {
                return;
            }

            result.Add(new WordComparisonFinding(
                scope,
                WordComparisonChangeKind.Modified,
                FeatureLocation(locationPrefix, targetItem.Index),
                sourceItem.Index,
                targetItem.Index,
                sourceItem.DisplayText,
                targetItem.DisplayText,
                label + " changed.",
                targetItem.DetailedLocation),
                targetItem.DocumentOrder);
        }

        private static List<FieldSnapshot> GetFieldSnapshots(WordDocument document, WordComparisonOptions options) {
            return document.InspectFields()
                .Select(field => new FieldSnapshot(
                    field.Index,
                    GetFieldLocationKey(field),
                    GetFieldMatchKey(field, options),
                    GetFieldSignature(field, options),
                    GetFieldDisplayText(field),
                    GetFieldDetailedLocation(field),
                    GetFeatureDocumentOrder(field.LocationKind, field.Index)))
                .ToList();
        }

        private static string GetFieldLocationKey(WordFieldInfo field) {
            return string.Join(
                "|",
                field.LocationKind.ToString(),
                field.PartUri,
                field.Representation.ToString(),
                field.IsInTable ? "table" : string.Empty,
                field.IsInContentControl ? "content-control" : string.Empty,
                field.IsInTextBox ? "text-box" : string.Empty);
        }

        private static string GetFieldMatchKey(WordFieldInfo field, WordComparisonOptions options) {
            return string.Join(
                "|",
                field.LocationKind.ToString(),
                field.PartUri,
                field.Representation.ToString(),
                field.FieldType?.ToString() ?? string.Empty,
                NormalizeComparisonText(field.InstructionText, options),
                field.NestingLevel.ToString(System.Globalization.CultureInfo.InvariantCulture),
                field.IsInTable ? "table" : string.Empty,
                field.IsInContentControl ? "content-control" : string.Empty,
                field.IsInTextBox ? "text-box" : string.Empty);
        }

        private static string GetFieldSignature(WordFieldInfo field, WordComparisonOptions options) {
            return string.Join(
                "|",
                field.LocationKind.ToString(),
                field.PartUri,
                field.Representation.ToString(),
                field.FieldType?.ToString() ?? string.Empty,
                NormalizeComparisonText(field.InstructionText, options),
                NormalizeComparisonText(field.ResultText, options),
                field.IsDirty ? "dirty" : "clean",
                field.IsLocked ? "locked" : "unlocked",
                field.NestingLevel.ToString(System.Globalization.CultureInfo.InvariantCulture),
                field.IsInTable ? "table" : string.Empty,
                field.IsInContentControl ? "content-control" : string.Empty,
                field.IsInTextBox ? "text-box" : string.Empty);
        }

        private static string GetFieldDisplayText(WordFieldInfo field) {
            string kind = field.FieldType?.ToString() ?? field.Representation.ToString();
            return kind + ": " + field.InstructionText + " => " + field.ResultText;
        }

        private static string GetFieldDetailedLocation(WordFieldInfo field) {
            return JoinFeatureLocation(
                field.LocationKind.ToString(),
                field.PartUri,
                FeatureLocation("field", field.Index),
                field.Representation.ToString(),
                field.IsInTable ? "table" : string.Empty,
                field.IsInContentControl ? "content-control" : string.Empty,
                field.IsInTextBox ? "text-box" : string.Empty);
        }

        private static List<ContentControlSnapshot> GetContentControlSnapshots(WordDocument document, WordComparisonOptions options) {
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            if (mainPart == null) {
                return new List<ContentControlSnapshot>();
            }

            var snapshots = new List<ContentControlSnapshot>();
            foreach (WordFieldInventory.FieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                foreach (OrderedElement ordered in EnumerateDescendantsWithOrder(root.Root, GetFeatureOrderBase(root.LocationKind))) {
                    if (ordered.Element is not SdtElement contentControl) {
                        continue;
                    }

                    SdtProperties? properties = contentControl.SdtProperties;
                    string alias = properties?.GetFirstChild<SdtAlias>()?.Val?.Value ?? string.Empty;
                    string tag = properties?.GetFirstChild<Tag>()?.Val?.Value ?? string.Empty;
                    DataBinding? binding = properties?.GetFirstChild<DataBinding>();
                    string bindingText = FormatDataBinding(binding);
                    string text = NormalizeFeatureText(contentControl.InnerText ?? string.Empty);
                    string key = string.Join(
                        "|",
                        root.LocationKind.ToString(),
                        root.PartUri,
                        contentControl.LocalName,
                        NormalizeComparisonText(alias, options),
                        NormalizeComparisonText(tag, options),
                        NormalizeComparisonText(bindingText, options),
                        IsInTableFeature(contentControl) ? "table" : string.Empty,
                        IsInContentControlFeature(contentControl) ? "nested-content-control" : string.Empty,
                        IsInTextBoxFeature(contentControl) ? "text-box" : string.Empty);
                    string signature = string.Join(
                        "|",
                        contentControl.LocalName,
                        NormalizeComparisonText(alias, options),
                        NormalizeComparisonText(tag, options),
                        NormalizeComparisonText(bindingText, options),
                        NormalizeComparisonText(text, options),
                        IsInTableFeature(contentControl) ? "table" : string.Empty,
                        IsInTextBoxFeature(contentControl) ? "text-box" : string.Empty);
                    string displayText = contentControl.LocalName + ": alias=" + alias + "; tag=" + tag + "; binding=" + bindingText + "; text=" + text;

                    snapshots.Add(new ContentControlSnapshot(
                        snapshots.Count,
                        GetContentControlLocationKey(root, contentControl),
                        key,
                        signature,
                        displayText,
                        JoinFeatureLocation(
                            root.LocationKind.ToString(),
                            root.PartUri,
                            FeatureLocation("content-control", snapshots.Count),
                            contentControl.LocalName,
                            IsInTableFeature(contentControl) ? "table" : string.Empty,
                            IsInContentControlFeature(contentControl) ? "nested-content-control" : string.Empty,
                            IsInTextBoxFeature(contentControl) ? "text-box" : string.Empty),
                        ordered.DocumentOrder));
                }
            }

            return snapshots;
        }

        private static string GetContentControlLocationKey(WordFieldInventory.FieldRoot root, SdtElement contentControl) {
            return string.Join(
                "|",
                root.LocationKind.ToString(),
                root.PartUri,
                contentControl.LocalName,
                IsInTableFeature(contentControl) ? "table" : string.Empty,
                IsInContentControlFeature(contentControl) ? "nested-content-control" : string.Empty,
                IsInTextBoxFeature(contentControl) ? "text-box" : string.Empty);
        }

        private static string FormatDataBinding(DataBinding? binding) {
            if (binding == null) {
                return string.Empty;
            }

            return string.Join(
                "|",
                binding.StoreItemId?.Value ?? string.Empty,
                binding.XPath?.Value ?? string.Empty,
                binding.PrefixMappings?.Value ?? string.Empty);
        }

        private static bool IsInTableFeature(OpenXmlElement element) => element.Ancestors<Table>().Any();

        private static int GetFeatureDocumentOrder(WordFieldLocationKind locationKind, int index) {
            return GetFeatureOrderBase(locationKind) + index;
        }

        private static int GetFeatureOrderBase(WordFieldLocationKind locationKind) {
            return locationKind switch {
                WordFieldLocationKind.Header => HeaderPartOrderBase,
                WordFieldLocationKind.Footer => FooterPartOrderBase,
                WordFieldLocationKind.Footnote => FootnotePartOrderBase,
                WordFieldLocationKind.Endnote => EndnotePartOrderBase,
                _ => BodyPartOrderBase
            };
        }

        private static bool IsInTextBoxFeature(OpenXmlElement element) =>
            element.Ancestors().Any(ancestor =>
                string.Equals(ancestor.LocalName, "txbxContent", StringComparison.Ordinal) ||
                string.Equals(ancestor.LocalName, "textbox", StringComparison.Ordinal));

        private static string NormalizeFeatureText(string value) =>
            string.IsNullOrEmpty(value) ? string.Empty : string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

        private static string FeatureLocation(string prefix, int index) {
            return prefix + "[" + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
        }

        private static string JoinFeatureLocation(params string[] parts) {
            return string.Join("/", parts.Where(part => !string.IsNullOrWhiteSpace(part)).ToArray());
        }

        private interface IFeatureSnapshot : IComparisonFingerprint {
            int Index { get; }

            string MatchKey { get; }

            string Signature { get; }

            string DisplayText { get; }

            string DetailedLocation { get; }

            int DocumentOrder { get; }
        }

        private sealed class FieldSnapshot : IFeatureSnapshot {
            internal FieldSnapshot(int index, string locationKey, string matchKey, string signature, string displayText, string detailedLocation, int documentOrder) {
                Index = index;
                LocationKey = locationKey;
                MatchKey = matchKey;
                Signature = signature;
                DisplayText = displayText;
                DetailedLocation = detailedLocation;
                DocumentOrder = documentOrder;
            }

            public int Index { get; }

            public string LocationKey { get; }

            public string MatchKey { get; }

            public string Signature { get; }

            public string DisplayText { get; }

            public string DetailedLocation { get; }

            public int DocumentOrder { get; }

            ulong IComparisonFingerprint.ComparisonFingerprint => GetOrdinalTextFingerprint(MatchKey);
        }

        private sealed class ContentControlSnapshot : IFeatureSnapshot {
            internal ContentControlSnapshot(int index, string locationKey, string matchKey, string signature, string displayText, string detailedLocation, int documentOrder) {
                Index = index;
                LocationKey = locationKey;
                MatchKey = matchKey;
                Signature = signature;
                DisplayText = displayText;
                DetailedLocation = detailedLocation;
                DocumentOrder = documentOrder;
            }

            public int Index { get; }

            public string LocationKey { get; }

            public string MatchKey { get; }

            public string Signature { get; }

            public string DisplayText { get; }

            public string DetailedLocation { get; }

            public int DocumentOrder { get; }

            ulong IComparisonFingerprint.ComparisonFingerprint => GetOrdinalTextFingerprint(MatchKey);
        }

        private sealed class FeatureSnapshotKeyComparer<TSnapshot> : IEqualityComparer<TSnapshot>
            where TSnapshot : IFeatureSnapshot {
            internal static readonly FeatureSnapshotKeyComparer<TSnapshot> Instance = new();

            public bool Equals(TSnapshot? x, TSnapshot? y) {
                if (ReferenceEquals(x, y)) {
                    return true;
                }

                if (x == null || y == null) {
                    return false;
                }

                return string.Equals(x.MatchKey, y.MatchKey, StringComparison.Ordinal);
            }

            public int GetHashCode(TSnapshot obj) {
                return StringComparer.Ordinal.GetHashCode(obj.MatchKey);
            }
        }
    }
}
