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
        private static WordContentControlDataBindingResult ExecuteContentControlDataBindingsCore(WordDocument document, IDictionary<string, string>? values, bool updateCustomXml) {
            MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart is missing.");
            var customXmlCache = new Dictionary<CustomXmlPart, XDocument>();
            var dirtyCustomXmlParts = new HashSet<CustomXmlPart>();
            var missingKeys = new List<string>();
            int bindingCount = 0;
            int updatedContentControls = 0;
            int updatedCustomXmlNodes = 0;

            foreach (var boundControl in EnumerateBoundContentControls(mainPart)) {
                bindingCount++;

                string? value = null;
                bool hasSuppliedValue = values != null && TryGetSuppliedBindingValue(values, boundControl, out value);
                if (!hasSuppliedValue && !TryGetCustomXmlBindingValue(mainPart, boundControl.Binding, customXmlCache, out value)) {
                    missingKeys.Add(GetBindingDisplayKey(boundControl));
                    continue;
                }

                if (UpdateContentControlText(boundControl.ContentControl, value!)) {
                    updatedContentControls++;
                }

                if (updateCustomXml && hasSuppliedValue
                    && TrySelectCustomXmlBindingElement(mainPart, boundControl.Binding, customXmlCache, out CustomXmlPart? customXmlPart, out XElement? element)) {
                    if (!string.Equals(element.Value, value, StringComparison.Ordinal)) {
                        element.Value = value!;
                        updatedCustomXmlNodes++;
                        dirtyCustomXmlParts.Add(customXmlPart);
                    }
                }
            }

            foreach (CustomXmlPart part in dirtyCustomXmlParts) {
                SaveCustomXmlPart(part, customXmlCache[part]);
            }

            return new WordContentControlDataBindingResult(
                bindingCount,
                updatedContentControls,
                updatedCustomXmlNodes,
                missingKeys.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(key => key, StringComparer.OrdinalIgnoreCase).ToList());
        }

        private static IEnumerable<BoundContentControl> EnumerateBoundContentControls(MainDocumentPart mainPart) {
            foreach (OpenXmlPart part in EnumerateParts(mainPart)) {
                OpenXmlPartRootElement? root = GetPartRootElement(part);
                if (root == null) {
                    continue;
                }

                foreach (SdtRun control in root.Descendants<SdtRun>()) {
                    DataBinding? binding = control.SdtProperties?.GetFirstChild<DataBinding>();
                    if (binding != null) {
                        yield return new BoundContentControl(part, control, control.SdtProperties!, binding);
                    }
                }

                foreach (SdtBlock control in root.Descendants<SdtBlock>()) {
                    DataBinding? binding = control.SdtProperties?.GetFirstChild<DataBinding>();
                    if (binding != null) {
                        yield return new BoundContentControl(part, control, control.SdtProperties!, binding);
                    }
                }

                foreach (SdtCell control in root.Descendants<SdtCell>()) {
                    DataBinding? binding = control.SdtProperties?.GetFirstChild<DataBinding>();
                    if (binding != null) {
                        yield return new BoundContentControl(part, control, control.SdtProperties!, binding);
                    }
                }
            }
        }

        private static OpenXmlPartRootElement? GetPartRootElement(OpenXmlPart part) {
            switch (part) {
                case MainDocumentPart mainPart:
                    return mainPart.Document;
                case HeaderPart headerPart:
                    return headerPart.Header;
                case FooterPart footerPart:
                    return footerPart.Footer;
                case FootnotesPart footnotesPart:
                    return footnotesPart.Footnotes;
                case EndnotesPart endnotesPart:
                    return endnotesPart.Endnotes;
                case WordprocessingCommentsPart commentsPart:
                    return commentsPart.Comments;
                default:
                    return part.RootElement;
            }
        }

        private static IEnumerable<OpenXmlPart> EnumerateParts(OpenXmlPart root) {
            yield return root;

            foreach (IdPartPair pair in root.Parts) {
                OpenXmlPart part = pair.OpenXmlPart;
                foreach (OpenXmlPart child in EnumerateParts(part)) {
                    yield return child;
                }
            }
        }

        private static bool TryGetSuppliedBindingValue(IDictionary<string, string> values, BoundContentControl boundControl, out string? value) {
            foreach (string key in GetBindingKeys(boundControl)) {
                if (TryGetValueOrdinalIgnoreCase(values, key, out value)) {
                    return true;
                }
            }

            value = null;
            return false;
        }

        private static bool TryGetValueOrdinalIgnoreCase(IDictionary<string, string> values, string key, out string? value) {
            if (values.TryGetValue(key, out string? exactValue)) {
                value = exactValue;
                return true;
            }

            foreach (var pair in values) {
                if (string.Equals(pair.Key, key, StringComparison.OrdinalIgnoreCase)) {
                    value = pair.Value;
                    return true;
                }
            }

            value = null;
            return false;
        }

        private static IEnumerable<string> GetBindingKeys(BoundContentControl boundControl) {
            string? alias = boundControl.Properties.Descendants<SdtAlias>().FirstOrDefault()?.Val?.ToString();
            if (!string.IsNullOrWhiteSpace(alias)) {
                yield return alias!;
            }

            string? tag = boundControl.Properties.Descendants<Tag>().FirstOrDefault()?.Val?.ToString();
            if (!string.IsNullOrWhiteSpace(tag)) {
                yield return tag!;
            }

            string? xpath = boundControl.Binding.XPath?.Value;
            if (!string.IsNullOrWhiteSpace(xpath)) {
                yield return xpath!;

                string? storeItemId = boundControl.Binding.StoreItemId?.Value;
                if (!string.IsNullOrWhiteSpace(storeItemId)) {
                    yield return storeItemId + "|" + xpath;
                }
            }
        }

        private static string GetBindingDisplayKey(BoundContentControl boundControl) {
            return GetBindingKeys(boundControl).FirstOrDefault()
                ?? boundControl.Binding.XPath?.Value
                ?? boundControl.Binding.StoreItemId?.Value
                ?? boundControl.Part.Uri.OriginalString;
        }

        private static bool TryGetCustomXmlBindingValue(MainDocumentPart mainPart, DataBinding binding, Dictionary<CustomXmlPart, XDocument> cache, out string? value) {
            if (TrySelectCustomXmlBindingElement(mainPart, binding, cache, out _, out XElement? element)) {
                value = element.Value;
                return true;
            }

            value = null;
            return false;
        }

        private static bool TrySelectCustomXmlBindingElement(MainDocumentPart mainPart, DataBinding binding, Dictionary<CustomXmlPart, XDocument> cache, out CustomXmlPart customXmlPart, out XElement element) {
            string? storeItemId = binding.StoreItemId?.Value;
            IEnumerable<CustomXmlPart> parts = mainPart.CustomXmlParts;
            if (!string.IsNullOrWhiteSpace(storeItemId)) {
                var matchingParts = parts
                    .Where(part => string.Equals(GetCustomXmlStoreItemId(part), storeItemId, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                if (matchingParts.Count > 0) {
                    parts = matchingParts;
                }
            }

            foreach (CustomXmlPart part in parts) {
                XDocument document = GetCustomXmlDocument(part, cache);
                XElement? selected = SelectBoundElement(document, binding);
                if (selected != null) {
                    customXmlPart = part;
                    element = selected;
                    return true;
                }
            }

            customXmlPart = null!;
            element = null!;
            return false;
        }

        private static string? GetCustomXmlStoreItemId(CustomXmlPart part) {
            return part.CustomXmlPropertiesPart?.DataStoreItem?.ItemId?.Value;
        }

        private static XDocument GetCustomXmlDocument(CustomXmlPart part, Dictionary<CustomXmlPart, XDocument> cache) {
            if (cache.TryGetValue(part, out XDocument? document)) {
                return document;
            }

            using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read)) {
                document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            }

            cache.Add(part, document);
            return document;
        }

        private static XElement? SelectBoundElement(XDocument document, DataBinding binding) {
            string? xpath = binding.XPath?.Value;
            if (string.IsNullOrWhiteSpace(xpath)) {
                return null;
            }

            XmlNamespaceManager namespaceManager = CreateNamespaceManager(binding.PrefixMappings?.Value);
            return document.XPathSelectElement(xpath!, namespaceManager);
        }

        private static XmlNamespaceManager CreateNamespaceManager(string? prefixMappings) {
            var manager = new XmlNamespaceManager(new NameTable());
            if (string.IsNullOrWhiteSpace(prefixMappings)) {
                return manager;
            }

            foreach (Match match in Regex.Matches(prefixMappings!, @"xmlns(?::(?<prefix>[\w.-]+))?\s*=\s*(?<quote>['""])(?<uri>.*?)\k<quote>")) {
                string prefix = match.Groups["prefix"].Success ? match.Groups["prefix"].Value : string.Empty;
                string uri = match.Groups["uri"].Value;
                if (!string.IsNullOrWhiteSpace(uri)) {
                    manager.AddNamespace(prefix, uri);
                }
            }

            return manager;
        }

        private static void SaveCustomXmlPart(CustomXmlPart part, XDocument document) {
            using (Stream stream = part.GetStream(FileMode.Create, FileAccess.Write)) {
                document.Save(stream);
            }
        }

        private static bool UpdateContentControlText(OpenXmlElement contentControl, string value) {
            switch (contentControl) {
                case SdtRun run:
                    run.SdtContentRun ??= new SdtContentRun();
                    SetTextInComposite(run.SdtContentRun, value, () => {
                        var newRun = new Run(new Text { Space = SpaceProcessingModeValues.Preserve });
                        run.SdtContentRun.Append(newRun);
                        return newRun.GetFirstChild<Text>()!;
                    });
                    return true;
                case SdtBlock block:
                    block.SdtContentBlock ??= new SdtContentBlock();
                    SetTextInComposite(block.SdtContentBlock, value, () => {
                        var paragraph = new Paragraph(new Run(new Text { Space = SpaceProcessingModeValues.Preserve }));
                        block.SdtContentBlock.Append(paragraph);
                        return paragraph.Descendants<Text>().First();
                    });
                    return true;
                case SdtCell cell:
                    cell.SdtContentCell ??= new SdtContentCell();
                    SetTextInComposite(cell.SdtContentCell, value, () => {
                        var tableCell = new TableCell(new Paragraph(new Run(new Text { Space = SpaceProcessingModeValues.Preserve })));
                        cell.SdtContentCell.Append(tableCell);
                        return tableCell.Descendants<Text>().First();
                    });
                    return true;
                default:
                    return false;
            }
        }

        private static void SetTextInComposite(OpenXmlCompositeElement container, string value, Func<Text> createText) {
            var textElements = container.Descendants<Text>().ToList();
            Text firstText = textElements.FirstOrDefault() ?? createText();
            firstText.Text = value;
            firstText.Space = SpaceProcessingModeValues.Preserve;

            foreach (Text extraText in textElements.Skip(1)) {
                extraText.Text = string.Empty;
            }
        }
        private sealed class BoundContentControl {
            internal BoundContentControl(OpenXmlPart part, OpenXmlElement contentControl, SdtProperties properties, DataBinding binding) {
                Part = part;
                ContentControl = contentControl;
                Properties = properties;
                Binding = binding;
            }

            internal OpenXmlPart Part { get; }
            internal OpenXmlElement ContentControl { get; }
            internal SdtProperties Properties { get; }
            internal DataBinding Binding { get; }
        }
    }
}
