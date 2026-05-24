using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed class SharedStringCache {
        private static readonly XmlReaderSettings SharedStringXmlReaderSettings = CreateSharedStringXmlReaderSettings();

        private readonly SharedStringTablePart? _part;
        private readonly bool _preferDom;
        private readonly Lazy<List<string>> _items;
        private List<string>? _loadedItems;
        private readonly object _containsCacheLock = new object();
        private Dictionary<(string Text, StringComparison Comparison), HashSet<int>?>? _containsCache;

        private SharedStringCache(SharedStringTablePart? part, bool preferDom) {
            _part = part;
            _preferDom = preferDom;
            _items = new Lazy<List<string>>(LoadItems, LazyThreadSafetyMode.ExecutionAndPublication);
        }

        public static SharedStringCache Build(SpreadsheetDocument doc) {
            return new SharedStringCache(doc.WorkbookPart!.SharedStringTablePart, doc.FileOpenAccess != FileAccess.Read);
        }

        private List<string> LoadItems() {
            if (_part == null) return new List<string>();
            if (_preferDom && _part.SharedStringTable != null) {
                return LoadItemsFromDom();
            }

            if (TryLoadItemsXmlFast(out var items)) {
                return items;
            }

            return LoadItemsFromDom();
        }

        private List<string> LoadItemsFromDom() {
            var part = _part;
            if (part == null || part.SharedStringTable == null) return new List<string>();
            var table = part.SharedStringTable;
            var list = new List<string>((int)(table.UniqueCount?.Value ?? table.Count?.Value ?? 0));
            foreach (var item in table.Elements<SharedStringItem>()) {
                if (item.Text?.Text != null)
                    list.Add(item.Text.Text);
                else if (item.HasChildren)
                    list.Add(GetRunText(item));
                else
                    list.Add(string.Empty);
            }

            return list;
        }

        private bool TryLoadItemsXmlFast(out List<string> items) {
            items = new List<string>();

            try {
                using var stream = _part!.GetStream(FileMode.Open, FileAccess.Read);
                using var reader = XmlReader.Create(stream, SharedStringXmlReaderSettings);
                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element) {
                        continue;
                    }

                    if (reader.LocalName == "sst") {
                        int capacity = ParsePositiveIntAttribute(reader.GetAttribute("uniqueCount"));
                        if (capacity <= 0) {
                            capacity = ParsePositiveIntAttribute(reader.GetAttribute("count"));
                        }

                        if (capacity > 0) {
                            items.Capacity = capacity;
                        }

                        continue;
                    }

                    if (reader.LocalName == "si") {
                        items.Add(ReadSharedStringItemXml(reader));
                    }
                }

                return true;
            } catch (XmlException) {
                items = null!;
                return false;
            } catch (IOException) {
                items = null!;
                return false;
            } catch (UnauthorizedAccessException) {
                items = null!;
                return false;
            } catch (ObjectDisposedException) {
                items = null!;
                return false;
            }
        }

        private static XmlReaderSettings CreateSharedStringXmlReaderSettings() {
            return new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true,
                CloseInput = false
            };
        }

        private static string ReadSharedStringItemXml(XmlReader reader) {
            if (reader.IsEmptyElement) {
                return string.Empty;
            }

            int depth = reader.Depth;
            string? first = null;
            StringBuilder? builder = null;
            int phoneticRunDepth = -1;

            bool hasNode = reader.Read();
            if (hasNode && reader.NodeType == XmlNodeType.Element && reader.LocalName == "t") {
                first = reader.ReadElementContentAsString();
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == "si") {
                    return first;
                }

                hasNode = true;
            }

            while (hasNode) {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == "si") {
                    break;
                }

                if (phoneticRunDepth >= 0) {
                    if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == phoneticRunDepth && reader.LocalName == "rPh") {
                        phoneticRunDepth = -1;
                    }

                    hasNode = reader.Read();
                    continue;
                }

                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "rPh") {
                    if (!reader.IsEmptyElement) {
                        phoneticRunDepth = reader.Depth;
                    }

                    hasNode = reader.Read();
                    continue;
                }

                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "t") {
                    hasNode = reader.Read();
                    continue;
                }

                string text = reader.ReadElementContentAsString();
                if (builder != null) {
                    builder.Append(text);
                } else if (first == null) {
                    first = text;
                } else {
                    builder = new StringBuilder(first.Length + text.Length);
                    builder.Append(first);
                    builder.Append(text);
                }

                hasNode = true;
            }

            return builder?.ToString() ?? first ?? string.Empty;
        }

        internal static string GetRunText(OpenXmlElement parent) {
            string? first = null;
            StringBuilder? builder = null;

            foreach (var run in parent.Elements<Run>()) {
                string text = run.Text?.Text ?? string.Empty;
                if (builder != null) {
                    builder.Append(text);
                } else if (first == null) {
                    first = text;
                } else {
                    builder = new StringBuilder(first.Length + text.Length);
                    builder.Append(first);
                    builder.Append(text);
                }
            }

            return builder?.ToString() ?? first ?? string.Empty;
        }

        public string? Get(int index) {
            var items = GetLoadedItems();
            if ((uint)index < (uint)items.Count) return items[index];
            return null;
        }

        internal void EnsureLoaded() {
            _ = GetLoadedItems();
        }

        internal HashSet<int>? FindIndexesContaining(string text, StringComparison comparison) {
            if (string.IsNullOrEmpty(text)) {
                return null;
            }

            var key = (text, comparison);
            lock (_containsCacheLock) {
                if (_containsCache != null && _containsCache.TryGetValue(key, out var cachedIndexes)) {
                    return cachedIndexes;
                }
            }

            var items = GetLoadedItems();
            HashSet<int>? indexes = null;
            for (int i = 0; i < items.Count; i++) {
                if (items[i].IndexOf(text, comparison) >= 0) {
                    indexes ??= new HashSet<int>();
                    indexes.Add(i);
                }
            }

            lock (_containsCacheLock) {
                _containsCache ??= new Dictionary<(string Text, StringComparison Comparison), HashSet<int>?>();
                if (_containsCache.Count >= 32) {
                    _containsCache.Clear();
                }

                _containsCache[key] = indexes;
            }

            return indexes;
        }

        private List<string> GetLoadedItems() {
            return _loadedItems ??= _items.Value;
        }

        private static int ParsePositiveIntAttribute(string? value) {
            if (value == null || value.Length == 0) {
                return 0;
            }

            string text = value;
            int parsed = 0;
            for (int i = 0; i < text.Length; i++) {
                int digit = text[i] - '0';
                if ((uint)digit > 9U) {
                    return 0;
                }

                if (parsed > (int.MaxValue - digit) / 10) {
                    return 0;
                }

                parsed = (parsed * 10) + digit;
            }

            return parsed;
        }
    }
}

