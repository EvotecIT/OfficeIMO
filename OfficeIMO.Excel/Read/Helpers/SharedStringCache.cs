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

        private readonly Lazy<SharedStringTablePart?> _part;
        private readonly bool _preferDom;
        private readonly int _maxSharedStringItems;
        private readonly int _maxSharedStringItemCharacters;
        private readonly long _maxSharedStringCharacters;
        private readonly Lazy<List<string>> _items;
        private List<string>? _loadedItems;
        private readonly object _containsCacheLock = new object();
        private Dictionary<(string Text, StringComparison Comparison), HashSet<int>?>? _containsCache;

        private SharedStringCache(WorkbookPart? workbookPart, bool preferDom, ExcelReadOptions options) {
            _part = new Lazy<SharedStringTablePart?>(() => workbookPart?.SharedStringTablePart, LazyThreadSafetyMode.ExecutionAndPublication);
            _preferDom = preferDom;
            _maxSharedStringItems = options.MaxSharedStringItems;
            _maxSharedStringItemCharacters = options.MaxSharedStringItemCharacters;
            _maxSharedStringCharacters = options.MaxSharedStringCharacters;
            _items = new Lazy<List<string>>(LoadItems, LazyThreadSafetyMode.ExecutionAndPublication);
        }

        public static SharedStringCache Build(SpreadsheetDocument doc, ExcelReadOptions? options = null) {
            return new SharedStringCache(doc.WorkbookPart, doc.FileOpenAccess != FileAccess.Read, options ?? new ExcelReadOptions());
        }

        private List<string> LoadItems() {
            SharedStringTablePart? part = GetSharedStringTablePart();
            if (part == null) return new List<string>();
            if (_preferDom && part.SharedStringTable != null) {
                return LoadItemsFromDom();
            }

            if (TryLoadItemsXmlFast(part, out var items)) {
                return items;
            }

            return LoadItemsFromDom();
        }

        private List<string> LoadItemsFromDom() {
            var part = GetSharedStringTablePart();
            if (part == null || part.SharedStringTable == null) return new List<string>();
            var table = part.SharedStringTable;
            var list = new List<string>(GetBoundedCapacity(table.UniqueCount?.Value, table.Count?.Value));
            long totalCharacters = 0;
            foreach (var item in table.Elements<SharedStringItem>()) {
                EnsureCanAddSharedString(list);
                string value;
                if (item.Text?.Text != null) {
                    value = item.Text.Text;
                } else if (item.HasChildren) {
                    value = GetRunText(item, _maxSharedStringItemCharacters);
                } else {
                    value = string.Empty;
                }

                ValidateSharedStringText(value, ref totalCharacters);
                list.Add(value);
            }

            return list;
        }

        private SharedStringTablePart? GetSharedStringTablePart() {
            return _part.Value;
        }

        private bool TryLoadItemsXmlFast(SharedStringTablePart part, out List<string> items) {
            items = new List<string>();
            long totalCharacters = 0;

            try {
                using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
                using var reader = XmlReader.Create(stream, SharedStringXmlReaderSettings);
                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element) {
                        continue;
                    }

                    if (reader.LocalName == "sst") {
                        int capacity = GetBoundedCapacity(ParsePositiveLongAttribute(reader.GetAttribute("uniqueCount")));
                        if (capacity <= 0) {
                            capacity = GetBoundedCapacity(ParsePositiveLongAttribute(reader.GetAttribute("count")));
                        }

                        if (capacity > 0) {
                            items.Capacity = capacity;
                        }

                        continue;
                    }

                    if (reader.LocalName == "si") {
                        EnsureCanAddSharedString(items);
                        string value = ReadSharedStringItemXml(reader, _maxSharedStringItemCharacters);
                        ValidateSharedStringText(value, ref totalCharacters);
                        items.Add(value);
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

        private static string ReadSharedStringItemXml(XmlReader reader, int maxItemCharacters) {
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
                EnsureItemCharacterBudget(0, first.Length, maxItemCharacters);
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
                    EnsureItemCharacterBudget(builder.Length, text.Length, maxItemCharacters);
                    builder.Append(text);
                } else if (first == null) {
                    EnsureItemCharacterBudget(0, text.Length, maxItemCharacters);
                    first = text;
                } else {
                    EnsureItemCharacterBudget(first.Length, text.Length, maxItemCharacters);
                    builder = new StringBuilder(first.Length + text.Length);
                    builder.Append(first);
                    builder.Append(text);
                }

                hasNode = true;
            }

            return builder?.ToString() ?? first ?? string.Empty;
        }

        internal static string GetRunText(OpenXmlElement parent) {
            return GetRunText(parent, int.MaxValue);
        }

        private static string GetRunText(OpenXmlElement parent, int maxItemCharacters) {
            string? first = null;
            StringBuilder? builder = null;

            foreach (var run in parent.Elements<Run>()) {
                string text = run.Text?.Text ?? string.Empty;
                if (builder != null) {
                    EnsureItemCharacterBudget(builder.Length, text.Length, maxItemCharacters);
                    builder.Append(text);
                } else if (first == null) {
                    EnsureItemCharacterBudget(0, text.Length, maxItemCharacters);
                    first = text;
                } else {
                    EnsureItemCharacterBudget(first.Length, text.Length, maxItemCharacters);
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

        internal List<string> GetItems() {
            return GetLoadedItems();
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

        private int GetBoundedCapacity(uint? uniqueCount, uint? count) {
            if (uniqueCount.HasValue && uniqueCount.Value > 0U) {
                return GetBoundedCapacity(uniqueCount.Value);
            }

            return count.HasValue && count.Value > 0U ? GetBoundedCapacity(count.Value) : 0;
        }

        private int GetBoundedCapacity(long declaredCount) {
            if (declaredCount <= 0) {
                return 0;
            }

            return declaredCount > _maxSharedStringItems ? _maxSharedStringItems : (int)declaredCount;
        }

        private void EnsureCanAddSharedString(List<string> items) {
            if (items.Count >= _maxSharedStringItems) {
                throw new InvalidDataException($"Shared string table exceeds the configured limit of {_maxSharedStringItems} entries.");
            }
        }

        private void ValidateSharedStringText(string value, ref long totalCharacters) {
            if (value.Length > _maxSharedStringItemCharacters) {
                throw new InvalidDataException($"Shared string item exceeds the configured limit of {_maxSharedStringItemCharacters} characters.");
            }

            if (totalCharacters > _maxSharedStringCharacters - value.Length) {
                throw new InvalidDataException($"Shared string table exceeds the configured aggregate limit of {_maxSharedStringCharacters} characters.");
            }

            totalCharacters += value.Length;
        }

        private static void EnsureItemCharacterBudget(int currentLength, int additionalLength, int maxItemCharacters) {
            if ((long)currentLength + additionalLength > maxItemCharacters) {
                throw new InvalidDataException($"Shared string item exceeds the configured limit of {maxItemCharacters} characters.");
            }
        }

        private static long ParsePositiveLongAttribute(string? value) {
            if (value == null || value.Length == 0) {
                return 0;
            }

            string text = value;
            long parsed = 0;
            for (int i = 0; i < text.Length; i++) {
                int digit = text[i] - '0';
                if ((uint)digit > 9U) {
                    return 0;
                }

                if (parsed > (long.MaxValue - digit) / 10) {
                    return long.MaxValue;
                }

                parsed = (parsed * 10) + digit;
            }

            return parsed;
        }
    }
}

