using OfficeIMO.Excel.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using System.Text;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class SharedStringCache : IDisposable {
        internal const int Utf8CacheSlotCount = ExcelUtf8TextCache.SlotCount;
        internal const int MaximumCachedUtf8ItemBytes = ExcelUtf8TextCache.MaximumCachedItemBytes;
        internal const int MaximumIndexedPartBytes = 64 * 1024 * 1024;
        private const int MaximumRetainedRunTextCapacity = 64 * 1024;
        private static readonly XmlReaderSettings SharedStringXmlReaderSettings = CreateSharedStringXmlReaderSettings();
        private static readonly XmlReaderSettings SharedStringWhitespaceXmlReaderSettings = CreateSharedStringWhitespaceXmlReaderSettings();

        private readonly ExcelPartRef _workbook;
        private readonly Func<Stream>? _openPartStream;
        private readonly bool _preferDom;
        private readonly bool _reportMalformedXml;
        private readonly bool _keepUnmarkedWhitespace;
        private readonly int _maxSharedStringItems;
        private readonly int _maxSharedStringItemCharacters;
        private readonly long _maxSharedStringCharacters;
        private CancellationToken _cancellationToken;
        private readonly Lazy<bool> _initialized;
        private readonly object _lifetimeLock = new object();
        private List<string>? _loadedItems;
        private IndexedAsciiItems? _indexedItems;
        private volatile bool _disposed;
        private readonly object _containsCacheLock = new object();
        private Dictionary<(string Text, StringComparison Comparison), HashSet<int>?>? _containsCache;
        private ExcelUtf8TextCache? _utf8Cache;

        private SharedStringCache(ExcelPartRef workbook, ExcelReaderOptions options) {
            _workbook = workbook;
            _preferDom = !workbook.IsReadOnly;
            _maxSharedStringItems = options.MaxSharedStringItems;
            _maxSharedStringItemCharacters = options.MaxSharedStringItemCharacters;
            _maxSharedStringCharacters = options.MaxSharedStringCharacters;
            _cancellationToken = options.CancellationToken;
            _initialized = new Lazy<bool>(Initialize, LazyThreadSafetyMode.ExecutionAndPublication);
        }

        private SharedStringCache(Func<Stream> openPartStream, ExcelReaderOptions options, bool keepUnmarkedWhitespace = false, bool reportMalformedXml = false) {
            _openPartStream = openPartStream ?? throw new ArgumentNullException(nameof(openPartStream));
            _keepUnmarkedWhitespace = keepUnmarkedWhitespace;
            _reportMalformedXml = reportMalformedXml;
            _preferDom = false;
            _maxSharedStringItems = options.MaxSharedStringItems;
            _maxSharedStringItemCharacters = options.MaxSharedStringItemCharacters;
            _maxSharedStringCharacters = options.MaxSharedStringCharacters;
            _cancellationToken = options.CancellationToken;
            _initialized = new Lazy<bool>(Initialize, LazyThreadSafetyMode.ExecutionAndPublication);
        }

        private SharedStringCache(List<string> items, ExcelReaderOptions options) {
            _openPartStream = null;
            _preferDom = false;
            _maxSharedStringItems = options.MaxSharedStringItems;
            _maxSharedStringItemCharacters = options.MaxSharedStringItemCharacters;
            _maxSharedStringCharacters = options.MaxSharedStringCharacters;
            _cancellationToken = options.CancellationToken;
            _loadedItems = items ?? throw new ArgumentNullException(nameof(items));
            _initialized = new Lazy<bool>(() => true, LazyThreadSafetyMode.ExecutionAndPublication);
        }

        public static SharedStringCache Build(object document, ExcelReaderOptions? options = null) =>
            Build(ExcelPartRef.FromSdk(document), options ?? new ExcelReaderOptions());

        internal static SharedStringCache Build(ExcelPartRef workbook, ExcelReaderOptions options) =>
            new(workbook, options);

        internal static SharedStringCache Build(Func<Stream> openPartStream, ExcelReaderOptions options) =>
            new(openPartStream, options);

        /// <summary>Builds the table from the part's bytes with the text the SDK model gives an editable document.</summary>
        internal static SharedStringCache BuildKeepingUnmarkedWhitespace(Func<Stream> openPartStream, ExcelReaderOptions options) =>
            new(openPartStream, options, keepUnmarkedWhitespace: true);

        internal static SharedStringCache BuildForPackage(Func<Stream> openPartStream, ExcelReaderOptions options, bool keepUnmarkedWhitespace) =>
            new(openPartStream, options, keepUnmarkedWhitespace, reportMalformedXml: true);

        internal static SharedStringCache Empty(ExcelReaderOptions options) =>
            new(new List<string>(), options);

        internal bool HasSameLimits(ExcelReaderOptions options) =>
            _maxSharedStringItems == options.MaxSharedStringItems
            && _maxSharedStringItemCharacters == options.MaxSharedStringItemCharacters
            && _maxSharedStringCharacters == options.MaxSharedStringCharacters;

        /// <summary>
        /// Parks a completed parse for another reader. An unread table is not worth keeping, and the
        /// finished reader's token must not reach its successor.
        /// </summary>
        internal bool TryDetachCancellationAfterLoad() {
            lock (_lifetimeLock) {
                if (_disposed || (_loadedItems == null && _indexedItems == null)) return false;
                _cancellationToken = CancellationToken.None;
                return true;
            }
        }

        private bool Initialize() {
            // Disposal and initial loading share a gate so a late load cannot retain
            // a package buffer after its workbook has closed.
            lock (_lifetimeLock) {
                ThrowIfDisposed();
                _loadedItems = LoadItems();
                return true;
            }
        }

        private List<string>? LoadItems() {
            _cancellationToken.ThrowIfCancellationRequested();
            if (_openPartStream != null) {
                Stream? stream = _openPartStream();
                try {
                    if (stream is OpenXmlPooledPartStream pooled
                        && TryIndexSimpleAsciiItems(pooled, out IndexedAsciiItems? indexedItems)) {
                        _indexedItems = indexedItems;
                        stream = null; // The validated index now owns the pooled stream.
                        return null;
                    }
                    if (TryLoadItemsXmlFast(stream, out List<string> nativeItems)) {
                        return nativeItems;
                    }

                    if (_reportMalformedXml) throw new XmlException("The shared-string part has no shared-string table.");
                    throw new XlsxTabularFastPathNotSupportedException(
                        "The shared-string table requires the general workbook reader.");
                } finally {
                    stream?.Dispose();
                }
            }

            return LoadItemsFromPackage();
        }

        private List<string> LoadItemsFromPackage() {
            ExcelPartRef part = _workbook.FindWorkbookPart(ExcelReadWorkbookParts.SharedStrings);
            if (part.IsNone) return new List<string>();

            if (!_preferDom) {
                using Stream stream = part.OpenRead();
                if (TryLoadItemsXmlFast(stream, out List<string> items)) return items;
            }

            OfficeXmlNode table = part.Root;
            if (table.IsNone) return new List<string>();
            const int spreadsheet = OfficeXmlNamespaces.SpreadsheetML;
            var list = new List<string>(GetBoundedCapacity(table.UInt32Attribute("uniqueCount"), table.UInt32Attribute("count")));
            long totalCharacters = 0;
            StringBuilder? runTextBuffer = null;
            foreach (OfficeXmlNode item in table.Children(spreadsheet, "si")) {
                CheckCancellation(list.Count);
                EnsureCanAddSharedString(list);
                OfficeXmlNode text = item.Child(spreadsheet, "t");
                string value = text.Exists ? Utilities.ExcelSanitizer.DecodeXString(text.InnerText)
                    : GetRunText(item, _maxSharedStringItemCharacters, ref runTextBuffer);
                ValidateSharedStringText(value, ref totalCharacters);
                list.Add(value);
            }

            return list;
        }

        private bool TryLoadItemsXmlFast(Stream stream, out List<string> items) {
            items = new List<string>();
            long totalCharacters = 0;
            StringBuilder? runTextBuffer = null;

            try {
                using var reader = XmlReader.Create(stream, OpenXmlReadNameTable.WithSchemaNames(
                    _keepUnmarkedWhitespace ? SharedStringWhitespaceXmlReaderSettings : SharedStringXmlReaderSettings));
                int nodesRead = 0;
                int tableDepth = -1;
                while (reader.Read()) {
                    CheckCancellation(nodesRead++);
                    if (reader.NodeType != XmlNodeType.Element) {
                        continue;
                    }

                    if (tableDepth < 0) {
                        if (reader.Depth != 0 || !SpreadsheetXmlContent.IsSpreadsheetElement(reader, "sst")) {
                            items = null!;
                            return false;
                        }

                        tableDepth = reader.Depth;
                        int capacity = GetBoundedCapacity(ParsePositiveLongAttribute(reader.GetAttribute("uniqueCount")));
                        if (capacity <= 0) {
                            capacity = GetBoundedCapacity(ParsePositiveLongAttribute(reader.GetAttribute("count")));
                        }

                        if (capacity > 0) {
                            items.Capacity = capacity;
                        }

                        continue;
                    }

                    if (SpreadsheetXmlContent.IsDirectChildElement(reader, tableDepth, "si")) {
                        EnsureCanAddSharedString(items);
                        string value = ReadSharedStringItemXml(
                            reader,
                            _maxSharedStringItemCharacters,
                            _cancellationToken,
                            ref runTextBuffer);
                        ValidateSharedStringText(value, ref totalCharacters);
                        items.Add(value);
                    }
                }

                if (tableDepth >= 0) return true;
                items = null!;
                return false;
            } catch (XmlException) when (!_reportMalformedXml) {
                items = null!;
                return false;
            } catch (IOException exception) when (!ExcelReadLimitFailure.Is(exception)) {
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

        private static XmlReaderSettings CreateSharedStringWhitespaceXmlReaderSettings() {
            XmlReaderSettings settings = CreateSharedStringXmlReaderSettings();
            settings.IgnoreWhitespace = false;
            return settings;
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

        private static string ReadSharedStringItemXml(
            XmlReader reader,
            int maxItemCharacters,
            CancellationToken cancellationToken,
            ref StringBuilder? runTextBuffer) {
            if (reader.IsEmptyElement) {
                return string.Empty;
            }

            int depth = reader.Depth;
            string? first = null;
            StringBuilder? builder = null;
            int richRunDepth = -1;

            bool hasNode = reader.Read();
            int nodesRead = 0;
            if (hasNode && SpreadsheetXmlContent.IsDirectChildElement(reader, depth, "t")) {
                first = Utilities.ExcelSanitizer.DecodeXString(reader.ReadElementContentAsString());
                EnsureItemCharacterBudget(0, first.Length, maxItemCharacters);
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == "si") {
                    return first;
                }

                hasNode = true;
            }

            while (hasNode) {
                if ((nodesRead++ & 255) == 0) {
                    cancellationToken.ThrowIfCancellationRequested();
                }
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == "si") {
                    break;
                }

                if (!SpreadsheetXmlContent.IsRichTextElement(reader, depth, ref richRunDepth)) {
                    hasNode = reader.Read();
                    continue;
                }

                string text = Utilities.ExcelSanitizer.DecodeXString(reader.ReadElementContentAsString());
                if (builder != null) {
                    EnsureItemCharacterBudget(builder.Length, text.Length, maxItemCharacters);
                    builder.Append(text);
                } else if (first == null) {
                    EnsureItemCharacterBudget(0, text.Length, maxItemCharacters);
                    first = text;
                } else {
                    EnsureItemCharacterBudget(first.Length, text.Length, maxItemCharacters);
                    builder = runTextBuffer ??= new StringBuilder(first.Length + text.Length);
                    builder.Clear();
                    builder.Append(first);
                    builder.Append(text);
                }

                hasNode = true;
            }

            return CompleteRunText(builder, first, ref runTextBuffer);
        }

        private void CheckCancellation(int iteration) {
            if ((iteration & 1023) == 0) {
                _cancellationToken.ThrowIfCancellationRequested();
            }
        }

        internal static string GetRunText(object parent) => GetRunText(OfficeXmlNode.FromSdk(parent));

        internal static string GetRunText(OfficeXmlNode parent) {
            StringBuilder? runTextBuffer = null;
            return GetRunText(parent, int.MaxValue, ref runTextBuffer);
        }

        private static string GetRunText(OfficeXmlNode parent, int maxItemCharacters, ref StringBuilder? runTextBuffer) {
            string? first = null;
            StringBuilder? builder = null;

            foreach (OfficeXmlNode run in parent.Children(OfficeXmlNamespaces.SpreadsheetML, "r")) {
                string text = Utilities.ExcelSanitizer.DecodeXString(run.Child(OfficeXmlNamespaces.SpreadsheetML, "t").InnerText);
                if (builder != null) {
                    EnsureItemCharacterBudget(builder.Length, text.Length, maxItemCharacters);
                    builder.Append(text);
                } else if (first == null) {
                    EnsureItemCharacterBudget(0, text.Length, maxItemCharacters);
                    first = text;
                } else {
                    EnsureItemCharacterBudget(first.Length, text.Length, maxItemCharacters);
                    builder = runTextBuffer ??= new StringBuilder(first.Length + text.Length);
                    builder.Clear();
                    builder.Append(first);
                    builder.Append(text);
                }
            }

            return CompleteRunText(builder, first, ref runTextBuffer);
        }

        // The buffer belongs to one table load, not the cache or individual items.
        // Keep each resulting string independent and avoid retaining an unusually large run.
        private static string CompleteRunText(StringBuilder? builder, string? first, ref StringBuilder? runTextBuffer) {
            string value = builder?.ToString() ?? first ?? string.Empty;
            if (runTextBuffer?.Capacity > MaximumRetainedRunTextCapacity) runTextBuffer = null;
            return value;
        }

        public string? Get(int index) {
            EnsureLoaded();
            if (_indexedItems != null) return _indexedItems.Get(index);
            var items = _loadedItems!;
            if ((uint)index < (uint)items.Count) return items[index];
            return null;
        }

        internal bool TryGetUtf8(int index, out ArraySegment<byte> value) {
            EnsureLoaded();
            if (_indexedItems != null) return _indexedItems.TryGetUtf8(index, out value);
            var items = _loadedItems!;
            if ((uint)index >= (uint)items.Count) {
                value = default;
                return false;
            }

            ExcelUtf8TextCache utf8Cache = Volatile.Read(ref _utf8Cache) ?? InitializeUtf8Cache();
            value = utf8Cache.Get(index, items[index]);
            return true;
        }

        private ExcelUtf8TextCache InitializeUtf8Cache() {
            var cache = new ExcelUtf8TextCache();
            return Interlocked.CompareExchange(ref _utf8Cache, cache, null) ?? cache;
        }

        internal int Count {
            get {
                EnsureLoaded();
                return _indexedItems?.Count ?? _loadedItems!.Count;
            }
        }

        /// <summary>True after the table has loaded successfully; background readers never trigger loading.</summary>
        internal bool IsLoaded => _initialized.IsValueCreated;

        internal void EnsureLoaded() {
            _ = _initialized.Value;
            ThrowIfDisposed();
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

            int count = Count;
            IndexedAsciiItems? indexedItems = _indexedItems;
            List<string>? loadedItems = _loadedItems;
            HashSet<int>? indexes = null;
            for (int i = 0; i < count; i++) {
                string value = indexedItems != null ? indexedItems.Get(i)! : loadedItems![i];
                if (value.IndexOf(text, comparison) >= 0) {
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

        public void Dispose() {
            lock (_lifetimeLock) {
                if (_disposed) return;
                _disposed = true;
                _indexedItems?.Dispose();
                _indexedItems = null;
                _loadedItems = null;
                _utf8Cache = null;
                _containsCache = null;
            }
        }

        private void ThrowIfDisposed() {
            if (_disposed) throw new ObjectDisposedException(nameof(SharedStringCache));
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

        private void EnsureCanAddSharedString(List<string> items) => EnsureCanAddSharedString(items.Count);

        private void EnsureCanAddSharedString(int itemCount) {
            if (itemCount >= _maxSharedStringItems) {
                throw ExcelReadLimitFailure.Create(nameof(ExcelReaderOptions.MaxSharedStringItems), _maxSharedStringItems, $"Shared string table exceeds the configured limit of {_maxSharedStringItems} entries.");
            }
        }

        private void ValidateSharedStringText(string value, ref long totalCharacters) {
            ValidateSharedStringLength(value.Length, ref totalCharacters);
        }

        private void ValidateSharedStringLength(int length, ref long totalCharacters) {
            if (length > _maxSharedStringItemCharacters) {
                throw ExcelReadLimitFailure.Create(nameof(ExcelReaderOptions.MaxSharedStringItemCharacters), _maxSharedStringItemCharacters, $"Shared string item exceeds the configured limit of {_maxSharedStringItemCharacters} characters.");
            }

            if (totalCharacters > _maxSharedStringCharacters - length) {
                throw ExcelReadLimitFailure.Create(nameof(ExcelReaderOptions.MaxSharedStringCharacters), _maxSharedStringCharacters, $"Shared string table exceeds the configured aggregate limit of {_maxSharedStringCharacters} characters.");
            }

            totalCharacters += length;
        }

        private static void EnsureItemCharacterBudget(int currentLength, int additionalLength, int maxItemCharacters) {
            if ((long)currentLength + additionalLength > maxItemCharacters) {
                throw ExcelReadLimitFailure.Create(nameof(ExcelReaderOptions.MaxSharedStringItemCharacters), maxItemCharacters, $"Shared string item exceeds the configured limit of {maxItemCharacters} characters.");
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

