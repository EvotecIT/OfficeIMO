using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed class StylesCacheProvider {
        private readonly SpreadsheetDocument? _doc;
        private readonly Func<Stream>? _openPartStream;
        private readonly object _gate = new();
        private StylesCache? _value;

        public StylesCacheProvider(SpreadsheetDocument doc) {
            _doc = doc;
            _openPartStream = null;
        }

        internal StylesCacheProvider(Func<Stream> openPartStream) {
            _openPartStream = openPartStream ?? throw new ArgumentNullException(nameof(openPartStream));
        }

        internal StylesCacheProvider(StylesCache value) {
            _value = value ?? throw new ArgumentNullException(nameof(value));
        }

        internal StylesCache GetValue(CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            var value = _value;
            if (value != null) {
                return value;
            }

            lock (_gate) {
                cancellationToken.ThrowIfCancellationRequested();
                if (_value != null) {
                    return _value;
                }

                if (_openPartStream != null) {
                    using Stream stream = _openPartStream();
                    return _value = StylesCache.Build(stream, cancellationToken);
                }

                return _value = StylesCache.Build(_doc!, cancellationToken);
            }
        }
    }

    internal sealed class StylesCache {
        private static readonly bool[] EmptyDateStyleIndexes = Array.Empty<bool>();
        private static readonly XmlReaderSettings StylesXmlReaderSettings = CreateStylesXmlReaderSettings();
        private bool[] _dateStyleIndexes = EmptyDateStyleIndexes;
        private bool[] _dateSystemShiftStyleIndexes = EmptyDateStyleIndexes;

        private StylesCache() { }

        public bool HasDateStyles { get; private set; }

        public int CellFormatCount => _dateStyleIndexes.Length;

        public static StylesCache Build(SpreadsheetDocument doc, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            var cache = new StylesCache();
            var sp = doc.WorkbookPart!.WorkbookStylesPart;
            if (sp == null) return cache;

            if (doc.FileOpenAccess == FileAccess.Read) {
                bool classified = TryBuildXmlFast(sp, cache, cancellationToken);
                if (classified && doc.MarkupCompatibilityProcessSettings.ProcessMode == MarkupCompatibilityProcessMode.NoProcess) {
                    return cache;
                }
                // The XML scan validates the complete part even when SDK processing
                // selects a version-specific branch for classification.
                ClearDateStyles(cache);
            }

            var stylesheet = sp.Stylesheet;
            cancellationToken.ThrowIfCancellationRequested();
            if (stylesheet == null) return cache;

            Dictionary<uint, string>? nf = null;
            var numbering = stylesheet.NumberingFormats;
            if (numbering != null) {
                foreach (var n in numbering.Elements<NumberingFormat>()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (n.NumberFormatId?.Value is uint id && n.FormatCode?.Value is string code) {
                        nf ??= new Dictionary<uint, string>();
                        nf[id] = code;
                    }
                }
            }

            var xfs = stylesheet.CellFormats;
            if (xfs != null) {
                int idx = 0;
                foreach (var cf in xfs.Elements<CellFormat>()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (idx >= cache._dateStyleIndexes.Length) {
                        Array.Resize(ref cache._dateStyleIndexes, Math.Max(4, (idx + 1) * 2));
                    }

                    var nId = (uint)(cf.NumberFormatId?.Value ?? 0);
                    bool dateLike = IsDateNumberFormat(nId, nf);
                    if (dateLike) {
                        cache._dateStyleIndexes[idx] = true;
                        cache.HasDateStyles = true;
                    }

                    if (IsDateSystemShiftNumberFormat(nId, nf)) {
                        if (idx >= cache._dateSystemShiftStyleIndexes.Length) {
                            Array.Resize(ref cache._dateSystemShiftStyleIndexes, Math.Max(4, (idx + 1) * 2));
                        }

                        cache._dateSystemShiftStyleIndexes[idx] = true;
                    }

                    idx++;
                }

                if (idx == 0) {
                    cache._dateStyleIndexes = EmptyDateStyleIndexes;
                    cache._dateSystemShiftStyleIndexes = EmptyDateStyleIndexes;
                } else if (idx < cache._dateStyleIndexes.Length) {
                    Array.Resize(ref cache._dateStyleIndexes, idx);
                    if (idx < cache._dateSystemShiftStyleIndexes.Length) {
                        Array.Resize(ref cache._dateSystemShiftStyleIndexes, idx);
                    }
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            return cache;
        }

        internal static StylesCache Build(Stream stream, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            var cache = new StylesCache();
            if (!TryBuildXmlFast(stream, cache, cancellationToken)) {
                throw new XlsxTabularFastPathNotSupportedException(
                    "The styles part requires the Open XML SDK fallback path.");
            }

            return cache;
        }

        internal static StylesCache Empty() => new();

        public bool IsDateLike(uint styleIndex) => styleIndex < (uint)_dateStyleIndexes.Length && _dateStyleIndexes[styleIndex];

        public bool IsDateSystemShiftStyle(uint styleIndex) => styleIndex < (uint)_dateSystemShiftStyleIndexes.Length && _dateSystemShiftStyleIndexes[styleIndex];

        private static bool TryBuildXmlFast(WorkbookStylesPart sp, StylesCache cache, CancellationToken cancellationToken) {
            using var stream = sp.GetStream(FileMode.Open, FileAccess.Read);
            return TryBuildXmlFast(stream, cache, cancellationToken);
        }

        private static bool TryBuildXmlFast(Stream stream, StylesCache cache, CancellationToken cancellationToken) {
            try {
                using var reader = XmlReader.Create(stream,
                    OpenXmlReadNameTable.WithSchemaNames(StylesXmlReaderSettings));
                reader.MoveToContent();
                cancellationToken.ThrowIfCancellationRequested();
                if (!SpreadsheetXmlContent.IsSpreadsheetElement(reader, "styleSheet")) return false;
                int childDepth = reader.Depth + 1;
                Dictionary<uint, string>? nf = null;
                List<(int Depth, string Namespace, string Name)>? processedContent = null;
                List<(int Depth, bool Selected)>? alternateContent = null;
                TrackProcessedStyleContent(reader, ref processedContent);
                while (reader.Read()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (reader.NodeType == XmlNodeType.EndElement) {
                        childDepth = Math.Min(childDepth, reader.Depth);
                        continue;
                    }
                    if (reader.NodeType != XmlNodeType.Element) continue;
                    TrackProcessedStyleContent(reader, ref processedContent);
                    if (reader.Depth != childDepth) continue;
                    if (IsCompatibilityStyleContainer(reader, processedContent, ref alternateContent)) {
                        if (!reader.IsEmptyElement) childDepth++;
                        continue;
                    }
                    if (SpreadsheetXmlContent.IsSpreadsheetElement(reader, "numFmts")) {
                        ReadNumberFormatsXml(reader, ref nf, ref processedContent, ref alternateContent, cancellationToken);
                        continue;
                    }
                    if (SpreadsheetXmlContent.IsSpreadsheetElement(reader, "cellXfs")) {
                        ReadCellFormatsXml(reader, cache, nf, ref processedContent, ref alternateContent, cancellationToken);
                        // Validate the remaining XML and declared package length
                        // before publishing the completed style map.
                        while (reader.Read()) cancellationToken.ThrowIfCancellationRequested();
                        cancellationToken.ThrowIfCancellationRequested();
                        return true;
                    }
                }

                ClearDateStyles(cache);
                return true;
            } catch (IOException) {
                ClearDateStyles(cache);
                return false;
            } catch (UnauthorizedAccessException) {
                ClearDateStyles(cache);
                return false;
            } catch (ObjectDisposedException) {
                ClearDateStyles(cache);
                return false;
            }
        }

        private const string CompatibilityNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        /// <summary>Retains inherited process-content names using their declaration's namespace bindings.</summary>
        private static void TrackProcessedStyleContent(XmlReader reader,
            ref List<(int Depth, string Namespace, string Name)>? processedContent) {
            if (processedContent != null) {
                while (processedContent.Count > 0 && processedContent[processedContent.Count - 1].Depth >= reader.Depth) {
                    processedContent.RemoveAt(processedContent.Count - 1);
                }
            }
            string? names = reader.GetAttribute("ProcessContent", CompatibilityNamespace);
            if (string.IsNullOrEmpty(names)) return;
            foreach (string name in names!.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)) {
                int colon = name.IndexOf(':');
                if (colon <= 0 || colon == name.Length - 1 || name.IndexOf(':', colon + 1) >= 0) continue;
                string? namespaceUri = reader.LookupNamespace(name.Substring(0, colon));
                if (string.IsNullOrEmpty(namespaceUri)) continue;
                processedContent ??= new List<(int Depth, string Namespace, string Name)>();
                processedContent.Add((reader.Depth, namespaceUri!, name.Substring(colon + 1)));
            }
        }

        /// <summary>Exposes only compatibility wrappers that preserve the owning style container.</summary>
        private static bool IsCompatibilityStyleContainer(XmlReader reader,
            List<(int Depth, string Namespace, string Name)>? processedContent,
            ref List<(int Depth, bool Selected)>? alternateContent) {
            if (alternateContent != null) {
                while (alternateContent.Count > 0 && alternateContent[alternateContent.Count - 1].Depth >= reader.Depth) {
                    alternateContent.RemoveAt(alternateContent.Count - 1);
                }
            }
            if (reader.NamespaceURI == CompatibilityNamespace) {
                if (reader.LocalName == "AlternateContent") {
                    if (!reader.IsEmptyElement) {
                        alternateContent ??= new List<(int Depth, bool Selected)>();
                        alternateContent.Add((reader.Depth, false));
                    }
                    return true;
                }
                if (reader.LocalName is "Choice" or "Fallback") {
                    if (alternateContent == null || alternateContent.Count == 0) return false;
                    int index = alternateContent.Count - 1;
                    var branch = alternateContent[index];
                    if (reader.Depth != branch.Depth + 1 || branch.Selected
                        || reader.LocalName == "Choice" && !UnderstandsStyleRequirements(reader)) return false;
                    alternateContent[index] = (branch.Depth, true);
                    return true;
                }
            }
            if (processedContent == null || SpreadsheetXmlContent.IsSpreadsheetElement(reader, reader.LocalName)) return false;
            foreach (var name in processedContent) {
                if (name.Namespace == reader.NamespaceURI && (name.Name == reader.LocalName || name.Name == "*")) return true;
            }
            return false;
        }

        /// <summary>Selects a choice only when every required namespace is understood by this reader.</summary>
        private static bool UnderstandsStyleRequirements(XmlReader reader) {
            string? requires = reader.GetAttribute("Requires");
            if (string.IsNullOrWhiteSpace(requires)) return false;
            foreach (string prefix in requires!.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)) {
                string? namespaceUri = reader.LookupNamespace(prefix);
                if (namespaceUri != SpreadsheetXmlContent.SpreadsheetNamespace
                    && namespaceUri != SpreadsheetXmlContent.StrictSpreadsheetNamespace
                    && namespaceUri != CompatibilityNamespace
                    && namespaceUri != "http://www.w3.org/XML/1998/namespace") return false;
            }
            return true;
        }

        private static void ReadNumberFormatsXml(XmlReader reader,
            ref Dictionary<uint, string>? numberingFormats,
            ref List<(int Depth, string Namespace, string Name)>? processedContent,
            ref List<(int Depth, bool Selected)>? alternateContent, CancellationToken cancellationToken) {
            if (reader.IsEmptyElement) return;
            int depth = reader.Depth;
            int childDepth = depth + 1;
            while (reader.Read()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.NodeType == XmlNodeType.EndElement) {
                    if (reader.Depth == depth) break;
                    childDepth = Math.Min(childDepth, reader.Depth);
                    continue;
                }
                if (reader.NodeType != XmlNodeType.Element) continue;
                TrackProcessedStyleContent(reader, ref processedContent);
                if (reader.Depth != childDepth) continue;
                if (IsCompatibilityStyleContainer(reader, processedContent, ref alternateContent)) {
                    if (!reader.IsEmptyElement) childDepth++;
                    continue;
                }
                if (SpreadsheetXmlContent.IsSpreadsheetElement(reader, "numFmt")
                    && TryParseUIntAttribute(reader.GetAttribute("numFmtId"), out uint id)
                    && reader.GetAttribute("formatCode") is string code) {
                    numberingFormats ??= new Dictionary<uint, string>();
                    numberingFormats[id] = code;
                }
            }
        }

        /// <summary>Removes partial XML classifications before an empty result or DOM fallback.</summary>
        private static void ClearDateStyles(StylesCache cache) {
            cache._dateStyleIndexes = EmptyDateStyleIndexes;
            cache._dateSystemShiftStyleIndexes = EmptyDateStyleIndexes;
            cache.HasDateStyles = false;
        }

        private static void ReadCellFormatsXml(XmlReader reader, StylesCache cache,
            Dictionary<uint, string>? numberingFormats,
            ref List<(int Depth, string Namespace, string Name)>? processedContent,
            ref List<(int Depth, bool Selected)>? alternateContent, CancellationToken cancellationToken) {
            if (reader.IsEmptyElement) {
                ClearDateStyles(cache);
                return;
            }

            int depth = reader.Depth;
            int childDepth = depth + 1;
            int index = 0;
            while (reader.Read()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.NodeType == XmlNodeType.EndElement) {
                    if (reader.Depth == depth) break;
                    childDepth = Math.Min(childDepth, reader.Depth);
                    continue;
                }
                if (reader.NodeType != XmlNodeType.Element) continue;
                TrackProcessedStyleContent(reader, ref processedContent);
                if (reader.Depth != childDepth) continue;
                if (IsCompatibilityStyleContainer(reader, processedContent, ref alternateContent)) {
                    if (!reader.IsEmptyElement) childDepth++;
                    continue;
                }
                if (SpreadsheetXmlContent.IsSpreadsheetElement(reader, "xf")) {
                    if (index == cache._dateStyleIndexes.Length) {
                        int newSize = Math.Max(4, index * 2);
                        Array.Resize(ref cache._dateStyleIndexes, newSize);
                        Array.Resize(ref cache._dateSystemShiftStyleIndexes, newSize);
                    }
                    if (TryParseUIntAttribute(reader.GetAttribute("numFmtId"), out uint numberFormatId)) {
                        if (IsDateNumberFormat(numberFormatId, numberingFormats)) {
                            cache._dateStyleIndexes[index] = true;
                            cache.HasDateStyles = true;
                        }
                        if (IsDateSystemShiftNumberFormat(numberFormatId, numberingFormats)) {
                            cache._dateSystemShiftStyleIndexes[index] = true;
                        }
                    }
                    index++;
                }
            }

            if (index == 0) {
                ClearDateStyles(cache);
            } else if (index < cache._dateStyleIndexes.Length) {
                Array.Resize(ref cache._dateStyleIndexes, index);
                if (index < cache._dateSystemShiftStyleIndexes.Length) {
                    Array.Resize(ref cache._dateSystemShiftStyleIndexes, index);
                }
            }
        }

        private static bool IsDateNumberFormat(uint id, Dictionary<uint, string>? numberingFormats) {
            return IsBuiltInDate(id)
                || (numberingFormats != null
                    && numberingFormats.TryGetValue(id, out string? code)
                    && ExcelNumberFormatClassifier.LooksLikeDateFormat(code));
        }

        private static bool IsDateSystemShiftNumberFormat(uint id, Dictionary<uint, string>? numberingFormats) {
            return IsBuiltInDateSystemShiftFormat(id)
                || (numberingFormats != null
                    && numberingFormats.TryGetValue(id, out string? code)
                    && ExcelNumberFormatClassifier.LooksLikeDateSystemFormat(code));
        }

        private static bool IsBuiltInDate(uint id)
            => ExcelBuiltInNumberFormats.IsDate(id);

        private static bool IsBuiltInDateSystemShiftFormat(uint id)
            => ExcelBuiltInNumberFormats.IsDateSystemShift(id);

        private static XmlReaderSettings CreateStylesXmlReaderSettings() {
            return new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true,
                CloseInput = false
            };
        }

        private static bool TryParseUIntAttribute(string? value, out uint result) {
            result = 0;
            if (string.IsNullOrEmpty(value)) {
                return false;
            }

            string text = value!;
            uint parsed = 0;
            for (int i = 0; i < text.Length; i++) {
                int digit = text[i] - '0';
                if ((uint)digit > 9U) {
                    return uint.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out result);
                }

                if (parsed > (uint.MaxValue - (uint)digit) / 10U) {
                    return uint.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out result);
                }

                parsed = (parsed * 10U) + (uint)digit;
            }

            result = parsed;
            return true;
        }

    }
}
