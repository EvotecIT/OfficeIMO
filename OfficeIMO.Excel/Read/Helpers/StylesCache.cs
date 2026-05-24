using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.IO;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed class StylesCacheProvider {
        private readonly SpreadsheetDocument _doc;
        private readonly object _gate = new();
        private StylesCache? _value;

        public StylesCacheProvider(SpreadsheetDocument doc) {
            _doc = doc;
        }

        public StylesCache Value {
            get {
                var value = _value;
                if (value != null) {
                    return value;
                }

                lock (_gate) {
                    return _value ??= StylesCache.Build(_doc);
                }
            }
        }
    }

    internal sealed class StylesCache {
        private static readonly bool[] EmptyDateStyleIndexes = Array.Empty<bool>();
        private static readonly XmlReaderSettings StylesXmlReaderSettings = CreateStylesXmlReaderSettings();
        private bool[] _dateStyleIndexes = EmptyDateStyleIndexes;

        private StylesCache() { }

        public bool HasDateStyles { get; private set; }

        public static StylesCache Build(SpreadsheetDocument doc) {
            var cache = new StylesCache();
            var sp = doc.WorkbookPart!.WorkbookStylesPart;
            if (sp == null) return cache;

            if (doc.FileOpenAccess == FileAccess.Read && TryBuildXmlFast(sp, cache)) {
                return cache;
            }

            if (sp.Stylesheet == null) return cache;

            Dictionary<uint, string>? nf = null;
            var numbering = sp.Stylesheet.NumberingFormats;
            if (numbering != null) {
                foreach (var n in numbering.Elements<NumberingFormat>()) {
                    if (n.NumberFormatId?.Value is uint id && n.FormatCode?.Value is string code) {
                        nf ??= new Dictionary<uint, string>();
                        nf[id] = code;
                    }
                }
            }

            var xfs = sp.Stylesheet.CellFormats;
            if (xfs != null) {
                int expectedCount = xfs.Count?.Value is uint declaredCount && declaredCount <= int.MaxValue
                    ? (int)declaredCount
                    : 0;
                if (expectedCount > 0) {
                    cache._dateStyleIndexes = new bool[expectedCount];
                }

                int idx = 0;
                foreach (var cf in xfs.Elements<CellFormat>()) {
                    if (idx == cache._dateStyleIndexes.Length) {
                        Array.Resize(ref cache._dateStyleIndexes, Math.Max(4, idx * 2));
                    }

                    var nId = (uint)(cf.NumberFormatId?.Value ?? 0);
                    bool dateLike = IsDateNumberFormat(nId, nf);
                    if (dateLike) {
                        cache._dateStyleIndexes[idx] = true;
                        cache.HasDateStyles = true;
                    }

                    idx++;
                }

                if (idx == 0) {
                    cache._dateStyleIndexes = EmptyDateStyleIndexes;
                } else if (idx < cache._dateStyleIndexes.Length) {
                    Array.Resize(ref cache._dateStyleIndexes, idx);
                }
            }

            return cache;
        }

        public bool IsDateLike(uint styleIndex) => styleIndex < (uint)_dateStyleIndexes.Length && _dateStyleIndexes[styleIndex];

        private static bool TryBuildXmlFast(WorkbookStylesPart sp, StylesCache cache) {
            try {
                using var stream = sp.GetStream(FileMode.Open, FileAccess.Read);
                using var reader = XmlReader.Create(stream, StylesXmlReaderSettings);
                Dictionary<uint, string>? nf = null;

                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element) {
                        continue;
                    }

                    if (reader.LocalName == "numFmt") {
                        if (TryParseUIntAttribute(reader.GetAttribute("numFmtId"), out uint id)
                            && reader.GetAttribute("formatCode") is string code) {
                            nf ??= new Dictionary<uint, string>();
                            nf[id] = code;
                        }

                        continue;
                    }

                    if (reader.LocalName == "cellXfs") {
                        ReadCellFormatsXml(reader, cache, nf);
                        return true;
                    }
                }

                cache._dateStyleIndexes = EmptyDateStyleIndexes;
                cache.HasDateStyles = false;
                return true;
            } catch (XmlException) {
                cache._dateStyleIndexes = EmptyDateStyleIndexes;
                cache.HasDateStyles = false;
                return false;
            } catch (IOException) {
                cache._dateStyleIndexes = EmptyDateStyleIndexes;
                cache.HasDateStyles = false;
                return false;
            } catch (UnauthorizedAccessException) {
                cache._dateStyleIndexes = EmptyDateStyleIndexes;
                cache.HasDateStyles = false;
                return false;
            } catch (ObjectDisposedException) {
                cache._dateStyleIndexes = EmptyDateStyleIndexes;
                cache.HasDateStyles = false;
                return false;
            }
        }

        private static void ReadCellFormatsXml(XmlReader reader, StylesCache cache, Dictionary<uint, string>? numberingFormats) {
            int expectedCount = TryParseIntAttribute(reader.GetAttribute("count"), out int parsedCount) ? parsedCount : 0;
            if (expectedCount > 0) {
                cache._dateStyleIndexes = new bool[expectedCount];
            }

            if (reader.IsEmptyElement) {
                cache._dateStyleIndexes = EmptyDateStyleIndexes;
                cache.HasDateStyles = false;
                return;
            }

            int depth = reader.Depth;
            int index = 0;
            while (reader.Read()) {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == "cellXfs") {
                    break;
                }

                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "xf") {
                    continue;
                }

                if (index == cache._dateStyleIndexes.Length) {
                    Array.Resize(ref cache._dateStyleIndexes, Math.Max(4, index * 2));
                }

                if (TryParseUIntAttribute(reader.GetAttribute("numFmtId"), out uint numberFormatId)
                    && IsDateNumberFormat(numberFormatId, numberingFormats)) {
                    cache._dateStyleIndexes[index] = true;
                    cache.HasDateStyles = true;
                }

                index++;
            }

            if (index == 0) {
                cache._dateStyleIndexes = EmptyDateStyleIndexes;
            } else if (index < cache._dateStyleIndexes.Length) {
                Array.Resize(ref cache._dateStyleIndexes, index);
            }
        }

        private static bool IsDateNumberFormat(uint id, Dictionary<uint, string>? numberingFormats) {
            return IsBuiltInDate(id)
                || (numberingFormats != null
                    && numberingFormats.TryGetValue(id, out string? code)
                    && ExcelNumberFormatClassifier.LooksLikeDateFormat(code));
        }

        private static bool IsBuiltInDate(uint id)
            => id is 14 or 15 or 16 or 17 or 18 or 19 or 20 or 21 or 22
                or 27 or 30 or 36 or 45 or 46 or 47;

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

        private static bool TryParseIntAttribute(string? value, out int result) {
            result = 0;
            if (!TryParseUIntAttribute(value, out uint parsed) || parsed > int.MaxValue) {
                return false;
            }

            result = (int)parsed;
            return true;
        }
    }
}

