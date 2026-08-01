using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>Format-neutral workbook named-style metadata.</summary>
    public sealed class ExcelNamedStyleInfo {
        internal ExcelNamedStyleInfo(string name, uint formatId, uint? builtInId, bool hidden) {
            Name = name;
            FormatId = formatId;
            BuiltInId = builtInId;
            Hidden = hidden;
        }
        /// <summary>Style name.</summary>
        public string Name { get; }
        /// <summary>Workbook style-format index.</summary>
        public uint FormatId { get; }
        /// <summary>Built-in Excel style identifier, when applicable.</summary>
        public uint? BuiltInId { get; }
        /// <summary>Whether the style is hidden from the gallery.</summary>
        public bool Hidden { get; }
    }

    public partial class ExcelDocument {
        /// <summary>Lists workbook named styles.</summary>
        public IReadOnlyList<ExcelNamedStyleInfo> GetNamedStyles() {
            Stylesheet? stylesheet = WorkbookPartRoot.WorkbookStylesPart?.Stylesheet;
            if (stylesheet?.CellStyles == null) return System.Array.Empty<ExcelNamedStyleInfo>();
            return new ReadOnlyCollection<ExcelNamedStyleInfo>(stylesheet.CellStyles.Elements<CellStyle>()
                .Where(style => !string.IsNullOrWhiteSpace(style.Name?.Value))
                .Select(style => new ExcelNamedStyleInfo(
                    style.Name!.Value!,
                    style.FormatId?.Value ?? 0U,
                    style.BuiltinId?.Value,
                    style.Hidden?.Value == true))
                .ToArray());
        }

        /// <summary>Removes a custom named-style catalog entry without changing already formatted cells.</summary>
        public bool RemoveNamedStyle(string name) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentNullException(nameof(name));
            bool removed = false;
            Locking.ExecuteWrite(EnsureLock(), () => {
                Stylesheet? stylesheet = WorkbookPartRoot.WorkbookStylesPart?.Stylesheet;
                CellStyle? style = stylesheet?.CellStyles?.Elements<CellStyle>()
                    .FirstOrDefault(item => string.Equals(item.Name?.Value, name.Trim(), StringComparison.OrdinalIgnoreCase));
                if (style == null) return;
                if (style.BuiltinId != null) throw new InvalidOperationException($"Built-in named style '{name}' cannot be removed.");
                style.Remove();
                stylesheet!.CellStyles!.Count = (uint)stylesheet.CellStyles.Count();
                stylesheet.Save();
                MarkPackageDirty();
                removed = true;
            });
            return removed;
        }
    }

    public partial class ExcelSheet {
        /// <summary>Creates or replaces a workbook named style from a cell's current formatting.</summary>
        public ExcelNamedStyleInfo DefineNamedStyle(string name, int sourceRow, int sourceColumn, bool hidden = false) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentNullException(nameof(name));
            if (name.Trim().Length > 255) throw new ArgumentOutOfRangeException(nameof(name));
            if (sourceRow < 1 || sourceRow > A1.MaxRows) throw new ArgumentOutOfRangeException(nameof(sourceRow));
            if (sourceColumn < 1 || sourceColumn > A1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(sourceColumn));
            ExcelNamedStyleInfo? result = null;
            WriteLock(() => {
                var stylesPart = _excelDocument.WorkbookPartRoot.WorkbookStylesPart
                    ?? _excelDocument.WorkbookPartRoot.AddNewPart<DocumentFormat.OpenXml.Packaging.WorkbookStylesPart>();
                Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureNamedStyleContainers(stylesheet);
                string normalizedName = name.Trim();
                CellStyle? style = stylesheet.CellStyles!.Elements<CellStyle>()
                    .FirstOrDefault(item => string.Equals(item.Name?.Value, normalizedName, StringComparison.OrdinalIgnoreCase));
                if (style?.BuiltinId != null) {
                    throw new InvalidOperationException($"Built-in named style '{name}' cannot be replaced.");
                }
                uint cellStyleIndex = TryGetExistingCell(sourceRow, sourceColumn)?.StyleIndex?.Value ?? 0U;
                CellFormat source = stylesheet.CellFormats!.Elements<CellFormat>().ElementAtOrDefault((int)cellStyleIndex)
                    ?? stylesheet.CellFormats.Elements<CellFormat>().First();
                var styleFormat = (CellFormat)source.CloneNode(true);
                styleFormat.FormatId = null;
                uint formatId = ReplaceOrAppendNamedStyleFormat(stylesheet, style, styleFormat);
                style ??= stylesheet.CellStyles.AppendChild(new CellStyle());
                style.Name = normalizedName;
                style.FormatId = formatId;
                style.Hidden = hidden;
                stylesheet.CellStyles.Count = (uint)stylesheet.CellStyles.Count();
                stylesheet.Save();
                result = new ExcelNamedStyleInfo(normalizedName, formatId, null, hidden);
            });
            return result!;
        }

        /// <summary>Applies a workbook named style to a range with cancellation and a hard cell budget.</summary>
        public void ApplyNamedStyle(
            string name,
            string range,
            int maximumCells = 1_000_000,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentNullException(nameof(name));
            if (maximumCells < 1) throw new ArgumentOutOfRangeException(nameof(maximumCells));
            var (r1, c1, r2, c2) = A1.ParseRange(range);
            long cellCount = ((long)r2 - r1 + 1L) * ((long)c2 - c1 + 1L);
            if (cellCount > maximumCells) {
                throw new InvalidOperationException($"Named-style application requires {cellCount} cells, exceeding maximumCells ({maximumCells}).");
            }
            WriteLock(() => {
                cancellationToken.ThrowIfCancellationRequested();
                Stylesheet? stylesheet = _excelDocument.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet;
                CellStyle? style = stylesheet?.CellStyles?.Elements<CellStyle>()
                    .FirstOrDefault(item => string.Equals(item.Name?.Value, name.Trim(), StringComparison.OrdinalIgnoreCase));
                if (style?.FormatId?.Value is not uint formatId
                    || stylesheet!.CellStyleFormats == null
                    || formatId >= stylesheet.CellStyleFormats.Count()) {
                    throw new KeyNotFoundException($"Named style '{name}' was not found.");
                }
                var cellFormat = (CellFormat)stylesheet.CellStyleFormats.Elements<CellFormat>().ElementAt((int)formatId).CloneNode(true);
                cellFormat.FormatId = formatId;
                uint cellFormatId = FindOrAppendFormat(stylesheet.CellFormats!, cellFormat);
                for (int row = r1; row <= r2; row++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    for (int column = c1; column <= c2; column++) GetCell(row, column).StyleIndex = cellFormatId;
                }
                stylesheet.Save();
                WorksheetRoot.Save();
            });
        }

        private static void EnsureNamedStyleContainers(Stylesheet stylesheet) {
            stylesheet.Fonts ??= new Fonts(new Font());
            stylesheet.Fills ??= new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
            stylesheet.Borders ??= new Borders(new Border());
            stylesheet.CellStyleFormats ??= new CellStyleFormats(new CellFormat());
            stylesheet.CellFormats ??= new CellFormats(new CellFormat());
            stylesheet.CellStyles ??= new CellStyles(new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U });
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();
            stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();
            stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();
            stylesheet.CellStyleFormats.Count = (uint)stylesheet.CellStyleFormats.Count();
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
            stylesheet.CellStyles.Count = (uint)stylesheet.CellStyles.Count();
        }

        private static uint FindOrAppendFormat(OpenXmlCompositeElement container, CellFormat candidate) {
            uint index = 0;
            foreach (CellFormat existing in container.Elements<CellFormat>()) {
                if (string.Equals(existing.OuterXml, candidate.OuterXml, StringComparison.Ordinal)) return index;
                index++;
            }
            container.Append(candidate);
            if (container is CellFormats cellFormats) cellFormats.Count = (uint)cellFormats.Count();
            if (container is CellStyleFormats styleFormats) styleFormats.Count = (uint)styleFormats.Count();
            return index;
        }

        private uint ReplaceOrAppendNamedStyleFormat(
            Stylesheet stylesheet,
            CellStyle? style,
            CellFormat replacement) {
            CellStyleFormats styleFormats = stylesheet.CellStyleFormats!;
            if (style?.FormatId?.Value is uint formatId
                && formatId < styleFormats.Count()) {
                bool shared = stylesheet.CellStyles!.Elements<CellStyle>()
                    .Any(candidate => !ReferenceEquals(candidate, style)
                        && candidate.FormatId?.Value == formatId);
                if (shared) {
                    throw new InvalidOperationException(
                        $"Named style '{style.Name?.Value}' shares its base format with another style and cannot be redefined safely.");
                }
                CellFormat previous = styleFormats.Elements<CellFormat>().ElementAt((int)formatId);
                ReplaceAppliedNamedStyleFormats(stylesheet.CellFormats!, previous, replacement, formatId, formatId);
                styleFormats.ReplaceChild(replacement, previous);
                styleFormats.Count = (uint)styleFormats.Count();
                return formatId;
            }

            uint appendedId = (uint)styleFormats.Count();
            styleFormats.Append(replacement);
            styleFormats.Count = appendedId + 1U;
            return appendedId;
        }

        private static void ReplaceAppliedNamedStyleFormats(
            CellFormats cellFormats,
            CellFormat previousStyle,
            CellFormat replacementStyle,
            uint previousFormatId,
            uint replacementFormatId) {
            var previousApplied = (CellFormat)previousStyle.CloneNode(true);
            previousApplied.FormatId = previousFormatId;
            string previousXml = previousApplied.OuterXml;
            foreach (CellFormat applied in cellFormats.Elements<CellFormat>().ToList()) {
                if (!string.Equals(applied.OuterXml, previousXml, StringComparison.Ordinal)) continue;
                var replacement = (CellFormat)replacementStyle.CloneNode(true);
                replacement.FormatId = replacementFormatId;
                cellFormats.ReplaceChild(replacement, applied);
            }
            cellFormats.Count = (uint)cellFormats.Count();
        }

    }
}
