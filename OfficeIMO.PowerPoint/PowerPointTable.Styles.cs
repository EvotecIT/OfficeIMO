using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTable {

        /// <summary>
        ///     Applies a style preset to the table.
        /// </summary>
        public void ApplyStyle(PowerPointTableStylePreset preset) {
            if (!string.IsNullOrWhiteSpace(preset.StyleId)) {
                StyleId = preset.StyleId;
            }
            if (preset.FirstRow.HasValue) {
                FirstRow = preset.FirstRow.Value;
            }
            if (preset.LastRow.HasValue) {
                LastRow = preset.LastRow.Value;
            }
            if (preset.FirstColumn.HasValue) {
                FirstColumn = preset.FirstColumn.Value;
            }
            if (preset.LastColumn.HasValue) {
                LastColumn = preset.LastColumn.Value;
            }
            if (preset.BandedRows.HasValue) {
                BandedRows = preset.BandedRows.Value;
            }
            if (preset.BandedColumns.HasValue) {
                BandedColumns = preset.BandedColumns.Value;
            }
        }

        /// <summary>
        ///     Applies a table style by name, with optional banding/heading toggles.
        /// </summary>
        public void ApplyStyleByName(string styleName, bool ignoreCase = true,
            bool? firstRow = null, bool? lastRow = null, bool? firstColumn = null, bool? lastColumn = null,
            bool? bandedRows = null, bool? bandedColumns = null) {
            if (!TryApplyStyleByName(styleName, ignoreCase, firstRow, lastRow, firstColumn, lastColumn, bandedRows, bandedColumns)) {
                throw new InvalidOperationException($"Table style '{styleName}' was not found in the presentation.");
            }
        }

        /// <summary>
        ///     Tries to apply a table style by name. Returns false if the style is not found.
        /// </summary>
        public bool TryApplyStyleByName(string styleName, bool ignoreCase = true,
            bool? firstRow = null, bool? lastRow = null, bool? firstColumn = null, bool? lastColumn = null,
            bool? bandedRows = null, bool? bandedColumns = null) {
            if (string.IsNullOrWhiteSpace(styleName)) {
                throw new ArgumentException("Style name cannot be null or empty.", nameof(styleName));
            }

            string? styleId = ResolveStyleId(styleName, ignoreCase);
            if (string.IsNullOrWhiteSpace(styleId)) {
                return false;
            }

            StyleId = styleId;
            if (firstRow.HasValue) {
                FirstRow = firstRow.Value;
            }
            if (lastRow.HasValue) {
                LastRow = lastRow.Value;
            }
            if (firstColumn.HasValue) {
                FirstColumn = firstColumn.Value;
            }
            if (lastColumn.HasValue) {
                LastColumn = lastColumn.Value;
            }
            if (bandedRows.HasValue) {
                BandedRows = bandedRows.Value;
            }
            if (bandedColumns.HasValue) {
                BandedColumns = bandedColumns.Value;
            }

            return true;
        }

        private string? ResolveStyleId(string styleName, bool ignoreCase) {
            PresentationPart? presentationPart = _slidePart?
                .GetParentParts()
                .OfType<PresentationPart>()
                .FirstOrDefault();

            if (presentationPart != null && presentationPart.TableStylesPart?.TableStyleList == null) {
                PowerPointUtils.CreateTableStylesPart(presentationPart);
            }

            TableStylesPart? stylesPart = presentationPart?.TableStylesPart;
            StringComparison comparison = ignoreCase
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;

            if (stylesPart?.TableStyleList != null) {
                foreach (A.TableStyle style in stylesPart.TableStyleList.Elements<A.TableStyle>()) {
                    string? styleId = style.StyleId?.Value;
                    if (!string.IsNullOrWhiteSpace(styleId) && string.Equals(styleId, styleName, comparison)) {
                        return styleId;
                    }

                    string? name = style.StyleName?.Value;
                    if (!string.IsNullOrWhiteSpace(name) && string.Equals(name, styleName, comparison)) {
                        return styleId;
                    }
                }
            }

            if (stylesPart != null) {
                using Stream stream = stylesPart.GetStream(FileMode.Open, FileAccess.Read);
                if (stream.Length > 0) {
                    string? resolved = ResolveStyleIdFromXml(styleName, comparison, stream);
                    if (!string.IsNullOrWhiteSpace(resolved)) {
                        return resolved;
                    }
                }
            }

            using Stream? resource = typeof(PowerPointTable).Assembly
                .GetManifestResourceStream("OfficeIMO.PowerPoint.Resources.tableStyles.xml");
            if (resource != null) {
                return ResolveStyleIdFromXml(styleName, comparison, resource);
            }

            return null;
        }

        private static string? ResolveStyleIdFromXml(string styleName, StringComparison comparison, Stream stream) {
            XDocument document = XDocument.Load(stream);
            XElement? root = document.Root;
            if (root == null) {
                return null;
            }

            XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
            foreach (XElement style in root.Elements(drawing + "tblStyle")) {
                string? styleId = style.Attribute("styleId")?.Value;
                if (!string.IsNullOrWhiteSpace(styleId) && string.Equals(styleId, styleName, comparison)) {
                    return styleId;
                }

                string? name = style.Attribute("styleName")?.Value;
                if (!string.IsNullOrWhiteSpace(name) && string.Equals(name, styleName, comparison)) {
                    return styleId;
                }
            }

            return null;
        }
    }
}
