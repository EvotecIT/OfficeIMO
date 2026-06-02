using System;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Visio {
/// <summary>
    /// Deterministic structural and semantic snapshot of a Visio document.
    /// </summary>
    public sealed class VisioInspectionSnapshot {
        internal VisioInspectionSnapshot(
            string? title,
            string? author,
            string? themeType,
            bool useMastersByDefault,
            bool writeMasterDeltasOnly,
            IReadOnlyList<VisioInspectionMasterSnapshot> masters,
            IReadOnlyList<VisioInspectionPageSnapshot> pages) {
            Title = title;
            Author = author;
            ThemeType = themeType;
            UseMastersByDefault = useMastersByDefault;
            WriteMasterDeltasOnly = writeMasterDeltasOnly;
            Masters = masters;
            Pages = pages;
        }

        /// <summary>Document title.</summary>
        public string? Title { get; }

        /// <summary>Document author.</summary>
        public string? Author { get; }

        /// <summary>Document theme type, when present.</summary>
        public string? ThemeType { get; }

        /// <summary>Whether generated shapes use masters by default.</summary>
        public bool UseMastersByDefault { get; }

        /// <summary>Whether page instances write only master deltas.</summary>
        public bool WriteMasterDeltasOnly { get; }

        /// <summary>Registered masters.</summary>
        public IReadOnlyList<VisioInspectionMasterSnapshot> Masters { get; }

        /// <summary>Document pages.</summary>
        public IReadOnlyList<VisioInspectionPageSnapshot> Pages { get; }

        /// <summary>Gets the total number of shapes, including group children.</summary>
        public int ShapeCount => Pages.Sum(page => page.Shapes.Count);

        /// <summary>Gets the total number of connectors.</summary>
        public int ConnectorCount => Pages.Sum(page => page.Connectors.Count);

        /// <summary>
        /// Compares this snapshot to another snapshot.
        /// </summary>
        public VisioInspectionDiff Diff(VisioInspectionSnapshot other) {
            return VisioInspectionDiff.Compare(this, other);
        }

        /// <summary>
        /// Writes a stable line-oriented representation suitable for golden snapshots and review diffs.
        /// </summary>
        public string ToText() {
            StringBuilder builder = new();
            AppendLine(builder, "document.title", Title);
            AppendLine(builder, "document.author", Author);
            AppendLine(builder, "document.theme", ThemeType);
            AppendLine(builder, "document.useMastersByDefault", UseMastersByDefault);
            AppendLine(builder, "document.writeMasterDeltasOnly", WriteMasterDeltasOnly);
            AppendLine(builder, "document.masterCount", Masters.Count);
            AppendLine(builder, "document.pageCount", Pages.Count);
            AppendLine(builder, "document.shapeCount", ShapeCount);
            AppendLine(builder, "document.connectorCount", ConnectorCount);

            foreach (VisioInspectionMasterSnapshot master in Masters) {
                string prefix = "master[" + EscapeKey(master.Id) + "]";
                AppendLine(builder, prefix + ".nameU", master.NameU);
                AppendLine(builder, prefix + ".shapeNameU", master.ShapeNameU);
                AppendLine(builder, prefix + ".text", master.Text);
                AppendLine(builder, prefix + ".width", master.Width);
                AppendLine(builder, prefix + ".height", master.Height);
                AppendLine(builder, prefix + ".packageBacked", master.IsPackageBacked);
                AppendLine(builder, prefix + ".stencilId", master.StencilId);
                AppendLine(builder, prefix + ".stencilName", master.StencilName);
                AppendLine(builder, prefix + ".stencilCategory", master.StencilCategory);
                AppendLine(builder, prefix + ".stencilCatalog", master.StencilCatalogName);
                AppendLine(builder, prefix + ".stencilSourcePackagePath", master.StencilSourcePackagePath);
                AppendLine(builder, prefix + ".stencilKeywords", string.Join(",", master.StencilKeywords));
                AppendLine(builder, prefix + ".stencilAliases", string.Join(",", master.StencilAliases));
                AppendLine(builder, prefix + ".stencilTags", string.Join(",", master.StencilTags));
                AppendLine(builder, prefix + ".stencilIconNameU", master.StencilIconNameU);
                AppendLine(builder, prefix + ".stencilDefaultWidth", master.StencilDefaultWidth);
                AppendLine(builder, prefix + ".stencilDefaultHeight", master.StencilDefaultHeight);
                AppendLine(builder, prefix + ".stencilDefaultUnit", master.StencilDefaultUnit);
                AppendLine(builder, prefix + ".stencilPreviewImageRelationshipId", master.StencilPreviewImageRelationshipId);
                AppendLine(builder, prefix + ".stencilPreviewImageTarget", master.StencilPreviewImageTarget);
                AppendLine(builder, prefix + ".stencilPreviewImageContentType", master.StencilPreviewImageContentType);
                AppendLine(builder, prefix + ".stencilPreviewImageExtension", master.StencilPreviewImageExtension);
                AppendLine(builder, prefix + ".stencilPreviewImageByteLength", master.StencilPreviewImageByteLength);
            }

            foreach (VisioInspectionPageSnapshot page in Pages) {
                page.AppendText(builder);
            }

            return builder.ToString();
        }

        /// <inheritdoc />
        public override string ToString() {
            return ToText();
        }

        internal static void AppendLine(StringBuilder builder, string key, object? value) {
            builder.Append(key);
            builder.Append('=');
            builder.Append(FormatLineValue(value));
            builder.AppendLine();
        }

        internal static string FormatValue(object? value) {
            if (value == null) {
                return string.Empty;
            }

            if (value is double doubleValue) {
                return doubleValue.ToString("0.######", CultureInfo.InvariantCulture);
            }

            if (value is bool boolValue) {
                return boolValue ? "true" : "false";
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        internal static string EscapeKey(string? value) {
            return EscapeText(value, escapeKeyDelimiters: true);
        }

        internal static string FormatLineValue(object? value) {
            return EscapeValue(FormatValue(value));
        }

        internal static string EscapeValue(string? value) {
            return EscapeText(value, escapeKeyDelimiters: false);
        }

        internal static string UnescapeValue(string? value) {
            if (string.IsNullOrEmpty(value)) {
                return string.Empty;
            }

            StringBuilder builder = new(value!.Length);
            bool escaped = false;
            foreach (char c in value) {
                if (escaped) {
                    switch (c) {
                        case 'r':
                            builder.Append('\r');
                            break;
                        case 'n':
                            builder.Append('\n');
                            break;
                        case '\\':
                        case '[':
                        case ']':
                        case '=':
                            builder.Append(c);
                            break;
                        default:
                            builder.Append('\\');
                            builder.Append(c);
                            break;
                    }

                    escaped = false;
                    continue;
                }

                if (c == '\\') {
                    escaped = true;
                    continue;
                }

                builder.Append(c);
            }

            if (escaped) {
                builder.Append('\\');
            }

            return builder.ToString();
        }

        private static string EscapeText(string? value, bool escapeKeyDelimiters) {
            if (string.IsNullOrEmpty(value)) {
                return string.Empty;
            }

            StringBuilder builder = new(value!.Length);
            foreach (char c in value) {
                switch (c) {
                    case '\\':
                        builder.Append(@"\\");
                        break;
                    case '\r':
                        builder.Append(@"\r");
                        break;
                    case '\n':
                        builder.Append(@"\n");
                        break;
                    case '[' when escapeKeyDelimiters:
                        builder.Append(@"\[");
                        break;
                    case ']' when escapeKeyDelimiters:
                        builder.Append(@"\]");
                        break;
                    case '=' when escapeKeyDelimiters:
                        builder.Append(@"\=");
                        break;
                    default:
                        builder.Append(c);
                        break;
                }
            }

            return builder.ToString();
        }
    }
}
