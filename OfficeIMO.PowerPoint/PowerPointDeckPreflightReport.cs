using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.PowerPoint {
    /// <summary>Severity assigned to a deck preflight finding.</summary>
    public enum PowerPointDeckPreflightSeverity {
        /// <summary>Useful measured context that does not require action.</summary>
        Info,
        /// <summary>The deck can be produced, but the result deserves review.</summary>
        Warning,
        /// <summary>The selected output policy should reject the deck.</summary>
        Error
    }

    /// <summary>One stable, machine-readable deck preflight finding.</summary>
    public sealed class PowerPointDeckPreflightFinding {
        internal PowerPointDeckPreflightFinding(
            PowerPointDeckPreflightSeverity severity,
            string code,
            string message,
            int slideIndex,
            int? shapeIndex = null,
            uint? shapeId = null,
            string? shapeName = null,
            PowerPointLayoutBox? bounds = null,
            double? resolvedFontSizePoints = null) {
            Severity = severity;
            Code = string.IsNullOrWhiteSpace(code) ? "Preflight.Unknown" : code;
            Message = message ?? string.Empty;
            SlideIndex = slideIndex;
            ShapeIndex = shapeIndex;
            ShapeId = shapeId;
            ShapeName = shapeName;
            Bounds = bounds;
            ResolvedFontSizePoints = resolvedFontSizePoints;
        }

        /// <summary>Finding severity.</summary>
        public PowerPointDeckPreflightSeverity Severity { get; }

        /// <summary>Stable diagnostic code suitable for CI policies.</summary>
        public string Code { get; }

        /// <summary>Human-readable explanation.</summary>
        public string Message { get; }

        /// <summary>Zero-based slide index.</summary>
        public int SlideIndex { get; }

        /// <summary>Zero-based explicit-shape index, when the finding refers to a shape.</summary>
        public int? ShapeIndex { get; }

        /// <summary>OOXML shape identifier, when available.</summary>
        public uint? ShapeId { get; }

        /// <summary>Shape name, when available.</summary>
        public string? ShapeName { get; }

        /// <summary>Authored shape bounds, when available.</summary>
        public PowerPointLayoutBox? Bounds { get; }

        /// <summary>Resolved measured font size in points, when the finding concerns text fit.</summary>
        public double? ResolvedFontSizePoints { get; }
    }

    /// <summary>Deterministic deck preflight result shared by code, designer, and markup workflows.</summary>
    public sealed class PowerPointDeckPreflightReport {
        private readonly ReadOnlyCollection<PowerPointDeckPreflightFinding> _findings;

        internal PowerPointDeckPreflightReport(int slideCount, IList<PowerPointDeckPreflightFinding> findings) {
            SlideCount = slideCount;
            _findings = new ReadOnlyCollection<PowerPointDeckPreflightFinding>(
                new List<PowerPointDeckPreflightFinding>(findings ?? throw new ArgumentNullException(nameof(findings))));
        }

        /// <summary>Report schema version.</summary>
        public int SchemaVersion => 1;

        /// <summary>Number of inspected slides.</summary>
        public int SlideCount { get; }

        /// <summary>Findings in deterministic slide and shape order.</summary>
        public IReadOnlyList<PowerPointDeckPreflightFinding> Findings => _findings;

        /// <summary>Number of error findings.</summary>
        public int ErrorCount => Count(PowerPointDeckPreflightSeverity.Error);

        /// <summary>Number of warning findings.</summary>
        public int WarningCount => Count(PowerPointDeckPreflightSeverity.Warning);

        /// <summary>Number of informational findings.</summary>
        public int InfoCount => Count(PowerPointDeckPreflightSeverity.Info);

        /// <summary>Whether the report contains no errors.</summary>
        public bool IsSuccessful => ErrorCount == 0;

        /// <summary>Returns whether any finding meets or exceeds the supplied severity.</summary>
        public bool HasFindingsAtOrAbove(PowerPointDeckPreflightSeverity severity) {
            for (int i = 0; i < _findings.Count; i++) {
                if (_findings[i].Severity >= severity) {
                    return true;
                }
            }

            return false;
        }

        /// <summary>Throws when findings meet or exceed the supplied severity.</summary>
        public void ThrowIfFindings(PowerPointDeckPreflightSeverity severity = PowerPointDeckPreflightSeverity.Error) {
            if (HasFindingsAtOrAbove(severity)) {
                throw new PowerPointDeckPreflightException(this, severity);
            }
        }

        /// <summary>Serializes the report as dependency-free JSON.</summary>
        public string ToJson(bool indented = true) {
            string newline = indented ? Environment.NewLine : string.Empty;
            string i1 = indented ? "  " : string.Empty;
            string i2 = indented ? "    " : string.Empty;
            var json = new StringBuilder();
            json.Append('{').Append(newline);
            AppendNumber(json, i1, "schemaVersion", SchemaVersion, true, newline);
            AppendNumber(json, i1, "slideCount", SlideCount, true, newline);
            AppendNumber(json, i1, "errorCount", ErrorCount, true, newline);
            AppendNumber(json, i1, "warningCount", WarningCount, true, newline);
            AppendNumber(json, i1, "infoCount", InfoCount, true, newline);
            json.Append(i1).Append("\"isSuccessful\": ").Append(IsSuccessful ? "true" : "false").Append(',').Append(newline);
            json.Append(i1).Append("\"findings\": [").Append(newline);
            for (int index = 0; index < _findings.Count; index++) {
                PowerPointDeckPreflightFinding finding = _findings[index];
                json.Append(i2).Append('{');
                AppendString(json, "severity", finding.Severity.ToString(), true);
                AppendString(json, "code", finding.Code, true);
                AppendString(json, "message", finding.Message, true);
                AppendRawNumber(json, "slideIndex", finding.SlideIndex, true);
                AppendNullableNumber(json, "shapeIndex", finding.ShapeIndex, true);
                AppendNullableNumber(json, "shapeId", finding.ShapeId, true);
                AppendNullableString(json, "shapeName", finding.ShapeName, true);
                AppendNullableDouble(json, "resolvedFontSizePoints", finding.ResolvedFontSizePoints, true);
                AppendBounds(json, finding.Bounds);
                json.Append('}');
                if (index < _findings.Count - 1) {
                    json.Append(',');
                }
                json.Append(newline);
            }
            json.Append(i1).Append(']').Append(newline).Append('}');
            return json.ToString();
        }

        /// <summary>Writes the JSON report to disk, creating the destination directory when needed.</summary>
        public void SaveJson(string path, bool indented = true) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("Output path cannot be null or whitespace.", nameof(path));
            }

            string fullPath = Path.GetFullPath(path);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(directory)) {
                Directory.CreateDirectory(directory!);
            }

            File.WriteAllText(fullPath, ToJson(indented), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        }

        private int Count(PowerPointDeckPreflightSeverity severity) {
            int count = 0;
            for (int i = 0; i < _findings.Count; i++) {
                if (_findings[i].Severity == severity) {
                    count++;
                }
            }

            return count;
        }

        private static void AppendNumber(StringBuilder json, string indent, string name, int value, bool comma, string newline) {
            json.Append(indent).Append('"').Append(name).Append("\": ")
                .Append(value.ToString(CultureInfo.InvariantCulture));
            if (comma) json.Append(',');
            json.Append(newline);
        }

        private static void AppendString(StringBuilder json, string name, string value, bool comma) {
            json.Append('"').Append(name).Append("\":\"").Append(Escape(value)).Append('"');
            if (comma) json.Append(',');
        }

        private static void AppendNullableString(StringBuilder json, string name, string? value, bool comma) {
            json.Append('"').Append(name).Append("\":");
            if (value == null) json.Append("null");
            else json.Append('"').Append(Escape(value)).Append('"');
            if (comma) json.Append(',');
        }

        private static void AppendRawNumber(StringBuilder json, string name, long value, bool comma) {
            json.Append('"').Append(name).Append("\":").Append(value.ToString(CultureInfo.InvariantCulture));
            if (comma) json.Append(',');
        }

        private static void AppendNullableNumber(StringBuilder json, string name, long? value, bool comma) {
            json.Append('"').Append(name).Append("\":");
            if (value.HasValue) json.Append(value.Value.ToString(CultureInfo.InvariantCulture));
            else json.Append("null");
            if (comma) json.Append(',');
        }

        private static void AppendNullableNumber(StringBuilder json, string name, uint? value, bool comma) {
            json.Append('"').Append(name).Append("\":");
            if (value.HasValue) json.Append(value.Value.ToString(CultureInfo.InvariantCulture));
            else json.Append("null");
            if (comma) json.Append(',');
        }

        private static void AppendNullableDouble(StringBuilder json, string name, double? value, bool comma) {
            json.Append('"').Append(name).Append("\":");
            if (value.HasValue) json.Append(value.Value.ToString("0.###", CultureInfo.InvariantCulture));
            else json.Append("null");
            if (comma) json.Append(',');
        }

        private static void AppendBounds(StringBuilder json, PowerPointLayoutBox? bounds) {
            json.Append("\"boundsPoints\":");
            if (!bounds.HasValue) {
                json.Append("null");
                return;
            }

            PowerPointLayoutBox value = bounds.Value;
            json.Append('{');
            AppendDouble(json, "left", value.LeftPoints, true);
            AppendDouble(json, "top", value.TopPoints, true);
            AppendDouble(json, "width", value.WidthPoints, true);
            AppendDouble(json, "height", value.HeightPoints, false);
            json.Append('}');
        }

        private static void AppendDouble(StringBuilder json, string name, double value, bool comma) {
            json.Append('"').Append(name).Append("\":").Append(value.ToString("0.###", CultureInfo.InvariantCulture));
            if (comma) json.Append(',');
        }

        private static string Escape(string value) {
            var escaped = new StringBuilder(value.Length + 8);
            for (int i = 0; i < value.Length; i++) {
                char character = value[i];
                switch (character) {
                    case '"': escaped.Append("\\\""); break;
                    case '\\': escaped.Append("\\\\"); break;
                    case '\b': escaped.Append("\\b"); break;
                    case '\f': escaped.Append("\\f"); break;
                    case '\n': escaped.Append("\\n"); break;
                    case '\r': escaped.Append("\\r"); break;
                    case '\t': escaped.Append("\\t"); break;
                    default:
                        if (character < 32) {
                            escaped.Append("\\u").Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                        } else {
                            escaped.Append(character);
                        }
                        break;
                }
            }

            return escaped.ToString();
        }
    }

    /// <summary>Raised when a selected preflight policy rejects a deck.</summary>
    public sealed class PowerPointDeckPreflightException : InvalidOperationException {
        internal PowerPointDeckPreflightException(PowerPointDeckPreflightReport report,
            PowerPointDeckPreflightSeverity severity)
            : base("PowerPoint deck preflight found " + report.Findings.Count +
                   " issue(s); at least one finding met the " + severity + " failure threshold.") {
            Report = report;
            FailureSeverity = severity;
        }

        /// <summary>Report that caused the operation to fail.</summary>
        public PowerPointDeckPreflightReport Report { get; }

        /// <summary>Severity threshold used by the failed operation.</summary>
        public PowerPointDeckPreflightSeverity FailureSeverity { get; }
    }
}
