using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>Severity assigned to an accessibility finding.</summary>
    public enum PowerPointAccessibilitySeverity {
        /// <summary>Useful inspection context.</summary>
        Info,
        /// <summary>Potential accessibility risk that deserves review.</summary>
        Warning,
        /// <summary>Policy violation that should fail the selected CI gate.</summary>
        Error
    }

    /// <summary>Inspectable accessibility metadata for one slide shape.</summary>
    public sealed class PowerPointAccessibilityShapeInfo {
        internal PowerPointAccessibilityShapeInfo(int readingOrder, uint? shapeId, string? name, string? title,
            string? description, bool decorative, string? language, PowerPointShapeContentType contentType) {
            ReadingOrder = readingOrder;
            ShapeId = shapeId;
            Name = name;
            Title = title;
            Description = description;
            Decorative = decorative;
            Language = language;
            ContentType = contentType;
        }

        /// <summary>Zero-based shape-tree reading order.</summary>
        public int ReadingOrder { get; }
        /// <summary>OOXML shape identifier.</summary>
        public uint? ShapeId { get; }
        /// <summary>Authored shape name.</summary>
        public string? Name { get; }
        /// <summary>Accessibility title.</summary>
        public string? Title { get; }
        /// <summary>Accessibility description.</summary>
        public string? Description { get; }
        /// <summary>Whether the shape is decorative.</summary>
        public bool Decorative { get; }
        /// <summary>First explicit text language.</summary>
        public string? Language { get; }
        /// <summary>Detected shape content type.</summary>
        public PowerPointShapeContentType ContentType { get; }
    }

    /// <summary>Inspectable accessibility metadata for one slide.</summary>
    public sealed class PowerPointAccessibilitySlideInfo {
        internal PowerPointAccessibilitySlideInfo(int slideIndex, string? title,
            IList<PowerPointAccessibilityShapeInfo> shapes) {
            SlideIndex = slideIndex;
            Title = title;
            Shapes = new ReadOnlyCollection<PowerPointAccessibilityShapeInfo>(
                new List<PowerPointAccessibilityShapeInfo>(shapes));
        }

        /// <summary>Zero-based slide index.</summary>
        public int SlideIndex { get; }
        /// <summary>Detected visible slide title.</summary>
        public string? Title { get; }
        /// <summary>Shapes in inspectable reading order.</summary>
        public IReadOnlyList<PowerPointAccessibilityShapeInfo> Shapes { get; }
    }

    /// <summary>One stable, machine-readable accessibility finding.</summary>
    public sealed class PowerPointAccessibilityFinding {
        internal PowerPointAccessibilityFinding(PowerPointAccessibilitySeverity severity, string code,
            string message, int? slideIndex = null, uint? shapeId = null, string? shapeName = null,
            double? measuredValue = null, double? requiredValue = null) {
            Severity = severity;
            Code = string.IsNullOrWhiteSpace(code) ? "Accessibility.Unknown" : code;
            Message = message ?? string.Empty;
            SlideIndex = slideIndex;
            ShapeId = shapeId;
            ShapeName = shapeName;
            MeasuredValue = measuredValue;
            RequiredValue = requiredValue;
        }

        /// <summary>Finding severity.</summary>
        public PowerPointAccessibilitySeverity Severity { get; }
        /// <summary>Stable diagnostic code suitable for CI policies.</summary>
        public string Code { get; }
        /// <summary>Human-readable explanation.</summary>
        public string Message { get; }
        /// <summary>Zero-based slide index, when applicable.</summary>
        public int? SlideIndex { get; }
        /// <summary>OOXML shape identifier, when applicable.</summary>
        public uint? ShapeId { get; }
        /// <summary>Shape name, when applicable.</summary>
        public string? ShapeName { get; }
        /// <summary>Measured numeric value such as a contrast ratio.</summary>
        public double? MeasuredValue { get; }
        /// <summary>Required numeric policy value.</summary>
        public double? RequiredValue { get; }
    }

    /// <summary>Structured accessibility inspection result for generated or imported presentations.</summary>
    public sealed class PowerPointAccessibilityReport {
        private readonly ReadOnlyCollection<PowerPointAccessibilityFinding> _findings;

        internal PowerPointAccessibilityReport(PowerPointAccessibilityPolicyProfile profile,
            IList<PowerPointAccessibilitySlideInfo> slides, IList<PowerPointAccessibilityFinding> findings) {
            Profile = profile;
            Slides = new ReadOnlyCollection<PowerPointAccessibilitySlideInfo>(
                new List<PowerPointAccessibilitySlideInfo>(slides));
            _findings = new ReadOnlyCollection<PowerPointAccessibilityFinding>(
                new List<PowerPointAccessibilityFinding>(findings));
        }

        /// <summary>Report schema version.</summary>
        public int SchemaVersion => 1;
        /// <summary>Policy profile used for inspection.</summary>
        public PowerPointAccessibilityPolicyProfile Profile { get; }
        /// <summary>Per-slide inspection metadata.</summary>
        public IReadOnlyList<PowerPointAccessibilitySlideInfo> Slides { get; }
        /// <summary>Findings in deterministic document, slide, and shape order.</summary>
        public IReadOnlyList<PowerPointAccessibilityFinding> Findings => _findings;
        /// <summary>Number of error findings.</summary>
        public int ErrorCount => Count(PowerPointAccessibilitySeverity.Error);
        /// <summary>Number of warning findings.</summary>
        public int WarningCount => Count(PowerPointAccessibilitySeverity.Warning);
        /// <summary>Whether no policy errors were found.</summary>
        public bool IsSuccessful => ErrorCount == 0;

        /// <summary>Throws when the report contains errors, or warnings when requested.</summary>
        public PowerPointAccessibilityReport EnsureCompliant(bool includeWarnings = false) {
            if (ErrorCount > 0 || includeWarnings && WarningCount > 0) {
                throw new PowerPointAccessibilityException(this, includeWarnings);
            }
            return this;
        }

        /// <summary>Serializes the report as dependency-free JSON.</summary>
        public string ToJson(bool indented = true) {
            string newline = indented ? Environment.NewLine : string.Empty;
            string i1 = indented ? "  " : string.Empty;
            string i2 = indented ? "    " : string.Empty;
            var json = new StringBuilder();
            json.Append('{').Append(newline);
            json.Append(i1).Append("\"schemaVersion\":").Append(SchemaVersion).Append(',').Append(newline);
            json.Append(i1).Append("\"profile\":\"").Append(Profile).Append("\",").Append(newline);
            json.Append(i1).Append("\"slideCount\":").Append(Slides.Count).Append(',').Append(newline);
            json.Append(i1).Append("\"errorCount\":").Append(ErrorCount).Append(',').Append(newline);
            json.Append(i1).Append("\"warningCount\":").Append(WarningCount).Append(',').Append(newline);
            json.Append(i1).Append("\"isSuccessful\":").Append(IsSuccessful ? "true" : "false").Append(',').Append(newline);
            json.Append(i1).Append("\"findings\": [").Append(newline);
            for (int index = 0; index < _findings.Count; index++) {
                PowerPointAccessibilityFinding finding = _findings[index];
                json.Append(i2).Append('{');
                AppendString(json, "severity", finding.Severity.ToString(), true);
                AppendString(json, "code", finding.Code, true);
                AppendString(json, "message", finding.Message, true);
                AppendNullable(json, "slideIndex", finding.SlideIndex, true);
                AppendNullable(json, "shapeId", finding.ShapeId, true);
                AppendNullableString(json, "shapeName", finding.ShapeName, true);
                AppendNullableDouble(json, "measuredValue", finding.MeasuredValue, true);
                AppendNullableDouble(json, "requiredValue", finding.RequiredValue, false);
                json.Append('}');
                if (index < _findings.Count - 1) json.Append(',');
                json.Append(newline);
            }
            json.Append(i1).Append(']').Append(newline).Append('}');
            return json.ToString();
        }

        /// <summary>Writes the JSON report to disk, creating the destination directory when needed.</summary>
        public void SaveJson(string path, bool indented = true) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Output path cannot be empty.", nameof(path));
            string fullPath = Path.GetFullPath(path);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory!);
            OfficeFileCommit.WriteAllBytes(fullPath, new UTF8Encoding(false).GetBytes(ToJson(indented)));
        }

        private int Count(PowerPointAccessibilitySeverity severity) {
            int count = 0;
            for (int i = 0; i < _findings.Count; i++) if (_findings[i].Severity == severity) count++;
            return count;
        }

        private static void AppendString(StringBuilder json, string name, string value, bool comma) {
            json.Append('"').Append(name).Append("\":\"").Append(Escape(value)).Append('"');
            if (comma) json.Append(',');
        }

        private static void AppendNullableString(StringBuilder json, string name, string? value, bool comma) {
            json.Append('"').Append(name).Append("\":");
            if (value == null) json.Append("null"); else json.Append('"').Append(Escape(value)).Append('"');
            if (comma) json.Append(',');
        }

        private static void AppendNullable(StringBuilder json, string name, long? value, bool comma) {
            json.Append('"').Append(name).Append("\":");
            if (value.HasValue) json.Append(value.Value.ToString(CultureInfo.InvariantCulture)); else json.Append("null");
            if (comma) json.Append(',');
        }

        private static void AppendNullable(StringBuilder json, string name, uint? value, bool comma) =>
            AppendNullable(json, name, value.HasValue ? (long?)value.Value : null, comma);

        private static void AppendNullableDouble(StringBuilder json, string name, double? value, bool comma) {
            json.Append('"').Append(name).Append("\":");
            if (value.HasValue) json.Append(value.Value.ToString("0.###", CultureInfo.InvariantCulture)); else json.Append("null");
            if (comma) json.Append(',');
        }

        private static string Escape(string value) => value.Replace("\\", "\\\\").Replace("\"", "\\\"")
            .Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
    }

    /// <summary>Raised when an accessibility policy rejects a presentation.</summary>
    public sealed class PowerPointAccessibilityException : InvalidOperationException {
        internal PowerPointAccessibilityException(PowerPointAccessibilityReport report, bool includeWarnings)
            : base("PowerPoint accessibility inspection found " + report.ErrorCount + " error(s) and " +
                   report.WarningCount + " warning(s).") {
            Report = report;
            IncludedWarnings = includeWarnings;
        }

        /// <summary>Report that caused the policy failure.</summary>
        public PowerPointAccessibilityReport Report { get; }
        /// <summary>Whether warnings participated in the failure threshold.</summary>
        public bool IncludedWarnings { get; }
    }
}
