using System.Collections.ObjectModel;
using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO;

/// <summary>Structured accessibility inspection result for a Word document, PowerPoint presentation or Excel workbook.</summary>
public sealed class OfficeAccessibilityReport {
    internal OfficeAccessibilityReport(string format, OfficeAccessibilityPolicyProfile profile, IList<OfficeAccessibilityFinding> findings) {
        Format = format;
        Profile = profile;
        Findings = new ReadOnlyCollection<OfficeAccessibilityFinding>(new List<OfficeAccessibilityFinding>(findings));
    }

    /// <summary>Version of the <see cref="ToJson"/> schema.</summary>
    public int SchemaVersion => 2;
    /// <summary>The inspected format: <c>Word</c>, <c>PowerPoint</c> or <c>Excel</c>.</summary>
    public string Format { get; }
    /// <summary>Policy profile used for inspection.</summary>
    public OfficeAccessibilityPolicyProfile Profile { get; }
    /// <summary>Findings in document order, after the document-level findings.</summary>
    public IReadOnlyList<OfficeAccessibilityFinding> Findings { get; }
    /// <summary>Number of error findings.</summary>
    public int ErrorCount => Count(OfficeAccessibilitySeverity.Error);
    /// <summary>Number of warning findings.</summary>
    public int WarningCount => Count(OfficeAccessibilitySeverity.Warning);
    /// <summary>Whether no policy errors were found.</summary>
    public bool IsSuccessful => ErrorCount == 0;

    /// <summary>Throws <see cref="OfficeAccessibilityException"/> when the report contains errors, or warnings when requested.</summary>
    public OfficeAccessibilityReport EnsureCompliant(bool includeWarnings = false) {
        if (ErrorCount > 0 || includeWarnings && WarningCount > 0) throw new OfficeAccessibilityException(this, includeWarnings);
        return this;
    }

    /// <summary>Serializes the report as dependency-free JSON.</summary>
    public string ToJson(bool indented = true) {
        string newline = indented ? "\n" : string.Empty;
        string i1 = indented ? "  " : string.Empty;
        string i2 = indented ? "    " : string.Empty;
        var json = new StringBuilder();
        json.Append('{').Append(newline);
        json.Append(i1).Append("\"schemaVersion\":").Append(SchemaVersion).Append(',').Append(newline);
        json.Append(i1).Append("\"format\":").Append(Quote(Format)).Append(',').Append(newline);
        json.Append(i1).Append("\"profile\":").Append(Quote(Profile.ToString())).Append(',').Append(newline);
        json.Append(i1).Append("\"errorCount\":").Append(ErrorCount).Append(',').Append(newline);
        json.Append(i1).Append("\"warningCount\":").Append(WarningCount).Append(',').Append(newline);
        json.Append(i1).Append("\"isSuccessful\":").Append(IsSuccessful ? "true" : "false").Append(',').Append(newline);
        json.Append(i1).Append("\"findings\":[").Append(newline);
        for (int index = 0; index < Findings.Count; index++) {
            OfficeAccessibilityFinding finding = Findings[index];
            json.Append(i2).Append('{')
                .Append("\"severity\":").Append(Quote(finding.Severity.ToString())).Append(',')
                .Append("\"code\":").Append(Quote(finding.Code)).Append(',')
                .Append("\"message\":").Append(Quote(finding.Message)).Append(',')
                .Append("\"location\":").Append(Quote(finding.Location)).Append(',')
                .Append("\"locationIndex\":").Append(Number(finding.LocationIndex)).Append(',')
                .Append("\"objectId\":").Append(Number(finding.ObjectId)).Append(',')
                .Append("\"objectName\":").Append(finding.ObjectName == null ? "null" : Quote(finding.ObjectName)).Append(',')
                .Append("\"isHeuristic\":").Append(finding.IsHeuristic ? "true" : "false").Append(',')
                .Append("\"measuredValue\":").Append(Number(finding.MeasuredValue)).Append(',')
                .Append("\"requiredValue\":").Append(Number(finding.RequiredValue))
                .Append('}');
            if (index < Findings.Count - 1) json.Append(',');
            json.Append(newline);
        }
        json.Append(i1).Append(']').Append(newline).Append('}');
        return json.ToString();
    }

    /// <summary>Writes the JSON report to disk, creating the destination directory when needed.</summary>
    public void SaveJson(string path, bool indented = true) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        string fullPath = Path.GetFullPath(path);
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory!);
        OfficeFileCommit.WriteAllBytes(fullPath, new UTF8Encoding(false).GetBytes(ToJson(indented)));
    }

    /// <summary>Serializes the report as a SARIF 2.1.0 log for code-scanning and CI tools.</summary>
    /// <param name="artifactPath">Inspected file. A fully qualified path becomes a <c>file://</c> URI; a relative path stays relative. When null, results carry only logical locations.</param>
    /// <param name="indented">Whether to indent the JSON.</param>
    /// <remarks>Each distinct finding code becomes a rule. Errors map to <c>error</c>, warnings to <c>warning</c> and information to <c>note</c>;
    /// the location, object and measured values are kept in each result's property bag.</remarks>
    public string ToSarif(string? artifactPath = null, bool indented = true) => OfficeAccessibilitySarif.Serialize(this, artifactPath, indented);

    /// <summary>Writes the SARIF 2.1.0 log to disk, creating the destination directory when needed.</summary>
    public void SaveSarif(string path, string? artifactPath = null, bool indented = true) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        string fullPath = Path.GetFullPath(path);
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory!);
        OfficeFileCommit.WriteAllBytes(fullPath, new UTF8Encoding(false).GetBytes(ToSarif(artifactPath, indented)));
    }

    private int Count(OfficeAccessibilitySeverity severity) {
        int count = 0;
        foreach (OfficeAccessibilityFinding finding in Findings) if (finding.Severity == severity) count++;
        return count;
    }

    private static string Number(int? value) => value.HasValue ? value.Value.ToString(CultureInfo.InvariantCulture) : "null";

    private static string Number(uint? value) => value.HasValue ? value.Value.ToString(CultureInfo.InvariantCulture) : "null";

    private static string Number(double? value) => value.HasValue ? value.Value.ToString("0.###", CultureInfo.InvariantCulture) : "null";

    internal static string Quote(string value) {
        var escaped = new StringBuilder(value.Length + 2).Append('"');
        foreach (char character in value) {
            switch (character) {
                case '"': escaped.Append("\\\""); break;
                case '\\': escaped.Append("\\\\"); break;
                case '\n': escaped.Append("\\n"); break;
                case '\r': escaped.Append("\\r"); break;
                case '\t': escaped.Append("\\t"); break;
                default:
                    if (character < ' ') escaped.Append("\\u").Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                    else escaped.Append(character);
                    break;
            }
        }
        return escaped.Append('"').ToString();
    }
}
