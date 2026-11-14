using System.Globalization;
using System.Text;
using OfficeIMO.Internal;

namespace OfficeIMO;

/// <summary>Writes <see cref="OfficeAccessibilityReport"/> findings as SARIF 2.1.0 logs for code-scanning and CI tools.</summary>
internal static class OfficeAccessibilitySarif {
    internal const string SchemaUri = "https://json.schemastore.org/sarif-2.1.0.json";

    private static readonly Dictionary<string, string> RuleDescriptions = new(StringComparer.Ordinal) {
        [OfficeDiagnosticCodes.AccessibilityMissingDocumentTitle] = "The document defines a title.",
        [OfficeDiagnosticCodes.AccessibilityMissingLanguage] = "Text defines its language, so screen readers pronounce it correctly.",
        [OfficeDiagnosticCodes.AccessibilityMissingAlternativeText] = "Pictures, charts, shapes and other non-text content have alternative text or are marked decorative.",
        [OfficeDiagnosticCodes.AccessibilityMissingShapeTitle] = "Informative visuals have a concise accessibility title.",
        [OfficeDiagnosticCodes.AccessibilityMissingSlideTitle] = "Every slide has a title.",
        [OfficeDiagnosticCodes.AccessibilityMissingTableHeader] = "Tables specify a header row.",
        [OfficeDiagnosticCodes.AccessibilityUnclearLinkLabel] = "Hyperlink text describes the destination.",
        [OfficeDiagnosticCodes.AccessibilityMissingLinkTooltip] = "Hyperlinks have a screen tip.",
        [OfficeDiagnosticCodes.AccessibilityLowContrast] = "Text has sufficient contrast against its background.",
        [OfficeDiagnosticCodes.AccessibilityChartColorOnlyMeaning] = "Charts do not rely on colour alone to convey meaning.",
        [OfficeDiagnosticCodes.AccessibilityHeadingLevelSkipped] = "Heading levels do not skip a level.",
        [OfficeDiagnosticCodes.AccessibilityMissingHeadings] = "Long documents use headings to structure their content.",
        [OfficeDiagnosticCodes.AccessibilityLayoutTable] = "Tables are not used for page layout.",
        [OfficeDiagnosticCodes.AccessibilityFloatingObject] = "Objects sit inline with text so their reading order is clear.",
        [OfficeDiagnosticCodes.AccessibilityBlankParagraphSpacing] = "Blank paragraphs are not used for spacing.",
        [OfficeDiagnosticCodes.AccessibilityMergedCells] = "Worksheets avoid merged cells.",
        [OfficeDiagnosticCodes.AccessibilityDefaultSheetName] = "Worksheets with content have meaningful names.",
        [OfficeDiagnosticCodes.AccessibilityHiddenSheetContent] = "Worksheets with content are not hidden.",
        [OfficeDiagnosticCodes.AccessibilityBlankCellLayout] = "Blank rows and columns are not used for spacing inside data."
    };

    internal static string Serialize(OfficeAccessibilityReport report, string? artifactPath, bool indented) {
        var rules = new List<object?>();
        var ruleIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (OfficeAccessibilityFinding finding in report.Findings) {
            if (ruleIndexes.ContainsKey(finding.Code)) continue;
            ruleIndexes[finding.Code] = rules.Count;
            rules.Add(Rule(finding.Code, finding.Severity, report.Profile));
        }

        string? uri = string.IsNullOrWhiteSpace(artifactPath) ? null : ToUri(artifactPath!);
        var results = new List<object?>();
        foreach (OfficeAccessibilityFinding finding in report.Findings) {
            var location = new JsonObject();
            if (uri != null) {
                location["physicalLocation"] = new JsonObject {
                    ["artifactLocation"] = new JsonObject { ["uri"] = uri, ["index"] = 0 }
                };
            }
            location["logicalLocations"] = new List<object?> {
                new JsonObject {
                    ["name"] = finding.ObjectName ?? finding.Location,
                    ["fullyQualifiedName"] = finding.ObjectName == null ? finding.Location : finding.Location + "/" + finding.ObjectName,
                    ["kind"] = finding.ObjectName == null ? "resource" : "object"
                }
            };

            var properties = new JsonObject { ["location"] = finding.Location, ["isHeuristic"] = finding.IsHeuristic };
            if (finding.LocationIndex.HasValue) properties["locationIndex"] = finding.LocationIndex.Value;
            if (finding.ObjectId.HasValue) properties["objectId"] = (long)finding.ObjectId.Value;
            if (finding.ObjectName != null) properties["objectName"] = finding.ObjectName;
            if (finding.MeasuredValue.HasValue) properties["measuredValue"] = finding.MeasuredValue.Value;
            if (finding.RequiredValue.HasValue) properties["requiredValue"] = finding.RequiredValue.Value;

            results.Add(new JsonObject {
                ["ruleId"] = finding.Code,
                ["ruleIndex"] = ruleIndexes[finding.Code],
                ["level"] = Level(finding.Severity),
                ["message"] = new JsonObject { ["text"] = finding.Message },
                ["locations"] = new List<object?> { location },
                ["properties"] = properties
            });
        }

        var run = new JsonObject {
            ["tool"] = new JsonObject {
                ["driver"] = new JsonObject {
                    ["name"] = "OfficeIMO Accessibility",
                    ["informationUri"] = "https://officeimo.com",
                    ["semanticVersion"] = DriverVersion(),
                    ["rules"] = rules
                }
            }
        };
        if (uri != null) {
            run["artifacts"] = new List<object?> { new JsonObject { ["location"] = new JsonObject { ["uri"] = uri } } };
        }
        run["results"] = results;
        run["invocations"] = new List<object?> { new JsonObject { ["executionSuccessful"] = true } };
        run["properties"] = new JsonObject {
            ["format"] = report.Format,
            ["profile"] = report.Profile.ToString(),
            ["errorCount"] = report.ErrorCount,
            ["warningCount"] = report.WarningCount,
            ["isSuccessful"] = report.IsSuccessful
        };

        var log = new JsonObject {
            ["$schema"] = SchemaUri,
            ["version"] = "2.1.0",
            ["runs"] = new List<object?> { run }
        };
        var builder = new StringBuilder();
        Write(builder, log, indented, 0);
        return builder.ToString();
    }

    private static JsonObject Rule(string code, OfficeAccessibilitySeverity firstSeverity, OfficeAccessibilityPolicyProfile profile) {
        string description = RuleDescriptions.TryGetValue(code, out string? known) ? known : "Accessibility rule " + code + ".";
        return new JsonObject {
            ["id"] = code,
            ["name"] = RuleName(code),
            ["shortDescription"] = new JsonObject { ["text"] = description },
            ["defaultConfiguration"] = new JsonObject { ["level"] = Level(firstSeverity) },
            ["properties"] = new JsonObject {
                ["tags"] = new List<object?> { "accessibility" },
                ["profile"] = profile.ToString()
            }
        };
    }

    /// <summary>Turns <c>accessibility-missing-table-header</c> into <c>MissingTableHeader</c>.</summary>
    private static string RuleName(string code) {
        const string prefix = "accessibility-";
        string trimmed = code.StartsWith(prefix, StringComparison.Ordinal) ? code.Substring(prefix.Length) : code;
        var name = new StringBuilder(trimmed.Length);
        foreach (string part in trimmed.Split(new[] { '-', '_', '.' }, StringSplitOptions.RemoveEmptyEntries)) {
            name.Append(char.ToUpperInvariant(part[0])).Append(part, 1, part.Length - 1);
        }
        return name.Length == 0 ? "Unknown" : name.ToString();
    }

    private static string Level(OfficeAccessibilitySeverity severity) => severity switch {
        OfficeAccessibilitySeverity.Error => "error",
        OfficeAccessibilitySeverity.Warning => "warning",
        _ => "note"
    };

    private static string DriverVersion() {
        Version? version = typeof(OfficeAccessibilitySarif).Assembly.GetName().Version;
        return version == null ? "0.0.0" : version.Major.ToString(CultureInfo.InvariantCulture) + "." +
            version.Minor.ToString(CultureInfo.InvariantCulture) + "." + Math.Max(0, version.Build).ToString(CultureInfo.InvariantCulture);
    }

    private static string ToUri(string path) {
        if (OfficeRuntimeCompatibility.IsPathFullyQualified(path)) return new Uri(path).AbsoluteUri;
        string[] segments = path.Replace('\\', '/').Split('/');
        for (int index = 0; index < segments.Length; index++) segments[index] = Uri.EscapeDataString(segments[index]);
        return string.Join("/", segments);
    }

    private sealed class JsonObject : List<KeyValuePair<string, object?>> {
        internal object? this[string key] {
            set {
                for (int index = 0; index < Count; index++) {
                    if (this[index].Key == key) {
                        this[index] = new KeyValuePair<string, object?>(key, value);
                        return;
                    }
                }
                Add(new KeyValuePair<string, object?>(key, value));
            }
        }
    }

    private static void Write(StringBuilder json, object? value, bool indented, int depth) {
        switch (value) {
            case null: json.Append("null"); break;
            case string text: json.Append(OfficeAccessibilityReport.Quote(text)); break;
            case bool flag: json.Append(flag ? "true" : "false"); break;
            case int number: json.Append(number.ToString(CultureInfo.InvariantCulture)); break;
            case long number: json.Append(number.ToString(CultureInfo.InvariantCulture)); break;
            case double number: json.Append(number.ToString("0.###", CultureInfo.InvariantCulture)); break;
            case JsonObject members:
                json.Append('{');
                for (int index = 0; index < members.Count; index++) {
                    if (index > 0) json.Append(',');
                    NewLine(json, indented, depth + 1);
                    json.Append(OfficeAccessibilityReport.Quote(members[index].Key)).Append(indented ? ": " : ":");
                    Write(json, members[index].Value, indented, depth + 1);
                }
                if (members.Count > 0) NewLine(json, indented, depth);
                json.Append('}');
                break;
            case List<object?> items:
                json.Append('[');
                for (int index = 0; index < items.Count; index++) {
                    if (index > 0) json.Append(',');
                    NewLine(json, indented, depth + 1);
                    Write(json, items[index], indented, depth + 1);
                }
                if (items.Count > 0) NewLine(json, indented, depth);
                json.Append(']');
                break;
            default: throw new InvalidOperationException("Unsupported SARIF value " + value.GetType().Name + ".");
        }
    }

    private static void NewLine(StringBuilder json, bool indented, int depth) {
        if (!indented) return;
        json.Append('\n').Append(' ', depth * 2);
    }
}
