using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Provenance;

namespace OfficeIMO.ContentSafety;

/// <summary>Shared bounded collector used by format-owned content-safety adapters.</summary>
public sealed class OfficeContentSafetyBuilder {
    private readonly string _format;
    private readonly OfficeContentSafetyOptions _options;
    private readonly List<OfficeContentSafetyFinding> _findings = new List<OfficeContentSafetyFinding>();
    private readonly List<OfficeTextIntegrityFinding> _textIntegrity = new List<OfficeTextIntegrityFinding>();
    private readonly List<string> _diagnostics = new List<string>();
    private int _characters;

    /// <summary>Creates a bounded collector for one format adapter.</summary>
    public OfficeContentSafetyBuilder(string format, OfficeContentSafetyOptions? options = null) {
        if (string.IsNullOrWhiteSpace(format)) throw new ArgumentException("A format name is required.", nameof(format));
        _format = format.Trim();
        _options = options ?? new OfficeContentSafetyOptions();
        _options.Validate();
    }

    /// <summary>Gets the validated inspection options.</summary>
    public OfficeContentSafetyOptions Options => _options;

    /// <summary>Adds exact format evidence and automatically evaluates Unicode and instruction-like content.</summary>
    public OfficeContentSafetyFinding Add(
        OfficeContentConcealmentKind kind,
        OfficeContentSafetyRisk risk,
        string location,
        string evidence,
        string? text,
        OfficeContentCleanupCapability cleanupCapability = OfficeContentCleanupCapability.ReportOnly,
        bool inspectTextIntegrityEvidence = true) {
        if (string.IsNullOrWhiteSpace(location)) throw new ArgumentException("A logical location is required.", nameof(location));
        if (string.IsNullOrWhiteSpace(evidence)) throw new ArgumentException("Exact concealment evidence is required.", nameof(evidence));
        text ??= string.Empty;
        Charge(text.Length);
        EnsureFindingCapacity();

        IReadOnlyList<string> instructionSignals = _options.DetectInstructionLikeText
            ? OfficeContentInstructionDetector.Detect(text)
            : Array.Empty<string>();
        bool instructionLike = instructionSignals.Count > 0;
        if (instructionLike) risk = OfficeContentSafetyRisk.PotentiallyDangerous;
        string contentHash = Hash(text);
        string id = Hash(_format + "\n" + location + "\n" + kind + "\n" + contentHash).Substring(0, 32);
        var finding = new OfficeContentSafetyFinding(
            id,
            _format,
            kind,
            risk,
            location,
            evidence,
            Preview(text, _options.MaxPreviewCharacters),
            text.Length,
            contentHash,
            instructionLike,
            instructionSignals,
            cleanupCapability);
        _findings.Add(finding);

        if (inspectTextIntegrityEvidence) InspectTextIntegrity(location, text, OfficeContentCleanupCapability.ReportOnly);
        return finding;
    }

    /// <summary>Inspects a visible text surface and returns selectable exact Unicode findings.</summary>
    public IReadOnlyList<OfficeContentSafetyFinding> InspectVisibleText(
        string location,
        string? text,
        OfficeContentCleanupCapability cleanupCapability = OfficeContentCleanupCapability.ReportOnly) {
        if (string.IsNullOrWhiteSpace(location)) throw new ArgumentException("A logical location is required.", nameof(location));
        text ??= string.Empty;
        Charge(text.Length);
        return InspectTextIntegrity(location, text, cleanupCapability);
    }

    /// <summary>Inspects an exact native text segment already charged by an owning concealment finding.</summary>
    internal IReadOnlyList<OfficeContentSafetyFinding> InspectChargedTextIntegrity(
        string location,
        string? text,
        OfficeContentCleanupCapability cleanupCapability) {
        if (string.IsNullOrWhiteSpace(location)) throw new ArgumentException("A logical location is required.", nameof(location));
        return InspectTextIntegrity(location, text ?? string.Empty, cleanupCapability);
    }

    private IReadOnlyList<OfficeContentSafetyFinding> InspectTextIntegrity(
        string location,
        string text,
        OfficeContentCleanupCapability cleanupCapability) {
        if (!_options.IncludeTextIntegrityEvidence || text.Length == 0) return Array.Empty<OfficeContentSafetyFinding>();
        int remaining = RemainingFindingCapacity();
        OfficeTextIntegrityReport unicode = OfficeTextIntegrityInspector.Inspect(text, new OfficeTextIntegrityOptions {
            MaxCharacters = Math.Max(1, text.Length),
            MaxFindings = Math.Max(1, remaining),
            IncludeTypographicSpaces = true,
            IncludeVariationSelectors = true
        }, location);
        if (remaining <= 0 && unicode.Findings.Count > 0) {
            throw new InvalidDataException("The asset exceeds the configured combined finding limit.");
        }
        _textIntegrity.AddRange(unicode.Findings);
        var contentFindings = new List<OfficeContentSafetyFinding>(unicode.Findings.Count);
        foreach (OfficeTextIntegrityFinding item in unicode.Findings) {
            EnsureFindingCapacity();
            string exact = text.Substring(item.TextOffset, item.TextLength);
            string contentHash = Hash(exact);
            OfficeContentSafetyRisk risk = item.Risk == OfficeTextIntegrityRisk.PotentiallyDangerous
                ? OfficeContentSafetyRisk.PotentiallyDangerous
                : item.Risk == OfficeTextIntegrityRisk.ContextDependent
                    ? OfficeContentSafetyRisk.ContextDependent
                    : OfficeContentSafetyRisk.Informational;
            string id = Hash(_format + "\n" + location + "\n" + OfficeContentConcealmentKind.NonPrintingUnicode + "\n" +
                item.Kind + "\n" + item.TextOffset.ToString(CultureInfo.InvariantCulture) + "\n" + item.CodePoint.ToString(CultureInfo.InvariantCulture) + "\n" + contentHash).Substring(0, 32);
            var contentFinding = new OfficeContentSafetyFinding(
                id,
                _format,
                OfficeContentConcealmentKind.NonPrintingUnicode,
                risk,
                location,
                "The native text contains " + item.Kind + " (" + item.UnicodeNotation + ") at UTF-16 offset " + item.TextOffset.ToString(CultureInfo.InvariantCulture) + ". This is text-integrity evidence, not proof of AI authorship.",
                "\\u" + item.CodePoint.ToString(item.CodePoint <= 0xFFFF ? "X4" : "X6", CultureInfo.InvariantCulture),
                item.TextLength,
                contentHash,
                false,
                Array.Empty<string>(),
                cleanupCapability,
                item.TextOffset,
                item.TextLength);
            _findings.Add(contentFinding);
            contentFindings.Add(contentFinding);
        }
        return contentFindings.AsReadOnly();
    }

    /// <summary>Adds one bounded adapter diagnostic.</summary>
    public void AddDiagnostic(string diagnostic) {
        if (string.IsNullOrWhiteSpace(diagnostic)) return;
        if (_diagnostics.Count >= 256) throw new InvalidDataException("The asset exceeds the configured content-safety diagnostic limit.");
        _diagnostics.Add(Preview(diagnostic, 512));
    }

    /// <summary>Creates an immutable report in deterministic insertion order.</summary>
    public OfficeContentSafetyReport Build() => new OfficeContentSafetyReport(
        _format,
        _findings.AsReadOnly(),
        _textIntegrity.AsReadOnly(),
        _diagnostics.AsReadOnly());

    /// <summary>Verifies that every selected id exists, is unique, and has the required cleanup capability.</summary>
    public static IReadOnlyList<OfficeContentSafetyFinding> ResolveSelection(
        OfficeContentSafetyReport report,
        OfficeContentCleanupSelection selection) {
        if (report == null) throw new ArgumentNullException(nameof(report));
        if (selection == null) throw new ArgumentNullException(nameof(selection));
        var byId = report.Findings.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var resolved = new List<OfficeContentSafetyFinding>();
        foreach (string id in selection.FindingIds) {
            if (!byId.TryGetValue(id, out OfficeContentSafetyFinding? finding)) {
                throw new ArgumentException("A selected finding does not belong to the current inspection snapshot.", nameof(selection));
            }
            if (finding.CleanupCapability == OfficeContentCleanupCapability.ReportOnly) {
                throw new InvalidOperationException("The selected finding is report-only and cannot be safely removed by this adapter.");
            }
            resolved.Add(finding);
        }
        return resolved.AsReadOnly();
    }

    private void Charge(int characters) {
        if (characters < 0 || _characters > _options.MaxCharacters - characters) {
            throw new InvalidDataException("The asset exceeds the configured decoded-character limit.");
        }
        _characters += characters;
    }

    private int RemainingFindingCapacity() => _options.MaxFindings - _findings.Count;

    private void EnsureFindingCapacity() {
        if (RemainingFindingCapacity() <= 0) {
            throw new InvalidDataException("The asset exceeds the configured combined finding limit.");
        }
    }

    private static string Hash(string value) {
        using SHA256 algorithm = SHA256.Create();
        byte[] hash = algorithm.ComputeHash(Encoding.UTF8.GetBytes(value));
        var builder = new StringBuilder(hash.Length * 2);
        foreach (byte item in hash) builder.Append(item.ToString("x2", CultureInfo.InvariantCulture));
        return builder.ToString();
    }

    private static string Preview(string value, int maximumCharacters) {
        var builder = new StringBuilder(Math.Min(value.Length, maximumCharacters));
        bool previousSpace = false;
        for (int index = 0; index < value.Length && builder.Length < maximumCharacters; index++) {
            char current = value[index];
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(current);
            if (category == UnicodeCategory.Control || category == UnicodeCategory.Format || char.IsSurrogate(current)) {
                string notation = "\\u" + ((int)current).ToString("X4", CultureInfo.InvariantCulture);
                if (builder.Length + notation.Length > maximumCharacters) break;
                builder.Append(notation);
                previousSpace = false;
            } else if (char.IsWhiteSpace(current)) {
                if (!previousSpace) builder.Append(' ');
                previousSpace = true;
            } else {
                builder.Append(current);
                previousSpace = false;
            }
        }
        return builder.ToString().Trim();
    }
}
