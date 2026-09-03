using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Provenance;

/// <summary>Classifies exact Unicode code points that can be invisible or context-sensitive.</summary>
public enum OfficeTextIntegrityFindingKind {
    /// <summary>A byte-order mark embedded in decoded text.</summary>
    EmbeddedByteOrderMark,
    /// <summary>Unicode ZERO WIDTH SPACE (U+200B).</summary>
    ZeroWidthSpace,
    /// <summary>Unicode ZERO WIDTH NON-JOINER (U+200C).</summary>
    ZeroWidthNonJoiner,
    /// <summary>Unicode ZERO WIDTH JOINER (U+200D).</summary>
    ZeroWidthJoiner,
    /// <summary>Unicode WORD JOINER (U+2060).</summary>
    WordJoiner,
    /// <summary>A Unicode bidirectional mark, isolate, embedding, or override control.</summary>
    BidirectionalControl,
    /// <summary>A Unicode tag character in the supplementary tag block.</summary>
    UnicodeTag,
    /// <summary>A Unicode variation selector.</summary>
    VariationSelector,
    /// <summary>A non-breaking, figure, thin, hair, or narrow no-break space.</summary>
    TypographicSpace,
    /// <summary>A soft hyphen, invisible operator, filler, or similar format character.</summary>
    InvisibleFormatCharacter,
    /// <summary>A C0 or C1 control other than common tab and line terminators.</summary>
    ControlCharacter,
    /// <summary>An unmatched UTF-16 surrogate.</summary>
    UnpairedSurrogate
}

/// <summary>Describes how cautiously a text-integrity finding should be interpreted.</summary>
public enum OfficeTextIntegrityRisk {
    /// <summary>The code point is usually typographic and is reported for exactness.</summary>
    Informational,
    /// <summary>The code point has legitimate language or presentation uses and needs context.</summary>
    ContextDependent,
    /// <summary>The code point can materially alter display or parsing and warrants review.</summary>
    PotentiallyDangerous
}

/// <summary>Bounds Unicode text-integrity inspection.</summary>
public sealed class OfficeTextIntegrityOptions {
    /// <summary>Maximum encoded bytes accepted by file inspection. Defaults to 64 MiB.</summary>
    public long MaxEncodedBytes { get; set; } = 64L * 1024L * 1024L;
    /// <summary>Maximum UTF-16 code units inspected. Defaults to 16 million.</summary>
    public int MaxCharacters { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum findings returned. Defaults to 4,096.</summary>
    public int MaxFindings { get; set; } = 4096;
    /// <summary>Whether a leading decoded byte-order mark is ignored. Defaults to true.</summary>
    public bool IgnoreLeadingByteOrderMark { get; set; } = true;
    /// <summary>Whether typographic spaces are reported. Defaults to true.</summary>
    public bool IncludeTypographicSpaces { get; set; } = true;
    /// <summary>Whether variation selectors are reported. Defaults to true.</summary>
    public bool IncludeVariationSelectors { get; set; } = true;
}

/// <summary>One exact Unicode text-integrity finding.</summary>
public sealed class OfficeTextIntegrityFinding {
    /// <summary>Creates a text-integrity finding.</summary>
    public OfficeTextIntegrityFinding(
        OfficeTextIntegrityFindingKind kind,
        OfficeTextIntegrityRisk risk,
        int textOffset,
        int textLength,
        int codePoint,
        string location = "Text") {
        if (textOffset < 0) throw new ArgumentOutOfRangeException(nameof(textOffset));
        if (textLength <= 0) throw new ArgumentOutOfRangeException(nameof(textLength));
        if (codePoint < 0 || codePoint > 0x10FFFF) throw new ArgumentOutOfRangeException(nameof(codePoint));
        if (string.IsNullOrWhiteSpace(location)) throw new ArgumentException("A finding location is required.", nameof(location));
        Kind = kind;
        Risk = risk;
        TextOffset = textOffset;
        TextLength = textLength;
        CodePoint = codePoint;
        Location = location;
    }

    /// <summary>Gets the classified code-point kind.</summary>
    public OfficeTextIntegrityFindingKind Kind { get; }
    /// <summary>Gets the interpretation risk.</summary>
    public OfficeTextIntegrityRisk Risk { get; }
    /// <summary>Gets the zero-based UTF-16 offset.</summary>
    public int TextOffset { get; }
    /// <summary>Gets the number of UTF-16 code units occupied by the finding.</summary>
    public int TextLength { get; }
    /// <summary>Gets the Unicode scalar value, or the unmatched surrogate value.</summary>
    public int CodePoint { get; }
    /// <summary>Gets the logical text location supplied by the inspector or adapter.</summary>
    public string Location { get; }
    /// <summary>Gets the conventional uppercase Unicode notation.</summary>
    public string UnicodeNotation => $"U+{CodePoint:X4}";
}

/// <summary>Immutable Unicode integrity evidence for one text surface.</summary>
public sealed class OfficeTextIntegrityReport {
    /// <summary>Creates a text-integrity report.</summary>
    public OfficeTextIntegrityReport(IReadOnlyList<OfficeTextIntegrityFinding> findings) {
        Findings = new List<OfficeTextIntegrityFinding>(findings ?? throw new ArgumentNullException(nameof(findings))).AsReadOnly();
    }

    /// <summary>Gets findings in source order.</summary>
    public IReadOnlyList<OfficeTextIntegrityFinding> Findings { get; }
    /// <summary>Gets whether any potentially dangerous control was found.</summary>
    public bool HasPotentiallyDangerousFindings => Findings.Any(item => item.Risk == OfficeTextIntegrityRisk.PotentiallyDangerous);
}

/// <summary>Inspects exact Unicode controls without inferring AI authorship.</summary>
public static class OfficeTextIntegrityInspector {
    private static readonly UTF8Encoding StrictUtf8 = new UTF8Encoding(false, true);

    /// <summary>Inspects a BOM-aware UTF-8, UTF-16, or UTF-32 text file.</summary>
    public static OfficeTextIntegrityReport InspectFile(
        string filePath,
        OfficeTextIntegrityOptions? options = null,
        string? location = null) => InspectFile(filePath, options, location, CancellationToken.None);

    /// <summary>Inspects a BOM-aware UTF-8, UTF-16, or UTF-32 text file with cooperative cancellation.</summary>
    public static OfficeTextIntegrityReport InspectFile(
        string filePath,
        OfficeTextIntegrityOptions? options,
        string? location,
        CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        string fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("The text file was not found.", fullPath);
        options ??= new OfficeTextIntegrityOptions();
        Validate(options);
        cancellationToken.ThrowIfCancellationRequested();
        byte[] data;
        using (var stream = File.OpenRead(fullPath)) {
            data = OfficeProvenanceBinary.ReadBounded(stream, options.MaxEncodedBytes, cancellationToken);
        }
        return Inspect(
            DecodeText(data, options.MaxCharacters, cancellationToken),
            options,
            location ?? fullPath,
            cancellationToken);
    }

    /// <summary>Inspects a string and reports exact, context-sensitive Unicode code points.</summary>
    public static OfficeTextIntegrityReport Inspect(
        string text,
        OfficeTextIntegrityOptions? options = null,
        string location = "Text") => Inspect(text, options, location, CancellationToken.None);

    /// <summary>Inspects a string with cooperative cancellation and reports exact, context-sensitive Unicode code points.</summary>
    public static OfficeTextIntegrityReport Inspect(
        string text,
        OfficeTextIntegrityOptions? options,
        string location,
        CancellationToken cancellationToken) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        options ??= new OfficeTextIntegrityOptions();
        Validate(options);
        if (text.Length > options.MaxCharacters) {
            throw new InvalidDataException("The text exceeds the configured character limit.");
        }
        if (string.IsNullOrWhiteSpace(location)) throw new ArgumentException("A text location is required.", nameof(location));

        cancellationToken.ThrowIfCancellationRequested();
        var findings = new List<OfficeTextIntegrityFinding>();
        for (int offset = 0; offset < text.Length;) {
            if ((offset & 0x3FF) == 0) cancellationToken.ThrowIfCancellationRequested();
            char current = text[offset];
            int codePoint;
            int length;
            if (char.IsHighSurrogate(current)) {
                if (offset + 1 < text.Length && char.IsLowSurrogate(text[offset + 1])) {
                    codePoint = char.ConvertToUtf32(current, text[offset + 1]);
                    length = 2;
                } else {
                    Add(findings, options, OfficeTextIntegrityFindingKind.UnpairedSurrogate,
                        OfficeTextIntegrityRisk.PotentiallyDangerous, offset, 1, current, location);
                    offset++;
                    continue;
                }
            } else if (char.IsLowSurrogate(current)) {
                Add(findings, options, OfficeTextIntegrityFindingKind.UnpairedSurrogate,
                    OfficeTextIntegrityRisk.PotentiallyDangerous, offset, 1, current, location);
                offset++;
                continue;
            } else {
                codePoint = current;
                length = 1;
            }

            if (TryClassify(codePoint, offset, options, out OfficeTextIntegrityFindingKind kind, out OfficeTextIntegrityRisk risk)) {
                Add(findings, options, kind, risk, offset, length, codePoint, location);
            }
            offset += length;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new OfficeTextIntegrityReport(findings.AsReadOnly());
    }

    private static bool TryClassify(
        int codePoint,
        int offset,
        OfficeTextIntegrityOptions options,
        out OfficeTextIntegrityFindingKind kind,
        out OfficeTextIntegrityRisk risk) {
        kind = default;
        risk = default;
        switch (codePoint) {
            case 0xFEFF:
                if (offset == 0 && options.IgnoreLeadingByteOrderMark) return false;
                kind = OfficeTextIntegrityFindingKind.EmbeddedByteOrderMark;
                risk = OfficeTextIntegrityRisk.ContextDependent;
                return true;
            case 0x200B:
                kind = OfficeTextIntegrityFindingKind.ZeroWidthSpace;
                risk = OfficeTextIntegrityRisk.ContextDependent;
                return true;
            case 0x200C:
                kind = OfficeTextIntegrityFindingKind.ZeroWidthNonJoiner;
                risk = OfficeTextIntegrityRisk.ContextDependent;
                return true;
            case 0x200D:
                kind = OfficeTextIntegrityFindingKind.ZeroWidthJoiner;
                risk = OfficeTextIntegrityRisk.ContextDependent;
                return true;
            case 0x2060:
                kind = OfficeTextIntegrityFindingKind.WordJoiner;
                risk = OfficeTextIntegrityRisk.ContextDependent;
                return true;
        }
        if (codePoint is 0x061C or 0x200E or 0x200F ||
            codePoint >= 0x202A && codePoint <= 0x202E ||
            codePoint >= 0x2066 && codePoint <= 0x2069) {
            kind = OfficeTextIntegrityFindingKind.BidirectionalControl;
            risk = codePoint is 0x202D or 0x202E
                ? OfficeTextIntegrityRisk.PotentiallyDangerous
                : OfficeTextIntegrityRisk.ContextDependent;
            return true;
        }
        if (codePoint >= 0xE0000 && codePoint <= 0xE007F) {
            kind = OfficeTextIntegrityFindingKind.UnicodeTag;
            risk = OfficeTextIntegrityRisk.PotentiallyDangerous;
            return true;
        }
        if ((codePoint >= 0xFE00 && codePoint <= 0xFE0F) ||
            (codePoint >= 0xE0100 && codePoint <= 0xE01EF)) {
            if (!options.IncludeVariationSelectors) return false;
            kind = OfficeTextIntegrityFindingKind.VariationSelector;
            risk = OfficeTextIntegrityRisk.ContextDependent;
            return true;
        }
        if (codePoint is 0x00A0 or 0x2007 or 0x2009 or 0x200A or 0x202F) {
            if (!options.IncludeTypographicSpaces) return false;
            kind = OfficeTextIntegrityFindingKind.TypographicSpace;
            risk = OfficeTextIntegrityRisk.Informational;
            return true;
        }
        if (codePoint is 0x00AD or 0x034F or 0x180E or 0x115F or 0x1160 or 0x3164 or 0xFFA0 ||
            codePoint >= 0x2061 && codePoint <= 0x2064) {
            kind = OfficeTextIntegrityFindingKind.InvisibleFormatCharacter;
            risk = OfficeTextIntegrityRisk.ContextDependent;
            return true;
        }
        if (GetUnicodeCategory(codePoint) == UnicodeCategory.Format) {
            kind = OfficeTextIntegrityFindingKind.InvisibleFormatCharacter;
            risk = OfficeTextIntegrityRisk.ContextDependent;
            return true;
        }
        if ((codePoint >= 0 && codePoint < 0x20 && codePoint is not 0x09 and not 0x0A and not 0x0D) ||
            codePoint >= 0x7F && codePoint <= 0x9F) {
            kind = OfficeTextIntegrityFindingKind.ControlCharacter;
            risk = OfficeTextIntegrityRisk.PotentiallyDangerous;
            return true;
        }
        return false;
    }

    private static UnicodeCategory GetUnicodeCategory(int codePoint) {
        if (codePoint <= 0xFFFF) return CharUnicodeInfo.GetUnicodeCategory((char)codePoint);
        string value = char.ConvertFromUtf32(codePoint);
        return CharUnicodeInfo.GetUnicodeCategory(value, 0);
    }

    private static void Add(
        List<OfficeTextIntegrityFinding> findings,
        OfficeTextIntegrityOptions options,
        OfficeTextIntegrityFindingKind kind,
        OfficeTextIntegrityRisk risk,
        int offset,
        int length,
        int codePoint,
        string location) {
        if (findings.Count >= options.MaxFindings) {
            throw new InvalidDataException("The text exceeds the configured finding limit.");
        }
        findings.Add(new OfficeTextIntegrityFinding(kind, risk, offset, length, codePoint, location));
    }

    private static void Validate(OfficeTextIntegrityOptions options) {
        if (options.MaxEncodedBytes <= 0 || options.MaxEncodedBytes > int.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(options), "MaxEncodedBytes must be between one and Int32.MaxValue.");
        }
        if (options.MaxCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(options), "MaxCharacters must be positive.");
        if (options.MaxFindings <= 0) throw new ArgumentOutOfRangeException(nameof(options), "MaxFindings must be positive.");
    }

    private static string DecodeText(byte[] data, int maximumCharacters, CancellationToken cancellationToken) {
        Encoding encoding = StrictUtf8;
        int offset = 0;
        if (StartsWith(data, new byte[] { 0x00, 0x00, 0xFE, 0xFF })) {
            encoding = new UTF32Encoding(true, true, true);
            offset = 4;
        } else if (StartsWith(data, new byte[] { 0xFF, 0xFE, 0x00, 0x00 })) {
            encoding = new UTF32Encoding(false, true, true);
            offset = 4;
        } else if (StartsWith(data, new byte[] { 0xEF, 0xBB, 0xBF })) {
            encoding = StrictUtf8;
            offset = 3;
        } else if (StartsWith(data, new byte[] { 0xFE, 0xFF })) {
            encoding = new UnicodeEncoding(true, true, true);
            offset = 2;
        } else if (StartsWith(data, new byte[] { 0xFF, 0xFE })) {
            encoding = new UnicodeEncoding(false, true, true);
            offset = 2;
        }
        try {
            cancellationToken.ThrowIfCancellationRequested();
            Decoder decoder = encoding.GetDecoder();
            var builder = new StringBuilder(Math.Min(data.Length - offset, maximumCharacters));
            var characters = new char[4096];
            int byteOffset = offset;
            bool completed;
            do {
                cancellationToken.ThrowIfCancellationRequested();
                decoder.Convert(
                    data,
                    byteOffset,
                    data.Length - byteOffset,
                    characters,
                    0,
                    characters.Length,
                    flush: true,
                    out int bytesUsed,
                    out int charactersUsed,
                    out completed);
                if (charactersUsed > maximumCharacters - builder.Length) {
                    throw new InvalidDataException("The text exceeds the configured character limit.");
                }
                builder.Append(characters, 0, charactersUsed);
                byteOffset += bytesUsed;
                if (!completed && bytesUsed == 0 && charactersUsed == 0) {
                    throw new InvalidDataException("The text file could not be decoded incrementally.");
                }
            } while (!completed);
            cancellationToken.ThrowIfCancellationRequested();
            return builder.ToString();
        } catch (DecoderFallbackException exception) {
            throw new InvalidDataException("The text file contains invalid encoded text.", exception);
        }
    }

    private static bool StartsWith(byte[] data, byte[] prefix) {
        if (data.Length < prefix.Length) return false;
        for (int index = 0; index < prefix.Length; index++) if (data[index] != prefix[index]) return false;
        return true;
    }
}

/// <summary>Removes only exact findings explicitly selected by a caller.</summary>
public static class OfficeTextIntegrityCleaner {
    /// <summary>
    /// Removes the selected source ranges after verifying that every finding still identifies the
    /// same code point. The method has no blanket-cleaning default.
    /// </summary>
    public static string RemoveSelected(string text, IEnumerable<OfficeTextIntegrityFinding> findings) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (findings == null) throw new ArgumentNullException(nameof(findings));
        OfficeTextIntegrityFinding[] ordered = findings.OrderBy(item => item.TextOffset).ToArray();
        int previousEnd = 0;
        foreach (OfficeTextIntegrityFinding finding in ordered) {
            if (finding.TextOffset < previousEnd || finding.TextOffset > text.Length - finding.TextLength) {
                throw new ArgumentException("Selected findings overlap or fall outside the supplied text.", nameof(findings));
            }
            int actual = ReadCodePoint(text, finding.TextOffset, finding.TextLength);
            if (actual != finding.CodePoint) {
                throw new ArgumentException("A selected finding no longer matches the supplied text.", nameof(findings));
            }
            previousEnd = finding.TextOffset + finding.TextLength;
        }
        if (ordered.Length == 0) return text;
        var builder = new StringBuilder(text.Length);
        int offset = 0;
        foreach (OfficeTextIntegrityFinding finding in ordered) {
            builder.Append(text, offset, finding.TextOffset - offset);
            offset = finding.TextOffset + finding.TextLength;
        }
        builder.Append(text, offset, text.Length - offset);
        return builder.ToString();
    }

    private static int ReadCodePoint(string text, int offset, int length) {
        if (length == 1) return text[offset];
        if (length == 2 && char.IsHighSurrogate(text[offset]) && char.IsLowSurrogate(text[offset + 1])) {
            return char.ConvertToUtf32(text[offset], text[offset + 1]);
        }
        return -1;
    }
}
