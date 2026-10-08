using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.ContentSafety;

/// <summary>Bounded heuristic evidence. Completion describes the scan, never a safety verdict.</summary>
public sealed class OfficeContentInstructionAnalysis {
    internal OfficeContentInstructionAnalysis(IReadOnlyList<string> signals, bool complete) {
        Signals = signals;
        IsComplete = complete;
    }
    /// <summary>Deterministic signal identifiers, without source or decoded payloads.</summary>
    public IReadOnlyList<string> Signals { get; }
    /// <summary>False when the source, encoded-candidate or decoded-character budget was reached.</summary>
    public bool IsComplete { get; }
}

public static partial class OfficeContentInstructionDetector {
    private static readonly UTF8Encoding StrictUtf8 = new UTF8Encoding(false, true);

    /// <summary>Inspects ordinary text and one layer of printable UTF-8 Base64, including line-wrapped tokens.
    /// Decoding is inspection-only; nested encodings and arbitrary obfuscation are not evaluated.</summary>
    public static OfficeContentInstructionAnalysis Analyze(string text, int maxCharacters = 1000000,
        int maxDecodedCharacters = 32768, int maxEncodedCandidates = 32) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (maxCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(maxCharacters));
        if (maxDecodedCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(maxDecodedCharacters));
        if (maxEncodedCandidates <= 0) throw new ArgumentOutOfRangeException(nameof(maxEncodedCandidates));
        return Analyze(text, maxCharacters, new OfficeContentInstructionBudget(maxDecodedCharacters, maxEncodedCandidates));
    }

    internal static OfficeContentInstructionAnalysis Analyze(string text, int maxCharacters, OfficeContentInstructionBudget budget) {
        bool complete = text.Length <= maxCharacters;
        string bounded = text.Length <= maxCharacters ? text : text.Substring(0, maxCharacters);
        var signals = new List<string>(DetectPlainText(bounded));
        // Format characters can split a token without appearing in an ordinary human view.
        var source = new StringBuilder(bounded.Length);
        foreach (char value in bounded) {
            if (CharUnicodeInfo.GetUnicodeCategory(value) != UnicodeCategory.Format) source.Append(value);
        }
        string scan = source.ToString();
        for (int index = 0; index < scan.Length;) {
            if (!IsBase64(scan[index])) { index++; continue; }
            int start = index;
            while (index < scan.Length && IsBase64(scan[index])) index++;
            if (index - start < 24) continue;
            int end = index;
            InspectCandidate(start, end);
            // Only join a single line break, never arbitrary words separated by spaces.
            while (end < scan.Length && scan[end - 1] != '=') {
                int next = end;
                while (next < scan.Length && (scan[next] == ' ' || scan[next] == '\t')) next++;
                if (next >= scan.Length || (scan[next] != '\r' && scan[next] != '\n')) break;
                if (scan[next++] == '\r' && next < scan.Length && scan[next] == '\n') next++;
                while (next < scan.Length && (scan[next] == ' ' || scan[next] == '\t')) next++;
                int tokenStart = next;
                while (next < scan.Length && IsBase64(scan[next])) next++;
                if (next - tokenStart < 4) break;
                // A valid token must remain evidence even when the next line is ordinary prose.
                if (next - tokenStart >= 24) InspectCandidate(tokenStart, next);
                end = next;
            }
            if (end != index) InspectCandidate(start, end);
            index = end;
        }
        return new OfficeContentInstructionAnalysis(signals.AsReadOnly(), complete && budget.IsComplete);

        void InspectCandidate(int start, int end) {
            if (!budget.TryCandidate()) return;
            // Bound allocation before decoding, including whitespace in wrapped candidates.
            if ((long)(end - start) > (long)budget.RemainingDecodedCharacters * 4 + 16) {
                budget.MarkIncomplete();
                return;
            }
            try {
                byte[] bytes = Convert.FromBase64String(scan.Substring(start, end - start));
                string decoded = StrictUtf8.GetString(bytes);
                if (!budget.TryDecodedCharacters(decoded.Length)) return;
                if (decoded.Any(value => char.IsControl(value) && value != '\n' && value != '\r' && value != '\t')) return;
                IReadOnlyList<string> inner = DetectPlainText(decoded);
                if (inner.Count == 0) return;
                foreach (string signal in inner) if (!signals.Contains(signal)) signals.Add(signal);
                if (!signals.Contains("encoded-instruction")) signals.Add("encoded-instruction");
            } catch (FormatException) {
                // Ordinary long words and binary encodings are not instruction evidence.
            } catch (DecoderFallbackException) {
                // Only printable UTF-8 text is supported by this bounded inspection.
            }
        }
    }

    private static bool IsBase64(char value) => value >= 'A' && value <= 'Z' ||
        value >= 'a' && value <= 'z' || value >= '0' && value <= '9' ||
        value == '+' || value == '/' || value == '=';
}

// A single operation can inspect several surfaces without multiplying its decoding budget.
internal sealed class OfficeContentInstructionBudget {
    private int _remainingCandidates;
    internal OfficeContentInstructionBudget(int maxDecodedCharacters = 32768, int maxEncodedCandidates = 32) {
        RemainingDecodedCharacters = maxDecodedCharacters;
        _remainingCandidates = maxEncodedCandidates;
    }
    internal int RemainingDecodedCharacters { get; private set; }
    internal bool IsComplete { get; private set; } = true;
    internal void MarkIncomplete() => IsComplete = false;
    internal bool TryCandidate() {
        if (_remainingCandidates == 0) { MarkIncomplete(); return false; }
        _remainingCandidates--;
        return true;
    }
    internal bool TryDecodedCharacters(int count) {
        if (count > RemainingDecodedCharacters) { MarkIncomplete(); return false; }
        RemainingDecodedCharacters -= count;
        return true;
    }
}
