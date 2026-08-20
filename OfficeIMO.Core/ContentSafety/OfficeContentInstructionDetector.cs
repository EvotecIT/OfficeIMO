using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.ContentSafety;

/// <summary>Performs bounded, explainable heuristic detection of instruction-like concealed text.</summary>
public static class OfficeContentInstructionDetector {
    private static readonly SignalRule[] Rules = {
        new SignalRule("instruction-override", new[] { "ignore previous", "ignore prior", "disregard previous", "override instructions", "forget previous" }),
        new SignalRule("prompt-reference", new[] { "system prompt", "developer prompt", "hidden prompt", "reveal prompt", "show prompt" }),
        new SignalRule("model-addressing", new[] { "language model", "llm", "ai assistant", "automated reviewer", "resume scanner", "cv scanner" }),
        new SignalRule("decision-manipulation", new[] { "approve candidate", "accept candidate", "reject other", "rank me", "highest score", "give me a score", "recommend this candidate", "shortlist this candidate" }),
        new SignalRule("concealment-request", new[] { "do not mention", "do not reveal", "hide this", "keep this secret", "without telling the user" }),
        new SignalRule("tool-or-data-exfiltration", new[] { "send the password", "send the token", "upload the secret", "exfiltrate", "forward the code", "retrieve the credential" }),
        new SignalRule("instruction-directive", new[] { "follow these instructions", "you must", "your task is", "when you summarize", "when you evaluate", "when assessing" })
    };

    /// <summary>Returns deterministic signal identifiers without claiming that the text is malicious.</summary>
    public static IReadOnlyList<string> Detect(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (text.Length == 0) return Array.Empty<string>();
        string normalized = Normalize(text);
        var signals = new List<string>();
        foreach (SignalRule rule in Rules) {
            if (rule.Phrases.Any(phrase => normalized.IndexOf(phrase, StringComparison.Ordinal) >= 0)) {
                signals.Add(rule.Id);
            }
        }
        return signals.AsReadOnly();
    }

    private static string Normalize(string text) {
        var builder = new StringBuilder(Math.Min(text.Length, 64 * 1024));
        bool previousSpace = false;
        for (int index = 0; index < text.Length; index++) {
            char value = text[index];
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(value);
            if (char.IsWhiteSpace(value) || char.IsPunctuation(value)) {
                if (!previousSpace) builder.Append(' ');
                previousSpace = true;
                continue;
            }
            if (category == UnicodeCategory.Format || category == UnicodeCategory.Control) continue;
            builder.Append(char.ToLowerInvariant(value));
            previousSpace = false;
        }
        return builder.ToString();
    }

    private sealed class SignalRule {
        internal SignalRule(string id, string[] phrases) { Id = id; Phrases = phrases; }
        internal string Id { get; }
        internal string[] Phrases { get; }
    }
}
