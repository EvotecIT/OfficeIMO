using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.ContentSafety;

/// <summary>Performs bounded, explainable heuristic detection without modifying source text.</summary>
public static partial class OfficeContentInstructionDetector {
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
        return Analyze(text).Signals;
    }

    private static IReadOnlyList<string> DetectPlainText(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (text.Length == 0) return Array.Empty<string>();
        string normalized = Normalize(text);
        var signals = new List<string>();
        foreach (SignalRule rule in Rules) {
            if (rule.Phrases.Any(phrase => HasPhrase(normalized, phrase))) {
                signals.Add(rule.Id);
            }
        }
        if (ContainsAny(normalized, "tool", "tools", "plugin", "plugins", "connector", "connectors") &&
            ContainsAny(normalized, "list every", "list all", "enumerate all", "available to you", "you can access", "write access", "permission scopes", "write send modify")) {
            signals.Add("tool-discovery");
        }
        if (ContainsAny(normalized, "contact list", "contacts", "address book", "email subjects", "recent subjects", "last 20") &&
            ContainsAny(normalized, "collect", "gather", "retrieve", "include", "send", "paste", "forward", "compile", "submit", "append") &&
            ContainsAny(normalized, "http", "https", "query string", "query parameters", "query parameter", "url", "endpoint")) {
            signals.Add("private-data-transfer");
        }
        return signals.AsReadOnly();
    }

    private static bool ContainsAny(string text, params string[] phrases) =>
        phrases.Any(phrase => HasPhrase(text, phrase));

    private static bool HasPhrase(string text, string phrase) {
        for (int start = 0; start <= text.Length - phrase.Length;) {
            int index = text.IndexOf(phrase, start, StringComparison.Ordinal);
            if (index < 0) return false;
            int end = index + phrase.Length;
            if ((index == 0 || !char.IsLetterOrDigit(text[index - 1])) &&
                (end == text.Length || !char.IsLetterOrDigit(text[end]))) return true;
            start = index + 1;
        }
        return false;
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
