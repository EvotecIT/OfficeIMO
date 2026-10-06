using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Html.Css;

/// <summary>Tokenizes CSS without a parser provider, resource loading or property validation.</summary>
public static class HtmlCssTokenizer {
    /// <summary>
    /// Returns immutable tokens, preserving comments, declaration order and original source spans.
    /// Escapes, nulls, unpaired surrogates and newlines are normalized in decoded values only.
    /// A canceled or budget-exhausted operation publishes no partial token list.
    /// </summary>
    public static IReadOnlyList<HtmlCssToken> Tokenize(string source, HtmlCssTokenizationOptions? options = null, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        HtmlCssTokenizationOptions effective = (options ?? new HtmlCssTokenizationOptions()).Clone();
        effective.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        if (effective.MaxInputCharacters.HasValue && source.Length > effective.MaxInputCharacters.Value)
            throw new HtmlCssTokenizationLimitException(nameof(effective.MaxInputCharacters), source.Length, effective.MaxInputCharacters.Value);
        var tokens = new List<HtmlCssToken>();
        foreach (HtmlCssToken token in Enumerate(source, cancellationToken)) {
            if (token.Kind != HtmlCssTokenKind.EndOfFile && effective.MaxTokens.HasValue && tokens.Count >= effective.MaxTokens.Value)
                throw new HtmlCssTokenizationLimitException(nameof(effective.MaxTokens), (long)tokens.Count + 1, effective.MaxTokens.Value);
            tokens.Add(token);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return tokens.AsReadOnly();
    }

    // Conversion already owns its input/complexity policy. Stream rather than materialize a second token list.
    internal static IEnumerable<HtmlCssToken> Enumerate(string source, CancellationToken cancellationToken = default) {
        var reader = new HtmlCssTokenReader(source, cancellationToken);
        HtmlCssToken token;
        do {
            token = reader.Read();
            yield return token;
        } while (token.Kind != HtmlCssTokenKind.EndOfFile);
    }
}

/// <summary>Source and token budgets for a standalone lexical operation.</summary>
public sealed class HtmlCssTokenizationOptions {
    /// <summary>Maximum UTF-16 input length, or null for caller-bounded input.</summary>
    public int? MaxInputCharacters { get; set; } = 8 * 1024 * 1024;
    /// <summary>Maximum tokens including comments and whitespace, excluding end-of-input; null disables this limit.</summary>
    public int? MaxTokens { get; set; } = 1_000_000;
    /// <summary>Copies the request settings.</summary>
    public HtmlCssTokenizationOptions Clone() => new HtmlCssTokenizationOptions { MaxInputCharacters = MaxInputCharacters, MaxTokens = MaxTokens };
    /// <summary>Rejects nonpositive limits.</summary>
    public void Validate() {
        if (MaxInputCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxInputCharacters));
        if (MaxTokens <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTokens));
    }
}

/// <summary>A CSS lexical operation exceeded a requested resource bound.</summary>
public sealed class HtmlCssTokenizationLimitException : InvalidOperationException {
    /// <summary>Creates a failure describing the exhausted budget.</summary>
    public HtmlCssTokenizationLimitException(string limitName, long actual, long maximum)
        : base($"CSS {limitName} limit exceeded ({actual} > {maximum}).") {
        LimitName = limitName;
        Actual = actual;
        Maximum = maximum;
    }
    /// <summary>The exhausted option.</summary>
    public string LimitName { get; }
    /// <summary>The observed input length or token count.</summary>
    public long Actual { get; }
    /// <summary>The configured ceiling.</summary>
    public long Maximum { get; }
}
