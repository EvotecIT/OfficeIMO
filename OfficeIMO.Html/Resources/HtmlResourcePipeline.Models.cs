using AngleSharp.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    private sealed class CssImportReference {
        internal CssImportReference(int start, int end, string source, string conditionText) {
            Start = start;
            End = end;
            Source = source;
            ConditionText = conditionText;
        }

        internal int Start { get; }
        internal int End { get; }
        internal string Source { get; }
        internal string ConditionText { get; }
    }

    private sealed class SourceRange {
        internal SourceRange(int start, int end) {
            Start = start;
            End = end;
        }

        internal int Start { get; }
        internal int End { get; }
    }

    private sealed class CssStringUrlReference {
        internal CssStringUrlReference(int start, int end, string source) {
            Start = start;
            End = end;
            Source = source;
        }

        internal int Start { get; }
        internal int End { get; }
        internal string Source { get; }
    }

    private sealed class CssCustomPropertyDefinition {
        internal CssCustomPropertyDefinition(string source, string selector, int declarationStart, bool hasUrl, bool isImportant, IReadOnlyList<string> aliases, bool isInline, IElement? inlineOwner, string valueText, string? fallbackAlias) {
            Source = source;
            Selector = selector;
            DeclarationStart = declarationStart;
            HasUrl = hasUrl;
            IsImportant = isImportant;
            Aliases = aliases;
            IsInline = isInline;
            InlineOwner = inlineOwner;
            ValueText = valueText;
            FallbackAlias = fallbackAlias;
        }

        internal string Source { get; }
        internal string Selector { get; }
        internal int DeclarationStart { get; }
        internal bool HasUrl { get; }
        internal bool IsImportant { get; }
        internal IReadOnlyList<string> Aliases { get; }
        internal bool IsInline { get; }
        internal IElement? InlineOwner { get; }
        internal string ValueText { get; }
        internal string? FallbackAlias { get; }
        internal bool IsInheritedKeyword => string.Equals(ValueText, "inherit", StringComparison.OrdinalIgnoreCase)
            || string.Equals(ValueText, "unset", StringComparison.OrdinalIgnoreCase);
        internal bool IsCssWideInvalidatingKeyword => string.Equals(ValueText, "initial", StringComparison.OrdinalIgnoreCase)
            || string.Equals(ValueText, "revert", StringComparison.OrdinalIgnoreCase)
            || string.Equals(ValueText, "revert-layer", StringComparison.OrdinalIgnoreCase);
    }
}
