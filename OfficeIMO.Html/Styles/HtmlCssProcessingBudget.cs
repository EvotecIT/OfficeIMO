namespace OfficeIMO.Html;

/// <summary>Tracks operation-wide CSS parsing and selector-matching complexity.</summary>
internal sealed class HtmlCssProcessingBudget {
    private readonly HtmlConversionLimits _limits;
    private readonly bool _hasConfiguredLimits;
    private long _declarations;
    private long _rules;
    private long _selectorEvaluations;

    internal HtmlCssProcessingBudget(HtmlConversionLimits? limits,
        System.Func<OfficeIMO.Html.Css.IHtmlCssSelectorElement, string, bool>? providerMatcher = null) {
        _hasConfiguredLimits = limits != null;
        _limits = (limits ?? HtmlConversionLimits.CreateTrustedProfile()).Clone();
        SelectorMatchContext = new OfficeIMO.Html.Css.HtmlCssSelectorMatchContext(
            RecordSelectorEvaluation, default, providerMatcher);
    }

    internal OfficeIMO.Html.Css.HtmlCssSelectorMatchContext SelectorMatchContext { get; }

    internal OfficeIMO.Html.Css.HtmlCssSyntaxOptions CreateInlineSyntaxOptions() =>
        new OfficeIMO.Html.Css.HtmlCssSyntaxOptions {
            MaxInputCharacters = null,
            MaxTokens = _hasConfiguredLimits ? _limits.MaxCssTokens : null,
            MaxNestingDepth = _hasConfiguredLimits ? _limits.MaxCssNestingDepth : null,
            MaxSyntaxNodes = _hasConfiguredLimits ? _limits.MaxCssSyntaxNodes : null
        };

    internal OfficeIMO.Html.Css.HtmlCssSyntaxOptions CreateStylesheetSyntaxOptions() => CreateInlineSyntaxOptions();

    internal void RecordDeclarations(int declarationCount) {
        _declarations += declarationCount;
        if (_limits.MaxCssDeclarations.HasValue && _declarations > _limits.MaxCssDeclarations.Value) {
            throw Limit(
                HtmlConversionDiagnosticCodes.CssDeclarationLimitExceeded,
                nameof(HtmlConversionLimits.MaxCssDeclarations),
                _declarations,
                _limits.MaxCssDeclarations.Value);
        }
    }

    internal void RecordRule(int declarationCount) {
        _rules++;
        if (_limits.MaxCssRules.HasValue && _rules > _limits.MaxCssRules.Value) {
            throw Limit(
                HtmlConversionDiagnosticCodes.CssRuleLimitExceeded,
                nameof(HtmlConversionLimits.MaxCssRules),
                _rules,
                _limits.MaxCssRules.Value);
        }

        RecordDeclarations(declarationCount);
    }

    internal void RecordSelectorEvaluation() {
        _selectorEvaluations++;
        if (_limits.MaxSelectorEvaluations.HasValue && _selectorEvaluations > _limits.MaxSelectorEvaluations.Value) {
            throw Limit(
                HtmlConversionDiagnosticCodes.CssSelectorEvaluationLimitExceeded,
                nameof(HtmlConversionLimits.MaxSelectorEvaluations),
                _selectorEvaluations,
                _limits.MaxSelectorEvaluations.Value);
        }
    }

    internal void RecordNestingDepth(int depth) {
        int maximum = _limits.MaxCssNestingDepth ?? 256;
        if (depth > maximum) {
            throw Limit(
                HtmlConversionDiagnosticCodes.CssNestingDepthLimitExceeded,
                nameof(HtmlConversionLimits.MaxCssNestingDepth),
                depth,
                maximum);
        }
    }

    internal void ValidateResolvedSelectorList(int selectors, long characters) {
        if (_limits.MaxCssSelectorsPerRule.HasValue && selectors > _limits.MaxCssSelectorsPerRule.Value) {
            throw Limit(
                HtmlConversionDiagnosticCodes.CssSelectorExpansionLimitExceeded,
                nameof(HtmlConversionLimits.MaxCssSelectorsPerRule),
                selectors,
                _limits.MaxCssSelectorsPerRule.Value);
        }
        if (_limits.MaxCssSelectorCharacters.HasValue && characters > _limits.MaxCssSelectorCharacters.Value) {
            throw Limit(
                HtmlConversionDiagnosticCodes.CssSelectorExpansionLimitExceeded,
                nameof(HtmlConversionLimits.MaxCssSelectorCharacters),
                characters,
                _limits.MaxCssSelectorCharacters.Value);
        }
    }

    internal HtmlDomLimitException TranslateSyntaxLimit(OfficeIMO.Html.Css.HtmlCssSyntaxLimitException exception) {
        string code;
        string source;
        switch (exception.LimitName) {
            case nameof(OfficeIMO.Html.Css.HtmlCssSyntaxOptions.MaxTokens):
                code = HtmlConversionDiagnosticCodes.CssTokenLimitExceeded;
                source = nameof(HtmlConversionLimits.MaxCssTokens);
                break;
            case nameof(OfficeIMO.Html.Css.HtmlCssSyntaxOptions.MaxSyntaxNodes):
                code = HtmlConversionDiagnosticCodes.CssSyntaxNodeLimitExceeded;
                source = nameof(HtmlConversionLimits.MaxCssSyntaxNodes);
                break;
            default:
                code = HtmlConversionDiagnosticCodes.CssNestingDepthLimitExceeded;
                source = nameof(HtmlConversionLimits.MaxCssNestingDepth);
                break;
        }

        return new HtmlDomLimitException(
            code,
            "CSS processing exceeded the configured conversion complexity limit.",
            source,
            exception.Actual,
            exception.Maximum,
            exception);
    }

    private static HtmlDomLimitException Limit(string code, string source, long actual, long limit) =>
        new HtmlDomLimitException(code, "CSS processing exceeded the configured conversion complexity limit.", source, actual, limit);
}
