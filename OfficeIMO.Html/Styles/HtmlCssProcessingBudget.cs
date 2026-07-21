namespace OfficeIMO.Html;

/// <summary>Tracks operation-wide CSS parsing and selector-matching complexity.</summary>
internal sealed class HtmlCssProcessingBudget {
    private readonly HtmlConversionLimits _limits;
    private long _declarations;
    private long _rules;
    private long _selectorEvaluations;

    internal HtmlCssProcessingBudget(HtmlConversionLimits? limits) {
        _limits = (limits ?? HtmlConversionLimits.CreateTrustedProfile()).Clone();
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

        _declarations += declarationCount;
        if (_limits.MaxCssDeclarations.HasValue && _declarations > _limits.MaxCssDeclarations.Value) {
            throw Limit(
                HtmlConversionDiagnosticCodes.CssDeclarationLimitExceeded,
                nameof(HtmlConversionLimits.MaxCssDeclarations),
                _declarations,
                _limits.MaxCssDeclarations.Value);
        }
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

    private static HtmlDomLimitException Limit(string code, string source, long actual, long limit) =>
        new HtmlDomLimitException(code, "CSS processing exceeded the configured conversion complexity limit.", source, actual, limit);
}
