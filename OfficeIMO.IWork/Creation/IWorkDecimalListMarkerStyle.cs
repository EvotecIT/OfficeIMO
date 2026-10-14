namespace OfficeIMO.IWork;

/// <summary>Chooses the punctuation around a non-tiered decimal list counter.</summary>
public enum IWorkDecimalListMarkerStyle {
    /// <summary>A counter followed by a period, such as 1.</summary>
    Period = 0,
    /// <summary>A counter inside parentheses, such as (1).</summary>
    Parentheses = 1,
    /// <summary>A counter followed by a closing parenthesis, such as 1).</summary>
    ClosingParenthesis = 2
}
