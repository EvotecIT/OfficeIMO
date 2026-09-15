using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Html.Css;

/// <summary>Provider-neutral origin of a candidate considered by the managed author cascade.</summary>
public enum HtmlCssCascadeSourceKind {
    /// <summary>A declaration from a matched style rule.</summary>
    StyleRule,
    /// <summary>A declaration from an element's style attribute.</summary>
    InlineStyle,
    /// <summary>A style synthesized from an HTML presentational hint.</summary>
    PresentationalHint,
    /// <summary>A computed value inherited from the parent element.</summary>
    Inherited
}

/// <summary>Why a retained cascade candidate did or did not provide the computed value.</summary>
public enum HtmlCssCascadeDecision {
    /// <summary>The candidate supplied the computed value.</summary>
    Selected,
    /// <summary>A candidate with greater cascade precedence supplied the value.</summary>
    Overridden,
    /// <summary>The candidate reset the property to its initial or inherited behavior.</summary>
    Reset,
    /// <summary>The candidate rolled back declarations in its cascade layer.</summary>
    RevertedLayer,
    /// <summary>The candidate rolled back declarations in its cascade origin.</summary>
    RevertedOrigin,
    /// <summary>The value came from the parent because no local candidate supplied it.</summary>
    Inherited,
    /// <summary>Custom-property substitution failed and the property used its inherited or initial fallback.</summary>
    InvalidAtComputedValue
}

/// <summary>Selector specificity retained without exposing a selector-provider type.</summary>
public readonly struct HtmlCssSpecificity : IEquatable<HtmlCssSpecificity> {
    /// <summary>Creates an ID/class/type specificity tuple.</summary>
    public HtmlCssSpecificity(int ids, int classes, int types) { Ids = ids; Classes = classes; Types = types; }
    /// <summary>ID selector count.</summary>
    public int Ids { get; }
    /// <summary>Class, attribute and pseudo-class count.</summary>
    public int Classes { get; }
    /// <summary>Type and pseudo-element count.</summary>
    public int Types { get; }
    /// <inheritdoc />
    public bool Equals(HtmlCssSpecificity other) => Ids == other.Ids && Classes == other.Classes && Types == other.Types;
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is HtmlCssSpecificity other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => unchecked((Ids * 397 ^ Classes) * 397 ^ Types);
    /// <inheritdoc />
    public override string ToString() => $"{Ids},{Classes},{Types}";
}

/// <summary>One supported declaration retained while resolving an owned property slice.</summary>
public sealed class HtmlCssCascadeCandidate {
    /// <summary>Creates an immutable provider-neutral cascade candidate.</summary>
    public HtmlCssCascadeCandidate(
        string declaredValue,
        string? cascadedValue,
        HtmlCssCascadeSourceKind source,
        HtmlCssCascadeDecision decision,
        bool isEffective,
        HtmlCssPropertyParseStatus grammarStatus,
        bool isImportant,
        HtmlCssSpecificity specificity,
        int ruleOrder,
        int declarationOrder,
        string? selector = null,
        string? layerName = null) {
        DeclaredValue = declaredValue ?? throw new ArgumentNullException(nameof(declaredValue));
        CascadedValue = cascadedValue;
        Source = source;
        Decision = decision;
        IsEffective = isEffective;
        GrammarStatus = grammarStatus;
        IsImportant = isImportant;
        Specificity = specificity;
        RuleOrder = ruleOrder;
        DeclarationOrder = declarationOrder;
        Selector = selector;
        LayerName = layerName;
    }

    /// <summary>Declaration value supplied to the cascade, including a CSS-wide keyword or unresolved var().</summary>
    /// <remarks>A retained parser may normalize this text. Use the lossless syntax tree when exact source spelling is required.</remarks>
    public string DeclaredValue { get; }
    /// <summary>Candidate value after CSS-wide keyword handling, or null for a reset or rollback.</summary>
    public string? CascadedValue { get; }
    /// <summary>Where the candidate came from.</summary>
    public HtmlCssCascadeSourceKind Source { get; }
    /// <summary>How the candidate participated in the final result.</summary>
    public HtmlCssCascadeDecision Decision { get; }
    /// <summary>Whether this authored declaration won the local cascade before keyword or substitution fallback was applied.</summary>
    /// <remarks>A reverting declaration remains effective even when another retained candidate supplies the fallback value.</remarks>
    public bool IsEffective { get; }
    /// <summary>Owned property-grammar status for the declaration value.</summary>
    public HtmlCssPropertyParseStatus GrammarStatus { get; }
    /// <summary>Whether the declaration was important.</summary>
    public bool IsImportant { get; }
    /// <summary>Selector specificity. Inline, inherited and presentational candidates also expose their stable internal tuple.</summary>
    public HtmlCssSpecificity Specificity { get; }
    /// <summary>Matched rule order, or -1 for inherited and presentational values.</summary>
    public int RuleOrder { get; }
    /// <summary>Order within the declaration block.</summary>
    public int DeclarationOrder { get; }
    /// <summary>Matched selector text when the source is a style rule.</summary>
    public string? Selector { get; }
    /// <summary>Qualified cascade-layer name, or null for an unlayered declaration.</summary>
    public string? LayerName { get; }
}

/// <summary>Explanation of the managed author cascade for one owned property.</summary>
public sealed class HtmlCssCascadeTrace {
    /// <summary>Creates an immutable trace.</summary>
    public HtmlCssCascadeTrace(
        string propertyName,
        string? computedValue,
        bool isInherited,
        bool isReset,
        IEnumerable<HtmlCssCascadeCandidate> candidates) {
        PropertyName = propertyName ?? throw new ArgumentNullException(nameof(propertyName));
        ComputedValue = computedValue;
        IsInherited = isInherited;
        IsReset = isReset;
        Candidates = new ReadOnlyCollection<HtmlCssCascadeCandidate>(new List<HtmlCssCascadeCandidate>(candidates ?? throw new ArgumentNullException(nameof(candidates))));
    }

    /// <summary>Canonical property name.</summary>
    public string PropertyName { get; }
    /// <summary>Computed string value exposed by the managed style engine, or null when reset without an explicit initial value.</summary>
    public string? ComputedValue { get; }
    /// <summary>Whether the computed value came from the parent.</summary>
    public bool IsInherited { get; }
    /// <summary>Whether a local CSS-wide keyword or invalid-at-computed-value fallback selected initial behavior.</summary>
    public bool IsReset { get; }
    /// <summary>Supported candidates in source order, with the selected candidate identified explicitly.</summary>
    public IReadOnlyList<HtmlCssCascadeCandidate> Candidates { get; }
}
