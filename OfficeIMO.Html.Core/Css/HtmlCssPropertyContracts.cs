using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Html.Css;

/// <summary>Outcome of parsing a value against an OfficeIMO-owned property grammar.</summary>
public enum HtmlCssPropertyParseStatus {
    /// <summary>The property and value are in the currently implemented grammar slice.</summary>
    Parsed,
    /// <summary>The value is syntactically deferred until custom properties are substituted.</summary>
    Deferred,
    /// <summary>The property does not yet have an OfficeIMO-owned grammar.</summary>
    UnknownProperty,
    /// <summary>The property is known, but this value is outside the implemented grammar slice.</summary>
    UnsupportedValue,
    /// <summary>The value contains malformed CSS tokens or unmatched component boundaries.</summary>
    InvalidSyntax
}

/// <summary>Typed shape of a parsed value in the first owned property grammar slice.</summary>
public enum HtmlCssPropertyValueKind {
    /// <summary>A keyword shared by every CSS property.</summary>
    CssWideKeyword,
    /// <summary>A property-specific keyword.</summary>
    Keyword,
    /// <summary>A unitless number.</summary>
    Number,
    /// <summary>A percentage.</summary>
    Percentage,
    /// <summary>A calculated number or percentage whose constant expression was evaluated.</summary>
    Calculation,
    /// <summary>A hexadecimal color.</summary>
    HexColor,
    /// <summary>A CSS named color.</summary>
    NamedColor,
    /// <summary>A system color whose used value comes from the user agent or host.</summary>
    SystemColor,
    /// <summary>The currentColor keyword.</summary>
    CurrentColor,
    /// <summary>An sRGB, HSL, or HWB color function represented by typed components.</summary>
    ColorFunction,
    /// <summary>A function whose value is resolved later, such as var().</summary>
    DeferredFunction
}

/// <summary>Numeric types supported by the first owned CSS math slice.</summary>
public enum HtmlCssNumericType {
    /// <summary>A unitless number.</summary>
    Number,
    /// <summary>A percentage retaining its percentage scale.</summary>
    Percentage
}

/// <summary>A constant numeric result parsed from a literal or CSS math function.</summary>
public sealed class HtmlCssNumericValue {
    internal HtmlCssNumericValue(HtmlCssNumericType type, double value, bool calculated) {
        Type = type; Value = value; IsCalculated = calculated;
    }
    /// <summary>Numeric type of the result.</summary>
    public HtmlCssNumericType Type { get; }
    /// <summary>Finite result; percentages retain their percentage scale.</summary>
    public double Value { get; }
    /// <summary>Whether the value came from calc(), min(), max(), or clamp().</summary>
    public bool IsCalculated { get; }
}

/// <summary>Color functions supported by the first owned typed-color slice.</summary>
public enum HtmlCssColorFunctionKind {
    /// <summary>rgb() or its rgba() alias.</summary>
    Rgb,
    /// <summary>hsl() or its hsla() alias.</summary>
    Hsl,
    /// <summary>hwb().</summary>
    Hwb
}

/// <summary>Typed component category in an owned color function.</summary>
public enum HtmlCssColorComponentKind {
    /// <summary>A unitless number.</summary>
    Number,
    /// <summary>A percentage.</summary>
    Percentage,
    /// <summary>An angle normalized to degrees.</summary>
    Angle,
    /// <summary>The CSS missing-component keyword none.</summary>
    None
}

/// <summary>One finite or missing component of a CSS color function.</summary>
public sealed class HtmlCssColorComponent {
    internal HtmlCssColorComponent(HtmlCssColorComponentKind kind, double? value) { Kind = kind; Value = value; }
    /// <summary>Component category.</summary>
    public HtmlCssColorComponentKind Kind { get; }
    /// <summary>Finite component value, or null for none. Percentages retain their percentage scale; angles use degrees.</summary>
    public double? Value { get; }
}

/// <summary>A provider-independent sRGB, HSL, or HWB functional color.</summary>
public sealed class HtmlCssColorFunctionValue {
    internal HtmlCssColorFunctionValue(
        HtmlCssColorFunctionKind kind,
        IReadOnlyList<HtmlCssColorComponent> components,
        HtmlCssColorComponent alpha,
        bool legacySyntax) {
        Kind = kind; Components = components; Alpha = alpha; UsesLegacyCommaSyntax = legacySyntax;
    }
    /// <summary>Color function model.</summary>
    public HtmlCssColorFunctionKind Kind { get; }
    /// <summary>Three model-specific components in source order.</summary>
    public IReadOnlyList<HtmlCssColorComponent> Components { get; }
    /// <summary>Alpha component; an omitted alpha is represented as the number 1.</summary>
    public HtmlCssColorComponent Alpha { get; }
    /// <summary>Whether rgb()/rgba() or hsl()/hsla() used the legacy comma grammar.</summary>
    public bool UsesLegacyCommaSyntax { get; }
}

/// <summary>Keywords accepted by every CSS property.</summary>
public enum HtmlCssWideKeyword {
    /// <summary>Use the property's initial value.</summary>
    Initial,
    /// <summary>Use the parent's computed value.</summary>
    Inherit,
    /// <summary>Inherit for inherited properties and use the initial value otherwise.</summary>
    Unset,
    /// <summary>Roll back the current cascade origin.</summary>
    Revert,
    /// <summary>Roll back the current cascade layer.</summary>
    RevertLayer
}

/// <summary>Metadata for one property whose value grammar is owned by OfficeIMO.</summary>
public sealed class HtmlCssPropertyDefinition {
    internal HtmlCssPropertyDefinition(string name, bool inherited, string initialValue, string valueSyntax) {
        Name = name;
        IsInherited = inherited;
        InitialValue = initialValue;
        ValueSyntax = valueSyntax;
    }

    /// <summary>Canonical ASCII-lowercase property name.</summary>
    public string Name { get; }
    /// <summary>Whether the computed value normally inherits.</summary>
    public bool IsInherited { get; }
    /// <summary>Canonical initial value for the implemented grammar slice.</summary>
    public string InitialValue { get; }
    /// <summary>Human-readable grammar implemented by this version.</summary>
    public string ValueSyntax { get; }
}

/// <summary>A parsed, provider-independent CSS property value.</summary>
public sealed class HtmlCssPropertyValue {
    internal HtmlCssPropertyValue(
        HtmlCssPropertyValueKind kind,
        string authoredText,
        string canonicalText,
        double? number = null,
        HtmlCssWideKeyword? cssWideKeyword = null,
        HtmlCssNumericValue? numericValue = null,
        HtmlCssColorFunctionValue? colorFunction = null) {
        Kind = kind;
        AuthoredText = authoredText;
        CanonicalText = canonicalText;
        Number = number;
        CssWideKeyword = cssWideKeyword;
        NumericValue = numericValue;
        ColorFunction = colorFunction;
    }

    /// <summary>Typed value category.</summary>
    public HtmlCssPropertyValueKind Kind { get; }
    /// <summary>Trimmed authored value without a declaration's trailing !important annotation.</summary>
    public string AuthoredText { get; }
    /// <summary>Stable lowercase or invariant representation when the grammar defines one.</summary>
    public string CanonicalText { get; }
    /// <summary>Parsed number, or null for nonnumeric values. Percentages retain their authored numeric scale.</summary>
    public double? Number { get; }
    /// <summary>Parsed CSS-wide keyword, or null for a property-specific value.</summary>
    public HtmlCssWideKeyword? CssWideKeyword { get; }
    /// <summary>Typed numeric literal or constant math result, when applicable.</summary>
    public HtmlCssNumericValue? NumericValue { get; }
    /// <summary>Typed functional color, when applicable.</summary>
    public HtmlCssColorFunctionValue? ColorFunction { get; }
}

/// <summary>Result of applying an owned property grammar without making a rendering claim.</summary>
public sealed class HtmlCssPropertyParseResult {
    internal HtmlCssPropertyParseResult(
        string propertyName,
        string authoredValue,
        HtmlCssPropertyParseStatus status,
        HtmlCssPropertyDefinition? definition,
        HtmlCssPropertyValue? value,
        bool isImportant) {
        PropertyName = propertyName;
        AuthoredValue = authoredValue;
        Status = status;
        Definition = definition;
        Value = value;
        IsImportant = isImportant;
    }

    /// <summary>Decoded property name. Custom-property casing is retained.</summary>
    public string PropertyName { get; }
    /// <summary>Trimmed authored value without a declaration's trailing !important annotation.</summary>
    public string AuthoredValue { get; }
    /// <summary>Grammar outcome.</summary>
    public HtmlCssPropertyParseStatus Status { get; }
    /// <summary>Known property metadata, or null for an unknown property.</summary>
    public HtmlCssPropertyDefinition? Definition { get; }
    /// <summary>Typed value when parsed or deferred.</summary>
    public HtmlCssPropertyValue? Value { get; }
    /// <summary>Whether the source declaration ended in !important.</summary>
    public bool IsImportant { get; }
    /// <summary>Whether OfficeIMO can consume this value now or after custom-property substitution.</summary>
    public bool IsAccepted => Status == HtmlCssPropertyParseStatus.Parsed || Status == HtmlCssPropertyParseStatus.Deferred;
}

/// <summary>Catalog for the incrementally owned CSS property grammar.</summary>
public static class HtmlCssPropertyCatalog {
    private static readonly IReadOnlyList<HtmlCssPropertyDefinition> AllDefinitions =
        new ReadOnlyCollection<HtmlCssPropertyDefinition>(new List<HtmlCssPropertyDefinition> {
            new HtmlCssPropertyDefinition("display", false, "inline", "supported single-keyword display values"),
            new HtmlCssPropertyDefinition("visibility", true, "visible", "visible | hidden | collapse"),
            new HtmlCssPropertyDefinition("opacity", false, "1", "<number> | <percentage> | calc() | min() | max() | clamp()"),
            new HtmlCssPropertyDefinition("color", true, "CanvasText", "named color | system color | hex color | currentColor | rgb() | hsl() | hwb()")
        });
    private static readonly Dictionary<string, HtmlCssPropertyDefinition> DefinitionsByName = CreateIndex();

    /// <summary>Properties with an implemented OfficeIMO-owned grammar.</summary>
    public static IReadOnlyList<HtmlCssPropertyDefinition> All => AllDefinitions;

    /// <summary>Finds a definition using CSS property-name casing rules.</summary>
    public static bool TryGet(string propertyName, out HtmlCssPropertyDefinition? definition) {
        if (string.IsNullOrWhiteSpace(propertyName)) {
            definition = null;
            return false;
        }
        return DefinitionsByName.TryGetValue(propertyName.Trim(), out definition);
    }

    private static Dictionary<string, HtmlCssPropertyDefinition> CreateIndex() {
        var definitions = new Dictionary<string, HtmlCssPropertyDefinition>(StringComparer.OrdinalIgnoreCase);
        foreach (HtmlCssPropertyDefinition definition in AllDefinitions) definitions.Add(definition.Name, definition);
        return definitions;
    }
}
