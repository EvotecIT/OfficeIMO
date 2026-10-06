using System;

namespace OfficeIMO.Html.Dom;

/// <summary>The kind of form element whose live properties were captured.</summary>
public enum HtmlFormControlStateKind {
    /// <summary>An input's current value, checkedness and indeterminate state.</summary>
    Input,
    /// <summary>A textarea's current value.</summary>
    TextArea,
    /// <summary>A select whose option states describe the live selection, including no selection.</summary>
    Select,
    /// <summary>An option's current selectedness.</summary>
    Option
}

/// <summary>Immutable live form properties, retained separately from authored attributes and text.</summary>
/// <remarks>Changing an attribute does not discard these properties. Set FormState to null to
/// return an owned element to attribute-based form semantics. File contents and text selections
/// are not represented by this state.</remarks>
public sealed class HtmlFormControlState {
    /// <summary>Creates validated state for one form element.</summary>
    public HtmlFormControlState(HtmlFormControlStateKind kind, string? value = null, bool isChecked = false, bool isIndeterminate = false, bool isSelected = false) {
        if (!Enum.IsDefined(typeof(HtmlFormControlStateKind), kind)) throw new ArgumentOutOfRangeException(nameof(kind));
        if (kind == HtmlFormControlStateKind.Input || kind == HtmlFormControlStateKind.TextArea) {
            if (value == null) throw new ArgumentNullException(nameof(value));
        } else if (value != null) throw new ArgumentException("Only input and textarea state has a value.", nameof(value));
        if (kind != HtmlFormControlStateKind.Input && (isChecked || isIndeterminate)) throw new ArgumentException("Checked and indeterminate state requires an input.");
        if (kind != HtmlFormControlStateKind.Option && isSelected) throw new ArgumentException("Selected state requires an option.", nameof(isSelected));
        Kind = kind;
        Value = value;
        IsChecked = isChecked;
        IsIndeterminate = isIndeterminate;
        IsSelected = isSelected;
    }
    /// <summary>The applicable element kind.</summary>
    public HtmlFormControlStateKind Kind { get; }
    /// <summary>Current input or textarea value; null for select and option state.</summary>
    public string? Value { get; }
    /// <summary>Current input checkedness, independent of its checked attribute.</summary>
    public bool IsChecked { get; }
    /// <summary>Current input indeterminate state.</summary>
    public bool IsIndeterminate { get; }
    /// <summary>Current option selectedness, independent of its selected attribute.</summary>
    public bool IsSelected { get; }

    internal string ElementName => Kind == HtmlFormControlStateKind.Input ? "input" : Kind == HtmlFormControlStateKind.TextArea ? "textarea" : Kind == HtmlFormControlStateKind.Select ? "select" : "option";
}
