namespace OfficeIMO.OpenDocument;

/// <summary>The page selected by a native page-number field before its adjustment is applied.</summary>
public enum OdfTextFieldPageSelection {
    /// <summary>The page containing the field.</summary>
    Current,
    /// <summary>The preceding page; a missing preceding page has no displayed number.</summary>
    Previous,
    /// <summary>The following page; a missing following page has no displayed number.</summary>
    Next
}

public sealed partial class OdfTextField {
    /// <summary>
    /// Native page-number/count format: 1, a, A, i, I, or an empty string for no number.
    /// Null removes the override. Other native formats are preserved but require a native formatter.
    /// </summary>
    public string? NumberFormat {
        get { RequireNumberField(); return (string?)_element.Attribute(OdfNamespaces.Style + "num-format"); }
        set { RequireNumberField(); SetNumberAttribute(OdfNamespaces.Style + "num-format", value); }
    }

    /// <summary>True uses repeated letters after Z (AA, BB); false uses AA, AB. Null removes the override.</summary>
    public bool? NumberLetterSync {
        get {
            RequireNumberField();
            string? raw = (string?)_element.Attribute(OdfNamespaces.Style + "num-letter-sync");
            if (raw == null) return null;
            return OdfBoolean.TryParseXml(raw, out bool value) ? value : throw new InvalidDataException("Invalid native number-letter-sync value.");
        }
        set { RequireNumberField(); SetNumberAttribute(OdfNamespaces.Style + "num-letter-sync", value.HasValue ? value.Value ? "true" : "false" : null); }
    }

    /// <summary>The native page selection. This property applies only to page-number fields.</summary>
    public OdfTextFieldPageSelection PageSelection {
        get {
            RequirePageNumber();
            return ((string?)_element.Attribute(OdfNamespaces.Text + "select-page")) switch {
                null or "current" => OdfTextFieldPageSelection.Current,
                "previous" => OdfTextFieldPageSelection.Previous,
                "next" => OdfTextFieldPageSelection.Next,
                _ => throw new InvalidDataException("Invalid native page selection.")
            };
        }
        set {
            RequirePageNumber();
            string? token = value switch {
                OdfTextFieldPageSelection.Current => null,
                OdfTextFieldPageSelection.Previous => "previous",
                OdfTextFieldPageSelection.Next => "next",
                _ => throw new ArgumentOutOfRangeException(nameof(value))
            };
            SetNumberAttribute(OdfNamespaces.Text + "select-page", token);
        }
    }

    /// <summary>Integer added to the selected page number. A result outside the drawing has no displayed number.</summary>
    public int PageAdjustment {
        get {
            RequirePageNumber();
            string? raw = (string?)_element.Attribute(OdfNamespaces.Text + "page-adjust");
            if (raw == null) return 0;
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value :
                throw new InvalidDataException("Invalid native page adjustment.");
        }
        set { RequirePageNumber(); SetNumberAttribute(OdfNamespaces.Text + "page-adjust", value == 0 ? null : value.ToString(CultureInfo.InvariantCulture)); }
    }

    private void RequireNumberField() {
        if (Kind is not (OdfTextFieldKind.PageNumber or OdfTextFieldKind.PageCount))
            throw new NotSupportedException("Numbering properties apply only to page-number and page-count fields.");
    }
    private void RequirePageNumber() {
        if (Kind != OdfTextFieldKind.PageNumber) throw new NotSupportedException("Page selection and adjustment apply only to page-number fields.");
    }
    private void SetNumberAttribute(XName name, string? value) {
        _element.SetAttributeValue(name, value); _document.MarkPartDirty(_document.GetPartPath(_element));
    }
}
