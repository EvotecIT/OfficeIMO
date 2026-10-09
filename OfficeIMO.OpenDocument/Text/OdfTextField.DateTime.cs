namespace OfficeIMO.OpenDocument;

public sealed partial class OdfTextField {
    /// <summary>Native date/time data-style binding. Null removes the binding.</summary>
    public string? DataStyleName {
        get { RequireDateTime(); return (string?)_element.Attribute(OdfNamespaces.Style + "data-style-name"); }
        set {
            RequireDateTime();
            if (value != null) { if (value.Length == 0) throw new ArgumentException("A data-style name must not be empty.", nameof(value)); XmlConvert.VerifyNCName(value); }
            SetDateTimeAttribute(OdfNamespaces.Style + "data-style-name", value);
        }
    }
    /// <summary>Saved native date/dateTime or time/dateTime lexical value. Assignment preserves the original offset notation.</summary>
    /// <remarks>The editable profile accepts Gregorian years 1–9999, clocks through 24:00:00, and at most seven fractional digits. Imported values outside it remain inspectable.</remarks>
    public string? DateTimeValueLexical {
        get { RequireDateTime(); return (string?)_element.Attribute(OdfNamespaces.Text + (Kind == OdfTextFieldKind.Date ? "date-value" : "time-value")); }
        set {
            RequireDateTime();
            if (value != null && !OdfDateTimeFieldValue.TryParse(value, Kind, out _)) throw new ArgumentException("The value is outside the supported native date/time lexical profile.", nameof(value));
            SetDateTimeAttribute(OdfNamespaces.Text + (Kind == OdfTextFieldKind.Date ? "date-value" : "time-value"), value);
        }
    }
    /// <summary>Native ISO duration adjustment. Time fields accept day/clock durations and truncate them to full minutes during projection.</summary>
    public string? DateTimeAdjustmentLexical {
        get { RequireDateTime(); return (string?)_element.Attribute(OdfNamespaces.Text + (Kind == OdfTextFieldKind.Date ? "date-adjust" : "time-adjust")); }
        set {
            RequireDateTime();
            if (value != null && !OdfDateTimeFieldValue.TryReadAdjustment(value, Kind, out _, out _, out _))
                throw new ArgumentException("The adjustment is outside the supported date/time duration profile.", nameof(value));
            SetDateTimeAttribute(OdfNamespaces.Text + (Kind == OdfTextFieldKind.Date ? "date-adjust" : "time-adjust"), value);
        }
    }
    private void RequireDateTime() { if (Kind is not (OdfTextFieldKind.Date or OdfTextFieldKind.Time)) throw new NotSupportedException("This property applies to date/time fields."); }
    private void SetDateTimeAttribute(XName name, string? value) { _element.SetAttributeValue(name, value); _document.MarkPartDirty(_document.GetPartPath(_element)); }
}
