namespace OfficeIMO.Rtf;

/// <summary>
/// Word form-field kind represented by the RTF <c>\fftype</c> control.
/// </summary>
public enum RtfFormFieldKind {
    /// <summary>Text input form field.</summary>
    Text = 0,

    /// <summary>Check box form field.</summary>
    CheckBox = 1,

    /// <summary>Drop-down form field.</summary>
    DropDown = 2
}
