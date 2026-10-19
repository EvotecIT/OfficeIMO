namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Form field types exposed by Add-OfficePdfFormField.</summary>
public enum OfficePdfFormFieldType
{
    /// <summary>Simple text field.</summary>
    Text,

    /// <summary>Simple check box field.</summary>
    CheckBox,

    /// <summary>Simple combo-box choice field.</summary>
    Choice,

    /// <summary>Simple multi-select choice field.</summary>
    MultiSelectChoice,

    /// <summary>Simple radio button group.</summary>
    RadioButton
}
