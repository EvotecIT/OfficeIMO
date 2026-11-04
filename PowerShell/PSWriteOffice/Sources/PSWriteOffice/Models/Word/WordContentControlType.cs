namespace PSWriteOffice.Models.Word;

/// <summary>Native content-control family selected by Get-OfficeWordContentControl.</summary>
public enum WordContentControlType {
    /// <summary>Text structured-document tags.</summary>
    Text,
    /// <summary>Checkbox controls.</summary>
    CheckBox,
    /// <summary>Editable combo boxes.</summary>
    ComboBox,
    /// <summary>Date pickers.</summary>
    DatePicker,
    /// <summary>Drop-down lists.</summary>
    DropDownList,
    /// <summary>Picture controls.</summary>
    Picture,
    /// <summary>Repeating sections.</summary>
    RepeatingSection
}

/// <summary>Checkbox state selected by Get-OfficeWordContentControl.</summary>
public enum WordCheckBoxState {
    /// <summary>Checked controls.</summary>
    Checked,
    /// <summary>Unchecked controls.</summary>
    Unchecked
}
