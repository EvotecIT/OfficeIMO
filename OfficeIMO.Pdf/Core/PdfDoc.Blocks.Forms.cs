using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDoc {
    /// <summary>Adds a simple AcroForm text field at the current flow position.</summary>
    public PdfDoc TextField(string name, double width = 180, double height = 22, string value = "", PdfAlign align = PdfAlign.Left, double fontSize = 10, double spacingBefore = 0, double spacingAfter = 6, PdfFormFieldStyle? style = null) {
        AddBlock(new TextFieldBlock(name, width, height, value, align, fontSize, spacingBefore, spacingAfter, style));
        return this;
    }

    /// <summary>Adds a simple AcroForm check box at the current flow position.</summary>
    public PdfDoc CheckBox(string name, bool isChecked = false, double size = 14, PdfAlign align = PdfAlign.Left, double spacingBefore = 0, double spacingAfter = 6, string checkedValueName = "Yes", PdfFormFieldStyle? style = null) {
        AddBlock(new CheckBoxBlock(name, isChecked, size, align, spacingBefore, spacingAfter, checkedValueName, style));
        return this;
    }

    /// <summary>Adds a simple AcroForm choice field at the current flow position.</summary>
    public PdfDoc ChoiceField(string name, System.Collections.Generic.IEnumerable<string> options, string? value = null, double width = 180, double height = 22, PdfAlign align = PdfAlign.Left, double fontSize = 10, double spacingBefore = 0, double spacingAfter = 6, bool isComboBox = true, PdfFormFieldStyle? style = null) {
        AddBlock(new ChoiceFieldBlock(name, options, value, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox, style));
        return this;
    }

    /// <summary>Adds a simple AcroForm multi-select choice field at the current flow position.</summary>
    public PdfDoc MultiSelectChoiceField(string name, System.Collections.Generic.IEnumerable<string> options, System.Collections.Generic.IEnumerable<string>? values = null, double width = 180, double height = 72, PdfAlign align = PdfAlign.Left, double fontSize = 10, double spacingBefore = 0, double spacingAfter = 6, PdfFormFieldStyle? style = null) {
        AddBlock(new ChoiceFieldBlock(name, options, values, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox: false, allowsMultipleSelection: true, style));
        return this;
    }

    /// <summary>Adds a simple AcroForm radio button group at the current flow position.</summary>
    public PdfDoc RadioButtonGroup(string name, System.Collections.Generic.IEnumerable<string> options, string? value = null, double size = 14, double gap = 6, PdfAlign align = PdfAlign.Left, double spacingBefore = 0, double spacingAfter = 6, PdfFormFieldStyle? style = null) {
        AddBlock(new RadioButtonGroupBlock(name, options, value, size, gap, align, spacingBefore, spacingAfter, style));
        return this;
    }
}
