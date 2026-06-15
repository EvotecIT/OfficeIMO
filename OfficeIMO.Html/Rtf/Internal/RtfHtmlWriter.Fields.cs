using System.Globalization;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendField(StringBuilder builder, RtfField field, RtfHtmlSaveOptions options, RtfDocument document) {
        builder.Append("<span data-officeimo-rtf-field=\"true\" data-officeimo-rtf-field-instruction=\"");
        builder.Append(EncodeAttribute(field.Instruction));
        builder.Append('"');
        AppendFormFieldAttributes(builder, field.FormFieldData);
        builder.Append('>');
        AppendInlines(builder, field.Result.Inlines, options, document);
        builder.Append("</span>");
    }

    private static void AppendFormFieldAttributes(StringBuilder builder, RtfFormFieldData? data) {
        if (data == null) {
            return;
        }

        builder.Append(" data-officeimo-rtf-form-field=\"true\"");
        string controls = FormatFormFieldControls(data);
        if (controls.Length > 0) {
            AppendAttribute(builder, "data-officeimo-rtf-form-controls", controls);
        }

        AppendAttribute(builder, "data-officeimo-rtf-form-name", data.Name);
        AppendAttribute(builder, "data-officeimo-rtf-form-default-text", data.DefaultText);
        AppendAttribute(builder, "data-officeimo-rtf-form-format", data.Format);
        AppendAttribute(builder, "data-officeimo-rtf-form-help-text", data.HelpText);
        AppendAttribute(builder, "data-officeimo-rtf-form-status-text", data.StatusText);
        AppendAttribute(builder, "data-officeimo-rtf-form-entry-macro", data.EntryMacro);
        AppendAttribute(builder, "data-officeimo-rtf-form-exit-macro", data.ExitMacro);
        string dropDownItems = FormatFormFieldDropDownItems(data);
        if (dropDownItems.Length > 0) {
            AppendAttribute(builder, "data-officeimo-rtf-form-dropdown-items", dropDownItems);
        }
    }

    private static string FormatFormFieldControls(RtfFormFieldData data) {
        var builder = new StringBuilder();
        foreach (RtfFormFieldDataControl control in data.Controls) {
            AppendFormFieldControl(builder, control.Name, control.Parameter, control.HasParameter);
        }

        AppendFormFieldControlIfMissing(builder, data, "fftype", data.TypeCode);
        AppendFormFieldControlIfMissing(builder, data, "ffenabled", ToFormFieldToggle(data.Enabled));
        AppendFormFieldControlIfMissing(builder, data, "ffownhelp", ToFormFieldToggle(data.OwnHelp));
        AppendFormFieldControlIfMissing(builder, data, "ffownstat", ToFormFieldToggle(data.OwnStatus));
        AppendFormFieldControlIfMissing(builder, data, "ffprot", ToFormFieldToggle(data.Protected));
        AppendFormFieldControlIfMissing(builder, data, "ffrecalc", ToFormFieldToggle(data.RecalculateOnExit));
        AppendFormFieldControlIfMissing(builder, data, "ffmaxlen", data.MaxLength);
        AppendFormFieldControlIfMissing(builder, data, "ffhps", data.CheckBoxSizeHalfPoints);
        AppendFormFieldControlIfMissing(builder, data, "ffdefres", data.DefaultResult);
        AppendFormFieldControlIfMissing(builder, data, "ffres", data.Result);
        return builder.ToString();
    }

    private static void AppendFormFieldControlIfMissing(StringBuilder builder, RtfFormFieldData data, string name, int? parameter) {
        if (!parameter.HasValue || data.Controls.Any(control => string.Equals(control.Name, name, StringComparison.Ordinal))) {
            return;
        }

        AppendFormFieldControl(builder, name, parameter.Value, hasParameter: true);
    }

    private static void AppendFormFieldControl(StringBuilder builder, string name, int? parameter, bool hasParameter) {
        if (builder.Length > 0) {
            builder.Append(';');
        }

        builder.Append(name);
        if (hasParameter) {
            builder.Append('=');
            if (parameter.HasValue) {
                builder.Append(parameter.Value.ToString(CultureInfo.InvariantCulture));
            }
        }
    }

    private static string FormatFormFieldDropDownItems(RtfFormFieldData data) {
        var builder = new StringBuilder();
        foreach (string item in data.DropDownItems) {
            if (builder.Length > 0) {
                builder.Append(';');
            }

            builder.Append(Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(item)));
        }

        return builder.ToString();
    }

    private static int? ToFormFieldToggle(bool? value) => value.HasValue ? value.Value ? 1 : 0 : null;

    private static void AppendAttribute(StringBuilder builder, string name, string? value) {
        if (value == null) {
            return;
        }

        builder.Append(' ');
        builder.Append(name);
        builder.Append("=\"");
        builder.Append(EncodeAttribute(value));
        builder.Append('"');
    }
}
