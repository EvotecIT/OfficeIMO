namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteFormFieldData(StringBuilder builder, RtfFormFieldData? formFieldData, int unicodeSkipCount) {
        if (formFieldData == null) {
            return;
        }

        builder.Append(@"{\*\ffdata");
        foreach (RtfFormFieldDataControl control in formFieldData.Controls) {
            WriteFormFieldControl(builder, control.Name, control.Parameter, control.HasParameter);
        }

        WriteKnownFormFieldControlIfMissing(builder, formFieldData, "fftype", formFieldData.TypeCode);
        WriteKnownFormFieldControlIfMissing(builder, formFieldData, "ffenabled", ToToggleParameter(formFieldData.Enabled));
        WriteKnownFormFieldControlIfMissing(builder, formFieldData, "ffownhelp", ToToggleParameter(formFieldData.OwnHelp));
        WriteKnownFormFieldControlIfMissing(builder, formFieldData, "ffownstat", ToToggleParameter(formFieldData.OwnStatus));
        WriteKnownFormFieldControlIfMissing(builder, formFieldData, "ffprot", ToToggleParameter(formFieldData.Protected));
        WriteKnownFormFieldControlIfMissing(builder, formFieldData, "ffrecalc", ToToggleParameter(formFieldData.RecalculateOnExit));
        WriteKnownFormFieldControlIfMissing(builder, formFieldData, "ffmaxlen", formFieldData.MaxLength);
        WriteKnownFormFieldControlIfMissing(builder, formFieldData, "ffhps", formFieldData.CheckBoxSizeHalfPoints);
        WriteKnownFormFieldControlIfMissing(builder, formFieldData, "ffdefres", formFieldData.DefaultResult);
        WriteKnownFormFieldControlIfMissing(builder, formFieldData, "ffres", formFieldData.Result);
        WriteFormFieldTextDestination(builder, "ffname", formFieldData.Name, unicodeSkipCount);
        WriteFormFieldTextDestination(builder, "ffdeftext", formFieldData.DefaultText, unicodeSkipCount);
        WriteFormFieldTextDestination(builder, "ffformat", formFieldData.Format, unicodeSkipCount);
        WriteFormFieldTextDestination(builder, "ffhelptext", formFieldData.HelpText, unicodeSkipCount);
        WriteFormFieldTextDestination(builder, "ffstattext", formFieldData.StatusText, unicodeSkipCount);
        WriteFormFieldTextDestination(builder, "ffentrymcr", formFieldData.EntryMacro, unicodeSkipCount);
        WriteFormFieldTextDestination(builder, "ffexitmcr", formFieldData.ExitMacro, unicodeSkipCount);
        foreach (string item in formFieldData.DropDownItems) {
            WriteFormFieldTextDestination(builder, "ffl", item, unicodeSkipCount);
        }

        builder.Append('}');
    }

    private static void WriteKnownFormFieldControlIfMissing(StringBuilder builder, RtfFormFieldData formFieldData, string name, int? parameter) {
        if (!parameter.HasValue || formFieldData.Controls.Any(control => control.Name == name)) {
            return;
        }

        WriteFormFieldControl(builder, name, parameter.Value, hasParameter: true);
    }

    private static int? ToToggleParameter(bool? value) => value.HasValue ? value.Value ? 1 : 0 : null;

    private static void WriteFormFieldControl(StringBuilder builder, string name, int? parameter, bool hasParameter) {
        builder.Append('\\');
        builder.Append(name);
        if (hasParameter && parameter.HasValue) {
            builder.Append(parameter.Value.ToString(CultureInfo.InvariantCulture));
        }
    }

    private static void WriteFormFieldTextDestination(StringBuilder builder, string destination, string? value, int unicodeSkipCount) {
        if (value == null) {
            return;
        }

        builder.Append(@"{\");
        builder.Append(destination);
        builder.Append(' ');
        builder.Append(EscapeText(value, unicodeSkipCount));
        builder.Append('}');
    }
}
