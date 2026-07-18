namespace OfficeIMO.Pdf;

internal static partial class PdfAcroFormEditor {
    private static void ValidateReadback(PdfDocumentInfo saved, IReadOnlyList<string> calculationOrder, IReadOnlyList<PdfAcroFormEditSession.EditCommand> commands) {
        var byName = saved.FormFields.Where(static field => !string.IsNullOrEmpty(field.Name)).ToDictionary(static field => field.Name!, StringComparer.Ordinal);
        for (int i = 0; i < commands.Count; i++) {
            PdfAcroFormEditSession.EditCommand command = commands[i];
            switch (command.Kind) {
                case PdfAcroFormEditSession.EditKind.Create:
                    if (!IsRemovedLater(commands, i, command.Options!.Name) && (!byName.TryGetValue(command.Options.Name, out PdfFormField? created) || created.Kind != ToFieldKind(command.Options.Kind))) throw new InvalidOperationException("AcroForm create readback validation failed for " + command.Options.Name + ".");
                    break;
                case PdfAcroFormEditSession.EditKind.Rename:
                    if (!IsRemovedLater(commands, i, command.Value!) && !byName.ContainsKey(command.Value!)) throw new InvalidOperationException("AcroForm rename readback validation failed for " + command.Value + ".");
                    break;
                case PdfAcroFormEditSession.EditKind.Remove:
                    if (byName.ContainsKey(command.Name!)) throw new InvalidOperationException("AcroForm remove readback validation failed for " + command.Name + ".");
                    break;
                case PdfAcroFormEditSession.EditKind.DefaultValue:
                    if (byName.TryGetValue(command.Name!, out PdfFormField? defaultField) && !string.Equals(defaultField.DefaultValue, command.Value, StringComparison.Ordinal)) throw new InvalidOperationException("AcroForm default-value readback validation failed for " + command.Name + ".");
                    break;
                case PdfAcroFormEditSession.EditKind.Flags:
                    if (byName.TryGetValue(command.Name!, out PdfFormField? flagsField) && flagsField.Flags != command.Number) throw new InvalidOperationException("AcroForm flags readback validation failed for " + command.Name + ".");
                    break;
                case PdfAcroFormEditSession.EditKind.TabOrder:
                    if (!string.Equals(saved.Pages[command.PageNumber - 1].TabOrder, GetTabOrderName((PdfPageTabOrder)command.Number), StringComparison.Ordinal)) throw new InvalidOperationException("AcroForm page tab-order readback validation failed.");
                    break;
                case PdfAcroFormEditSession.EditKind.CalculationOrder:
                    string[] expectedOrder = command.Names!.Distinct(StringComparer.Ordinal).ToArray();
                    if (!calculationOrder.SequenceEqual(expectedOrder, StringComparer.Ordinal)) throw new InvalidOperationException("AcroForm calculation-order readback validation failed.");
                    break;
                case PdfAcroFormEditSession.EditKind.Flatten:
                    for (int n = 0; n < command.Names!.Length; n++) if (byName.ContainsKey(command.Names[n])) throw new InvalidOperationException("AcroForm flatten readback validation failed for " + command.Names[n] + ".");
                    break;
            }
        }
    }

    private static bool IsRemovedLater(IReadOnlyList<PdfAcroFormEditSession.EditCommand> commands, int index, string name) {
        string current = name;
        for (int i = index + 1; i < commands.Count; i++) {
            PdfAcroFormEditSession.EditCommand command = commands[i];
            if (command.Kind == PdfAcroFormEditSession.EditKind.Rename && string.Equals(command.Name, current, StringComparison.Ordinal)) current = command.Value!;
            if (command.Kind == PdfAcroFormEditSession.EditKind.Remove && string.Equals(command.Name, current, StringComparison.Ordinal)) return true;
            if (command.Kind == PdfAcroFormEditSession.EditKind.Flatten && command.Names!.Contains(current, StringComparer.Ordinal)) return true;
        }
        return !string.Equals(current, name, StringComparison.Ordinal);
    }

    private static PdfFormFieldKind ToFieldKind(PdfFormFieldCreationKind kind) => kind == PdfFormFieldCreationKind.Text ? PdfFormFieldKind.Text : kind == PdfFormFieldCreationKind.Choice ? PdfFormFieldKind.Choice : kind == PdfFormFieldCreationKind.Signature ? PdfFormFieldKind.Signature : PdfFormFieldKind.Button;
}
