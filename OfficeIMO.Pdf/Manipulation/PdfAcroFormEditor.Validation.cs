namespace OfficeIMO.Pdf;

internal static partial class PdfAcroFormEditor {
    private static void ValidateReadback(PdfDocumentInfo saved, IReadOnlyList<string> calculationOrder, IReadOnlyList<PdfAcroFormEditSession.EditCommand> commands) {
        var byName = saved.FormFields.Where(static field => !string.IsNullOrEmpty(field.Name)).ToDictionary(static field => field.Name!, StringComparer.Ordinal);
        for (int i = 0; i < commands.Count; i++) {
            PdfAcroFormEditSession.EditCommand command = commands[i];
            switch (command.Kind) {
                case PdfAcroFormEditSession.EditKind.Create:
                    if (!IsRemovedLater(commands, i, command.Options!.Name)) {
                        if (!byName.TryGetValue(command.Options.Name, out PdfFormField? created) || created.Kind != ToFieldKind(command.Options.Kind)) throw new InvalidOperationException("AcroForm create readback validation failed for " + command.Options.Name + ".");
                        ValidateCreatedFieldReadback(created, command.Options, !HasLaterFlagsEdit(commands, i, command.Options.Name));
                    }
                    break;
                case PdfAcroFormEditSession.EditKind.Rename:
                    if (!IsRemovedLater(commands, i, command.Value!) &&
                        !IsFlattenedInTransaction(commands, command.Value!) &&
                        !byName.ContainsKey(command.Value!)) throw new InvalidOperationException("AcroForm rename readback validation failed for " + command.Value + ".");
                    break;
                case PdfAcroFormEditSession.EditKind.Remove:
                    if (byName.Keys.Any(candidate => IsFieldInSubtree(candidate, command.Name!) && !IsIntroducedLater(commands, i, candidate))) throw new InvalidOperationException("AcroForm remove readback validation failed for " + command.Name + ".");
                    break;
                case PdfAcroFormEditSession.EditKind.DefaultValue:
                    if (!HasLaterDefaultValueEdit(commands, i, command.Name!) &&
                        byName.TryGetValue(command.Name!, out PdfFormField? defaultField) &&
                        !string.Equals(defaultField.DefaultValue, command.Value, StringComparison.Ordinal)) throw new InvalidOperationException("AcroForm default-value readback validation failed for " + command.Name + ".");
                    break;
                case PdfAcroFormEditSession.EditKind.Flags:
                    if (!HasLaterFlagsEdit(commands, i, command.Name!) &&
                        byName.TryGetValue(command.Name!, out PdfFormField? flagsField) &&
                        flagsField.Flags != command.Number) throw new InvalidOperationException("AcroForm flags readback validation failed for " + command.Name + ".");
                    break;
                case PdfAcroFormEditSession.EditKind.TabOrder:
                    if (!HasLaterTabOrderEdit(commands, i, command.PageNumber) &&
                        !string.Equals(saved.Pages[command.PageNumber - 1].TabOrder, GetTabOrderName((PdfPageTabOrder)command.Number), StringComparison.Ordinal)) throw new InvalidOperationException("AcroForm page tab-order readback validation failed.");
                    break;
                case PdfAcroFormEditSession.EditKind.CalculationOrder:
                    string[] expectedOrder = ResolveFinalCalculationOrder(command.Names!, commands, i);
                    if (!HasLaterCalculationOrderEdit(commands, i) &&
                        !calculationOrder.SequenceEqual(expectedOrder, StringComparer.Ordinal)) throw new InvalidOperationException("AcroForm calculation-order readback validation failed.");
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
            if (command.Kind == PdfAcroFormEditSession.EditKind.Remove && IsFieldInSubtree(current, command.Name!)) return true;
            if (command.Kind == PdfAcroFormEditSession.EditKind.Flatten && command.Names!.Contains(current, StringComparer.Ordinal)) return true;
        }
        return !string.Equals(current, name, StringComparison.Ordinal);
    }

    private static bool IsIntroducedLater(IReadOnlyList<PdfAcroFormEditSession.EditCommand> commands, int index, string finalName) {
        for (int i = index + 1; i < commands.Count; i++) {
            PdfAcroFormEditSession.EditCommand command = commands[i];
            if (command.Kind == PdfAcroFormEditSession.EditKind.Create && string.Equals(command.Options!.Name, finalName, StringComparison.Ordinal)) return true;
            if (command.Kind == PdfAcroFormEditSession.EditKind.Rename && string.Equals(command.Value, finalName, StringComparison.Ordinal)) return true;
        }
        return false;
    }

    private static bool IsFlattenedInTransaction(IReadOnlyList<PdfAcroFormEditSession.EditCommand> commands, string finalName) {
        for (int index = 0; index < commands.Count; index++) {
            PdfAcroFormEditSession.EditCommand command = commands[index];
            if (command.Kind != PdfAcroFormEditSession.EditKind.Flatten) continue;
            for (int nameIndex = 0; nameIndex < command.Names!.Length; nameIndex++) {
                string current = command.Names[nameIndex];
                for (int later = index + 1; later < commands.Count; later++) {
                    PdfAcroFormEditSession.EditCommand laterCommand = commands[later];
                    if (laterCommand.Kind == PdfAcroFormEditSession.EditKind.Rename && string.Equals(laterCommand.Name, current, StringComparison.Ordinal)) {
                        current = laterCommand.Value!;
                    }
                }
                if (string.Equals(current, finalName, StringComparison.Ordinal)) return true;
            }
        }
        return false;
    }

    private static bool HasLaterFlagsEdit(IReadOnlyList<PdfAcroFormEditSession.EditCommand> commands, int index, string name) {
        string current = name;
        for (int i = index + 1; i < commands.Count; i++) {
            PdfAcroFormEditSession.EditCommand command = commands[i];
            if (command.Kind == PdfAcroFormEditSession.EditKind.Rename && string.Equals(command.Name, current, StringComparison.Ordinal)) {
                current = command.Value!;
            } else if (command.Kind == PdfAcroFormEditSession.EditKind.Flags && string.Equals(command.Name, current, StringComparison.Ordinal)) {
                return true;
            } else if (command.Kind == PdfAcroFormEditSession.EditKind.Remove && IsFieldInSubtree(current, command.Name!)) {
                return false;
            } else if (command.Kind == PdfAcroFormEditSession.EditKind.Flatten && command.Names!.Contains(current, StringComparer.Ordinal)) {
                return false;
            }
        }
        return false;
    }

    private static bool HasLaterDefaultValueEdit(IReadOnlyList<PdfAcroFormEditSession.EditCommand> commands, int index, string name) {
        string current = name;
        for (int i = index + 1; i < commands.Count; i++) {
            PdfAcroFormEditSession.EditCommand command = commands[i];
            if (command.Kind == PdfAcroFormEditSession.EditKind.Rename && string.Equals(command.Name, current, StringComparison.Ordinal)) {
                current = command.Value!;
            } else if (command.Kind == PdfAcroFormEditSession.EditKind.DefaultValue && string.Equals(command.Name, current, StringComparison.Ordinal)) {
                return true;
            } else if (command.Kind == PdfAcroFormEditSession.EditKind.Remove && IsFieldInSubtree(current, command.Name!)) {
                return false;
            } else if (command.Kind == PdfAcroFormEditSession.EditKind.Flatten && command.Names!.Contains(current, StringComparer.Ordinal)) {
                return false;
            }
        }
        return false;
    }

    private static bool HasLaterTabOrderEdit(IReadOnlyList<PdfAcroFormEditSession.EditCommand> commands, int index, int pageNumber) {
        for (int i = index + 1; i < commands.Count; i++) {
            PdfAcroFormEditSession.EditCommand command = commands[i];
            if (command.Kind == PdfAcroFormEditSession.EditKind.TabOrder && command.PageNumber == pageNumber) return true;
        }
        return false;
    }

    private static bool HasLaterCalculationOrderEdit(IReadOnlyList<PdfAcroFormEditSession.EditCommand> commands, int index) {
        for (int i = index + 1; i < commands.Count; i++) {
            if (commands[i].Kind == PdfAcroFormEditSession.EditKind.CalculationOrder) return true;
        }
        return false;
    }

    private static string[] ResolveFinalCalculationOrder(
        IReadOnlyList<string> names,
        IReadOnlyList<PdfAcroFormEditSession.EditCommand> commands,
        int index) {
        var expected = names.Distinct(StringComparer.Ordinal).ToList();
        for (int i = index + 1; i < commands.Count; i++) {
            PdfAcroFormEditSession.EditCommand command = commands[i];
            if (command.Kind == PdfAcroFormEditSession.EditKind.Rename) {
                for (int nameIndex = 0; nameIndex < expected.Count; nameIndex++) {
                    if (string.Equals(expected[nameIndex], command.Name, StringComparison.Ordinal)) {
                        expected[nameIndex] = command.Value!;
                    } else if (expected[nameIndex].StartsWith(command.Name + ".", StringComparison.Ordinal)) {
                        expected[nameIndex] = command.Value + expected[nameIndex].Remove(0, command.Name!.Length);
                    }
                }
            } else if (command.Kind == PdfAcroFormEditSession.EditKind.Remove) {
                expected.RemoveAll(name => IsFieldInSubtree(name, command.Name!));
            } else if (command.Kind == PdfAcroFormEditSession.EditKind.Flatten) {
                expected.RemoveAll(name => command.Names!.Any(flattened => IsFieldInSubtree(name, flattened)));
            }
        }
        return expected.Distinct(StringComparer.Ordinal).ToArray();
    }

    private static bool IsFieldInSubtree(string fieldName, string subtreeName) =>
        string.Equals(fieldName, subtreeName, StringComparison.Ordinal) ||
        fieldName.StartsWith(subtreeName + ".", StringComparison.Ordinal);

    private static PdfFormFieldKind ToFieldKind(PdfFormFieldCreationKind kind) => kind == PdfFormFieldCreationKind.Text ? PdfFormFieldKind.Text : kind == PdfFormFieldCreationKind.Choice ? PdfFormFieldKind.Choice : kind == PdfFormFieldCreationKind.Signature ? PdfFormFieldKind.Signature : PdfFormFieldKind.Button;

    private static void ValidateCreatedFieldReadback(PdfFormField field, PdfFormFieldCreateOptions options, bool validateCreationFlags) {
        if (validateCreationFlags && options.Kind == PdfFormFieldCreationKind.Text && options.Style?.IsMultiline == true && !field.IsMultiline) throw new InvalidOperationException("AcroForm multiline text-field readback validation failed for " + options.Name + ".");
        if (validateCreationFlags && options.Kind == PdfFormFieldCreationKind.Choice && field.IsCombo != IsChoiceComboBox(options)) throw new InvalidOperationException("AcroForm choice presentation readback validation failed for " + options.Name + ".");
        if (options.Kind == PdfFormFieldCreationKind.RadioButtonGroup && (validateCreationFlags && !field.IsRadioButton || field.WidgetCount != options.ChoiceOptions.Count)) throw new InvalidOperationException("AcroForm radio-button readback validation failed for " + options.Name + ".");
        if (validateCreationFlags && options.Kind == PdfFormFieldCreationKind.PushButton && !field.IsPushButton) throw new InvalidOperationException("AcroForm push-button readback validation failed for " + options.Name + ".");
        if (options.JavaScript is not null && !string.Equals(field.JavaScript, options.JavaScript, StringComparison.Ordinal)) throw new InvalidOperationException("AcroForm widget JavaScript readback validation failed for " + options.Name + ".");
    }
}
