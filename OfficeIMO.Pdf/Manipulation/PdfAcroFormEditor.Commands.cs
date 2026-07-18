namespace OfficeIMO.Pdf;

internal static partial class PdfAcroFormEditor {
    private static void ApplyCommands(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        int[] pageObjectNumbers,
        IReadOnlyList<PdfAcroFormEditSession.EditCommand> commands,
        Dictionary<string, string> refillValues,
        List<string> flattenNames,
        List<string> operations) {
        PdfDictionary catalog = RequireCatalog(objects, security);
        PdfDictionary acroForm = EnsureAcroForm(objects, catalog, out PdfArray fields);
        int nextObjectNumber = objects.Count == 0 ? 1 : objects.Keys.Max() + 1;

        foreach (PdfAcroFormEditSession.EditCommand command in commands) {
            switch (command.Kind) {
                case PdfAcroFormEditSession.EditKind.Create:
                    ApplyCreate(objects, acroForm, fields, pageObjectNumbers, command.Options!, refillValues, ref nextObjectNumber);
                    operations.Add("Create " + command.Options!.Name);
                    break;
                case PdfAcroFormEditSession.EditKind.Rename:
                    ApplyRename(objects, fields, command.Name!, command.Value!, refillValues);
                    operations.Add("Rename " + command.Name + " -> " + command.Value);
                    break;
                case PdfAcroFormEditSession.EditKind.Remove:
                    ApplyRemove(objects, acroForm, fields, command.Name!);
                    refillValues.Remove(command.Name!);
                    operations.Add("Remove " + command.Name);
                    break;
                case PdfAcroFormEditSession.EditKind.Move:
                    ApplyMove(objects, fields, pageObjectNumbers, command.Name!, command.PageNumber, command.Rectangle!, refillValues);
                    operations.Add("Move " + command.Name + " to page " + command.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    break;
                case PdfAcroFormEditSession.EditKind.DefaultValue:
                    ApplyDefaultValue(objects, fields, command.Name!, command.Value);
                    operations.Add("Set default " + command.Name);
                    break;
                case PdfAcroFormEditSession.EditKind.Flags:
                    ApplyFlags(objects, fields, command.Name!, command.Number, refillValues);
                    operations.Add("Set flags " + command.Name);
                    break;
                case PdfAcroFormEditSession.EditKind.CalculationOrder:
                    ApplyCalculationOrder(objects, acroForm, fields, command.Names!);
                    operations.Add("Set calculation order");
                    break;
                case PdfAcroFormEditSession.EditKind.TabOrder:
                    RequirePage(objects, pageObjectNumbers, command.PageNumber).Items["Tabs"] = new PdfName(GetTabOrderName((PdfPageTabOrder)command.Number));
                    operations.Add("Set page " + command.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " tab order");
                    break;
                case PdfAcroFormEditSession.EditKind.Flatten:
                    for (int i = 0; i < command.Names!.Length; i++) {
                        EditableField field = RequireField(objects, fields, command.Names[i]);
                        if (string.Equals(field.FieldType, "Sig", StringComparison.Ordinal)) throw new NotSupportedException("Signature fields cannot be flattened by the AcroForm editor.");
                        if (!flattenNames.Contains(field.FullName, StringComparer.Ordinal)) flattenNames.Add(field.FullName);
                    }
                    operations.Add("Flatten " + string.Join(", ", command.Names));
                    break;
            }
        }
    }

    private static void ApplyCreate(Dictionary<int, PdfIndirectObject> objects, PdfDictionary acroForm, PdfArray fields, int[] pages, PdfFormFieldCreateOptions options, Dictionary<string, string> refillValues, ref int nextObjectNumber) {
        ValidateCreateOptions(options, pages.Length);
        if (FindField(objects, fields, options.Name) is not null) throw new ArgumentException("PDF form field already exists: " + options.Name, nameof(options));
        PdfDictionary page = RequirePage(objects, pages, options.PageNumber);
        int objectNumber = nextObjectNumber++;
        var field = new PdfDictionary();
        field.Items["Type"] = new PdfName("Annot"); field.Items["Subtype"] = new PdfName("Widget");
        field.Items["FT"] = new PdfName(GetFieldType(options.Kind)); field.Items["T"] = new PdfStringObj(options.Name, true);
        field.Items["Rect"] = CreateRectangle(options.X, options.Y, options.X + options.Width, options.Y + options.Height);
        field.Items["P"] = CreateReference(objects, pages[options.PageNumber - 1]); field.Items["F"] = new PdfNumber(options.WidgetFlags);
        if (options.FieldFlags != 0) field.Items["Ff"] = new PdfNumber(options.FieldFlags);
        if (options.Kind == PdfFormFieldCreationKind.Choice) field.Items["Opt"] = CreateStringArray(options.ChoiceOptions);
        SetFieldValue(field, GetFieldType(options.Kind), options.Value, options.CheckedValueName, setAppearanceState: true);
        SetDefaultValue(field, GetFieldType(options.Kind), options.DefaultValue, options.CheckedValueName, normalizeButtonValue: true);
        objects[objectNumber] = new PdfIndirectObject(objectNumber, 0, field);
        var reference = new PdfReference(objectNumber, 0); fields.Items.Add(reference); EnsureAnnotationArray(objects, page).Items.Add(reference);
        if (options.Kind != PdfFormFieldCreationKind.Signature) refillValues[options.Name] = ReadSimpleValue(field) ?? string.Empty;
        if (!acroForm.Items.ContainsKey("NeedAppearances")) acroForm.Items["NeedAppearances"] = new PdfBoolean(false);
    }

    private static void ApplyRename(Dictionary<int, PdfIndirectObject> objects, PdfArray fields, string name, string newName, Dictionary<string, string> refillValues) {
        if (FindField(objects, fields, newName) is not null) throw new ArgumentException("PDF form field already exists: " + newName, nameof(newName));
        EditableField field = RequireField(objects, fields, name);
        string oldParent = ParentName(name); string newParent = ParentName(newName);
        if (!string.Equals(oldParent, newParent, StringComparison.Ordinal)) throw new NotSupportedException("Renaming a hierarchical field must preserve its parent path.");
        string partialName = ReadText(field.Dictionary, "T") ?? string.Empty;
        field.Dictionary.Items["T"] = new PdfStringObj(string.Equals(partialName, field.FullName, StringComparison.Ordinal) ? newName : LeafName(newName), true);
        string? value = ReadSimpleValue(field.Dictionary);
        refillValues.Remove(name); if (!string.Equals(field.FieldType, "Sig", StringComparison.Ordinal) && value is not null) refillValues[newName] = value;
    }

    private static void ApplyRemove(Dictionary<int, PdfIndirectObject> objects, PdfDictionary acroForm, PdfArray fields, string name) {
        EditableField field = RequireField(objects, fields, name);
        field.Owner.Items.Remove(field.Reference);
        var removed = new HashSet<int>(field.ObjectNumbers);
        RemoveWidgetReferences(objects, removed);
        FilterReferenceArray(objects, acroForm, "CO", removed);
        foreach (int objectNumber in removed) objects.Remove(objectNumber);
        RemoveEmptyParents(objects, fields);
    }

    private static void ApplyMove(Dictionary<int, PdfIndirectObject> objects, PdfArray fields, int[] pages, string name, int pageNumber, double[] rectangle, Dictionary<string, string> refillValues) {
        EditableField field = RequireField(objects, fields, name);
        if (field.WidgetObjectNumbers.Count != 1) throw new NotSupportedException("Moving a form field requires exactly one indirect widget.");
        PdfDictionary widget = RequireDictionary(objects, field.WidgetObjectNumbers[0]);
        RemoveWidgetReferences(objects, new HashSet<int>(field.WidgetObjectNumbers));
        PdfDictionary page = RequirePage(objects, pages, pageNumber);
        widget.Items["P"] = CreateReference(objects, pages[pageNumber - 1]); widget.Items["Rect"] = CreateRectangle(rectangle[0], rectangle[1], rectangle[2], rectangle[3]);
        EnsureAnnotationArray(objects, page).Items.Add(new PdfReference(field.WidgetObjectNumbers[0], 0));
        string? value = ReadSimpleValue(field.Dictionary); if (!string.Equals(field.FieldType, "Sig", StringComparison.Ordinal) && value is not null) refillValues[name] = value;
    }

    private static void ApplyDefaultValue(Dictionary<int, PdfIndirectObject> objects, PdfArray fields, string name, string? value) {
        EditableField field = RequireField(objects, fields, name);
        SetDefaultValue(field.Dictionary, field.FieldType, value, "Yes", normalizeButtonValue: false);
    }

    private static void ApplyFlags(Dictionary<int, PdfIndirectObject> objects, PdfArray fields, string name, int flags, Dictionary<string, string> refillValues) {
        EditableField field = RequireField(objects, fields, name); field.Dictionary.Items["Ff"] = new PdfNumber(flags);
        string? value = ReadSimpleValue(field.Dictionary); if (!string.Equals(field.FieldType, "Sig", StringComparison.Ordinal) && value is not null) refillValues[name] = value;
    }

    private static void ApplyCalculationOrder(Dictionary<int, PdfIndirectObject> objects, PdfDictionary acroForm, PdfArray fields, string[] names) {
        var order = new PdfArray(); var seen = new HashSet<int>();
        for (int i = 0; i < names.Length; i++) {
            EditableField field = RequireField(objects, fields, names[i]);
            if (field.Reference is not PdfReference reference) throw new NotSupportedException("Calculation-order fields must be indirect objects.");
            if (seen.Add(reference.ObjectNumber)) order.Items.Add(reference);
        }
        if (order.Items.Count == 0) acroForm.Items.Remove("CO"); else acroForm.Items["CO"] = order;
    }
}
