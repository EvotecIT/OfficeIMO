namespace OfficeIMO.Pdf;

internal static partial class PdfAcroFormEditor {
    private static void ApplyCommands(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        int[] pageObjectNumbers,
        IReadOnlyList<PdfAcroFormEditSession.EditCommand> commands,
        Dictionary<string, string> refillValues,
        List<string> flattenNames,
        List<string> operations,
        PdfFormFillerOptions? appearanceOptions) {
        PdfDictionary catalog = RequireCatalog(objects, security);
        PdfDictionary acroForm = EnsureAcroForm(objects, catalog, out PdfArray fields);
        int nextObjectNumber = objects.Count == 0 ? 1 : objects.Keys.Max() + 1;

        foreach (PdfAcroFormEditSession.EditCommand command in commands) {
            switch (command.Kind) {
                case PdfAcroFormEditSession.EditKind.Create:
                    ApplyCreate(objects, acroForm, fields, pageObjectNumbers, command.Options!, refillValues, appearanceOptions, ref nextObjectNumber);
                    operations.Add("Create " + command.Options!.Name);
                    break;
                case PdfAcroFormEditSession.EditKind.Rename:
                    ApplyRename(objects, fields, command.Name!, command.Value!, refillValues, flattenNames);
                    operations.Add("Rename " + command.Name + " -> " + command.Value);
                    break;
                case PdfAcroFormEditSession.EditKind.Remove:
                    ApplyRemove(objects, acroForm, fields, command.Name!);
                    RemoveQueuedSubtreeWork(refillValues, flattenNames, command.Name!);
                    operations.Add("Remove " + command.Name);
                    break;
                case PdfAcroFormEditSession.EditKind.Move:
                    ApplyMove(objects, acroForm, fields, pageObjectNumbers, command.Name!, command.PageNumber, command.Rectangle!, refillValues, appearanceOptions, ref nextObjectNumber);
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

    private static void ApplyCreate(Dictionary<int, PdfIndirectObject> objects, PdfDictionary acroForm, PdfArray fields, int[] pages, PdfFormFieldCreateOptions options, Dictionary<string, string> refillValues, PdfFormFillerOptions? appearanceOptions, ref int nextObjectNumber) {
        ValidateCreateOptions(options, pages.Length);
        if (FieldPathExists(objects, fields, options.Name)) throw new ArgumentException("PDF form field already exists: " + options.Name, nameof(options));
        (PdfArray fieldOwner, PdfReference? parentReference, string partialName) = EnsureCreatedFieldOwner(objects, fields, options.Name, ref nextObjectNumber);
        string appearanceFontName = EnsureAcroFormAppearanceDefaults(objects, acroForm);
        if (options.Kind == PdfFormFieldCreationKind.RadioButtonGroup) {
            ApplyCreateRadioButtonGroup(objects, acroForm, fieldOwner, parentReference, partialName, pages, options, appearanceFontName, refillValues, appearanceOptions, ref nextObjectNumber);
            if (!acroForm.Items.ContainsKey("NeedAppearances")) acroForm.Items["NeedAppearances"] = new PdfBoolean(false);
            return;
        }

        PdfDictionary page = RequirePage(objects, pages, options.PageNumber);
        int objectNumber = nextObjectNumber++;
        var field = new PdfDictionary();
        field.Items["Type"] = new PdfName("Annot"); field.Items["Subtype"] = new PdfName("Widget");
        field.Items["FT"] = new PdfName(GetFieldType(options.Kind)); field.Items["T"] = new PdfStringObj(partialName, true);
        if (parentReference is not null) field.Items["Parent"] = parentReference;
        field.Items["Rect"] = CreateRectangle(options.X, options.Y, options.X + options.Width, options.Y + options.Height);
        field.Items["P"] = CreateReference(objects, pages[options.PageNumber - 1]); field.Items["F"] = new PdfNumber(options.WidgetFlags);
        int fieldFlags = GetCreateFieldFlags(options);
        if (fieldFlags != 0) field.Items["Ff"] = new PdfNumber(fieldFlags);
        if (options.Kind == PdfFormFieldCreationKind.Choice) field.Items["Opt"] = CreateStringArray(options.ChoiceOptions);
        ApplyCreateFieldStyle(field, options, appearanceFontName, includeWidgetStyle: true);
        if (options.Kind != PdfFormFieldCreationKind.PushButton) {
            string initialValue = ResolveInitialValue(options);
            SetFieldValue(field, GetFieldType(options.Kind), initialValue, options.CheckedValueName, setAppearanceState: true);
            SetDefaultValue(field, GetFieldType(options.Kind), options.DefaultValue, options.CheckedValueName, normalizeButtonValue: true);
        }
        ApplyWidgetJavaScript(field, options.JavaScript, usePrimaryAction: options.Kind == PdfFormFieldCreationKind.PushButton);
        if (options.Kind == PdfFormFieldCreationKind.CheckBox) {
            AddCheckBoxAppearances(objects, field, options, ref nextObjectNumber);
        }
        if (options.Kind == PdfFormFieldCreationKind.Choice && string.IsNullOrEmpty(ResolveInitialValue(options))) {
            AddTextWidgetAppearance(objects, acroForm, page, field, options, string.Empty, appearanceOptions, ref nextObjectNumber);
        }
        objects[objectNumber] = new PdfIndirectObject(objectNumber, 0, field);
        var reference = new PdfReference(objectNumber, 0); fieldOwner.Items.Add(reference); EnsureAnnotationArray(objects, page).Items.Add(reference);
        if (options.Kind == PdfFormFieldCreationKind.PushButton) {
            AddPushButtonAppearance(objects, acroForm, page, field, options, appearanceOptions, ref nextObjectNumber);
        } else if (options.Kind != PdfFormFieldCreationKind.Signature) {
            bool normalizeEmptyMultiSelect = options.Kind == PdfFormFieldCreationKind.Choice &&
                                             (fieldFlags & FieldFlagMultiSelect) != 0;
            QueueRefillValue(refillValues, options.Name, GetFieldType(options.Kind), ReadSimpleValue(field), includeEmptyChoice: normalizeEmptyMultiSelect);
        }
        if (!acroForm.Items.ContainsKey("NeedAppearances")) acroForm.Items["NeedAppearances"] = new PdfBoolean(false);
    }

    private static void ApplyRename(Dictionary<int, PdfIndirectObject> objects, PdfArray fields, string name, string newName, Dictionary<string, string> refillValues, List<string> flattenNames) {
        if (FieldPathExists(objects, fields, newName)) throw new ArgumentException("PDF form field already exists: " + newName, nameof(newName));
        EditableField field = RequireField(objects, fields, name);
        string oldParent = ParentName(name); string newParent = ParentName(newName);
        if (!string.Equals(oldParent, newParent, StringComparison.Ordinal)) throw new NotSupportedException("Renaming a hierarchical field must preserve its parent path.");
        string partialName = ReadText(field.Dictionary, "T") ?? string.Empty;
        field.Dictionary.Items["T"] = new PdfStringObj(string.Equals(partialName, field.FullName, StringComparison.Ordinal) ? newName : LeafName(newName), true);
        string? value = refillValues.TryGetValue(name, out string? queuedValue)
            ? queuedValue
            : ReadInheritedSimpleValue(objects, field.Dictionary);
        refillValues.Remove(name);
        QueueRefillValue(refillValues, newName, field.FieldType, value);
        for (int i = 0; i < flattenNames.Count; i++) {
            if (string.Equals(flattenNames[i], name, StringComparison.Ordinal)) {
                flattenNames[i] = newName;
            } else if (flattenNames[i].StartsWith(name + ".", StringComparison.Ordinal)) {
                flattenNames[i] = newName + flattenNames[i].Remove(0, name.Length);
            }
        }
    }

    private static void ApplyRemove(Dictionary<int, PdfIndirectObject> objects, PdfDictionary acroForm, PdfArray fields, string name) {
        EditableField field = RequireFieldSubtree(objects, fields, name);
        field.Owner.Items.Remove(field.Reference);
        var removed = new HashSet<int>(field.ObjectNumbers);
        RemoveWidgetReferences(objects, removed);
        FilterReferenceArray(objects, acroForm, "CO", removed);
        foreach (int objectNumber in removed) objects.Remove(objectNumber);
        RemoveEmptyParents(objects, fields);
    }

    private static void RemoveQueuedSubtreeWork(Dictionary<string, string> refillValues, List<string> flattenNames, string name) {
        string descendantPrefix = name + ".";
        foreach (string queuedName in refillValues.Keys
                     .Where(candidate => string.Equals(candidate, name, StringComparison.Ordinal) || candidate.StartsWith(descendantPrefix, StringComparison.Ordinal))
                     .ToArray()) {
            refillValues.Remove(queuedName);
        }
        flattenNames.RemoveAll(candidate => string.Equals(candidate, name, StringComparison.Ordinal) || candidate.StartsWith(descendantPrefix, StringComparison.Ordinal));
    }

    private static void ApplyMove(Dictionary<int, PdfIndirectObject> objects, PdfDictionary acroForm, PdfArray fields, int[] pages, string name, int pageNumber, double[] rectangle, Dictionary<string, string> refillValues, PdfFormFillerOptions? appearanceOptions, ref int nextObjectNumber) {
        EditableField field = RequireField(objects, fields, name);
        if (field.WidgetObjectNumbers.Count != 1) throw new NotSupportedException("Moving a form field requires exactly one indirect widget.");
        PdfDictionary widget = RequireDictionary(objects, field.WidgetObjectNumbers[0]);
        bool pushButtonSizeChanged = IsPushButton(objects, field) && !HasSameRectangleSize(objects, widget, rectangle);
        if (pushButtonSizeChanged && HasPushButtonIcon(objects, widget)) {
            throw new NotSupportedException("Resizing a push button with an icon is not supported because its icon and layout semantics cannot be preserved safely.");
        }
        if (pushButtonSizeChanged && HasPushButtonInteractiveAppearances(objects, widget)) {
            throw new NotSupportedException("Resizing a push button with rollover or down appearances is not supported because every interactive appearance state must preserve the authored bounds.");
        }
        PdfDictionary page = RequirePage(objects, pages, pageNumber);
        PdfArray destinationAnnotations = EnsureAnnotationArray(objects, page);
        int destinationIndex = FindReferenceIndex(destinationAnnotations, field.WidgetObjectNumbers[0]);
        RemoveWidgetReferences(objects, new HashSet<int>(field.WidgetObjectNumbers));
        widget.Items["P"] = CreateReference(objects, pages[pageNumber - 1]); widget.Items["Rect"] = CreateRectangle(rectangle[0], rectangle[1], rectangle[2], rectangle[3]);
        destinationAnnotations = EnsureAnnotationArray(objects, page);
        PdfReference widgetReference = CreateReference(objects, field.WidgetObjectNumbers[0]);
        if (destinationIndex >= 0) {
            destinationAnnotations.Items.Insert(Math.Min(destinationIndex, destinationAnnotations.Items.Count), widgetReference);
        } else {
            destinationAnnotations.Items.Add(widgetReference);
        }
        if (IsPushButton(objects, field)) {
            if (pushButtonSizeChanged) {
                RebuildPushButtonAppearance(objects, acroForm, page, field, widget, rectangle, appearanceOptions, ref nextObjectNumber);
            }
            refillValues.Remove(name);
            return;
        }
        QueueRefillValue(refillValues, name, field.FieldType, ReadInheritedSimpleValue(objects, field.Dictionary), includeEmptyChoice: true);
    }

    private static bool IsPushButton(Dictionary<int, PdfIndirectObject> objects, EditableField field) =>
        string.Equals(field.FieldType, "Btn", StringComparison.Ordinal) &&
        (ReadInheritedFieldFlags(objects, field.Dictionary) & FieldFlagPushButton) != 0;

    private static int FindReferenceIndex(PdfArray array, int objectNumber) {
        for (int index = 0; index < array.Items.Count; index++) {
            if (array.Items[index] is PdfReference reference && reference.ObjectNumber == objectNumber) return index;
        }
        return -1;
    }

    private static bool HasPushButtonIcon(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget) {
        PdfDictionary? characteristics = ResolveDictionary(
            objects,
            widget.Items.TryGetValue("MK", out PdfObject? characteristicsObject) ? characteristicsObject : null);
        return characteristics != null &&
            (characteristics.Items.ContainsKey("I") ||
             characteristics.Items.ContainsKey("RI") ||
             characteristics.Items.ContainsKey("IX"));
    }

    private static bool HasPushButtonInteractiveAppearances(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget) {
        PdfDictionary? appearances = ResolveDictionary(
            objects,
            widget.Items.TryGetValue("AP", out PdfObject? appearanceObject) ? appearanceObject : null);
        return appearances != null &&
            (appearances.Items.TryGetValue("R", out PdfObject? rollover) && PdfObjectLookup.Resolve(objects, rollover) is not PdfNull ||
             appearances.Items.TryGetValue("D", out PdfObject? down) && PdfObjectLookup.Resolve(objects, down) is not PdfNull);
    }

    private static bool HasSameRectangleSize(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary widget,
        double[] rectangle) {
        if (!widget.Items.TryGetValue("Rect", out PdfObject? rectangleObject) ||
            PdfObjectLookup.Resolve(objects, rectangleObject) is not PdfArray existing ||
            existing.Items.Count != 4 ||
            existing.Items[0] is not PdfNumber x1 || existing.Items[1] is not PdfNumber y1 ||
            existing.Items[2] is not PdfNumber x2 || existing.Items[3] is not PdfNumber y2) {
            return false;
        }
        double oldWidth = Math.Abs(x2.Value - x1.Value);
        double oldHeight = Math.Abs(y2.Value - y1.Value);
        double newWidth = Math.Abs(rectangle[2] - rectangle[0]);
        double newHeight = Math.Abs(rectangle[3] - rectangle[1]);
        return Math.Abs(oldWidth - newWidth) <= 1e-9D && Math.Abs(oldHeight - newHeight) <= 1e-9D;
    }

    private static int ReadInheritedFieldFlags(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary field) {
        var visited = new HashSet<PdfDictionary>();
        PdfDictionary? current = field;
        while (current is not null && visited.Add(current)) {
            if (current.Items.TryGetValue("Ff", out PdfObject? flagsObject) &&
                PdfObjectLookup.Resolve(objects, flagsObject) is PdfNumber flags &&
                flags.Value >= int.MinValue && flags.Value <= int.MaxValue &&
                Math.Truncate(flags.Value) == flags.Value) {
                return (int)flags.Value;
            }
            current = current.Items.TryGetValue("Parent", out PdfObject? parentObject)
                ? ResolveDictionary(objects, parentObject)
                : null;
        }
        return 0;
    }

    private static void RebuildPushButtonAppearance(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary acroForm,
        PdfDictionary page,
        EditableField field,
        PdfDictionary widget,
        double[] rectangle,
        PdfFormFillerOptions? appearanceOptions,
        ref int nextObjectNumber) {
        double width = rectangle[2] - rectangle[0];
        double height = rectangle[3] - rectangle[1];
        string? defaultAppearance = ReadResolvedText(objects, widget, "DA") ?? ReadResolvedText(objects, field.Dictionary, "DA") ?? ReadResolvedText(objects, acroForm, "DA");
        PdfFormFieldStyle style = PdfAcroFormEditor.CreateButtonCaptionStyle(
            PdfFormFiller.ReadWidgetAppearanceStyle(objects, widget, inheritedDefaultAppearance: defaultAppearance));
        style.TextAlignment = PdfFormFieldTextAlignment.Center;
        string caption = ResolveDictionary(objects, widget.Items.TryGetValue("MK", out PdfObject? characteristicsObject) ? characteristicsObject : null) is PdfDictionary characteristics
            ? ReadResolvedText(objects, characteristics, "CA") ?? string.Empty
            : string.Empty;
        PdfStream appearance = PdfFormFiller.CreateAuthoredTextWidgetAppearance(
            objects,
            acroForm,
            page,
            widget,
            caption,
            width,
            height,
            style,
            PdfFormFiller.ReadWidgetAppearanceFontSize(defaultAppearance, height),
            field.FullName,
            appearanceOptions,
            ref nextObjectNumber,
            inheritedDefaultAppearance: defaultAppearance);
        int appearanceObjectNumber = nextObjectNumber++;
        objects[appearanceObjectNumber] = new PdfIndirectObject(appearanceObjectNumber, 0, appearance);
        PdfDictionary? appearances = widget.Items.TryGetValue("AP", out PdfObject? appearanceObject)
            ? ResolveDictionary(objects, appearanceObject)
            : null;
        if (appearances is null) {
            appearances = new PdfDictionary();
        } else {
            var detachedAppearances = new PdfDictionary();
            foreach (KeyValuePair<string, PdfObject> item in appearances.Items) {
                detachedAppearances.Items[item.Key] = item.Value;
            }
            appearances = detachedAppearances;
        }
        widget.Items["AP"] = appearances;
        appearances.Items["N"] = new PdfReference(appearanceObjectNumber, 0);
    }

    private static string? ReadResolvedText(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) =>
        dictionary.Items.TryGetValue(key, out PdfObject? value) && PdfObjectLookup.Resolve(objects, value) is PdfStringObj text ? text.Value : null;

    private static void ApplyDefaultValue(Dictionary<int, PdfIndirectObject> objects, PdfArray fields, string name, string? value) {
        EditableField field = RequireField(objects, fields, name);
        int flags = ReadInheritedFieldFlags(objects, field.Dictionary);
        if (value is not null &&
            string.Equals(field.FieldType, "Btn", StringComparison.Ordinal) &&
            (flags & FieldFlagPushButton) == 0 &&
            !string.Equals(value, "Off", StringComparison.Ordinal) &&
            !CollectButtonAppearanceStates(objects, field.Dictionary, new HashSet<int>()).Contains(value)) {
            throw new ArgumentException($"PDF button field cannot use default value '{value}' because it is not one of the available appearance states.", nameof(value));
        }
        if (value is not null &&
            string.Equals(field.FieldType, "Ch", StringComparison.Ordinal) &&
            (flags & (FieldFlagCombo | FieldFlagEdit)) != (FieldFlagCombo | FieldFlagEdit) &&
            !CollectChoiceExportValues(objects, field.Dictionary).Contains(value)) {
            throw new ArgumentException($"PDF choice field cannot use default value '{value}' because it is not one of the available option values.", nameof(value));
        }
        SetDefaultValue(field.Dictionary, field.FieldType, value, "Yes", normalizeButtonValue: false);
    }

    private static HashSet<string> CollectChoiceExportValues(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary field) {
        var visited = new HashSet<PdfDictionary>();
        PdfDictionary? current = field;
        while (current is not null && visited.Add(current)) {
            if (current.Items.TryGetValue("Opt", out PdfObject? optionsObject)) {
                var result = new HashSet<string>(StringComparer.Ordinal);
                if (PdfObjectLookup.Resolve(objects, optionsObject) is not PdfArray options) return result;
                for (int index = 0; index < options.Items.Count; index++) {
                    PdfObject? option = PdfObjectLookup.Resolve(objects, options.Items[index]);
                    if (option is PdfStringObj text) {
                        result.Add(text.Value);
                    } else if (option is PdfArray pair &&
                               pair.Items.Count > 0 &&
                               PdfObjectLookup.Resolve(objects, pair.Items[0]) is PdfStringObj exportValue) {
                        result.Add(exportValue.Value);
                    }
                }
                return result;
            }
            current = current.Items.TryGetValue("Parent", out PdfObject? parentObject)
                ? ResolveDictionary(objects, parentObject)
                : null;
        }
        return new HashSet<string>(StringComparer.Ordinal);
    }

    private static HashSet<string> CollectButtonAppearanceStates(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary field,
        HashSet<int> visited) {
        var states = new HashSet<string>(StringComparer.Ordinal);
        Collect(field);
        return states;

        void Collect(PdfDictionary current) {
            if (current.Items.TryGetValue("AP", out PdfObject? appearanceObject) &&
                ResolveDictionary(objects, appearanceObject) is PdfDictionary appearances &&
                appearances.Items.TryGetValue("N", out PdfObject? normalObject) &&
                ResolveDictionary(objects, normalObject) is PdfDictionary normalAppearances) {
                foreach (string state in normalAppearances.Items.Keys) states.Add(state);
            }
            if (!current.Items.TryGetValue("Kids", out PdfObject? kidsObject) ||
                PdfObjectLookup.Resolve(objects, kidsObject) is not PdfArray kids) return;
            foreach (PdfObject kidObject in kids.Items) {
                if (kidObject is PdfReference reference && !visited.Add(reference.ObjectNumber)) continue;
                if (PdfObjectLookup.Resolve(objects, kidObject) is PdfDictionary kid) Collect(kid);
            }
        }
    }

    private static void ApplyFlags(Dictionary<int, PdfIndirectObject> objects, PdfArray fields, string name, int flags, Dictionary<string, string> refillValues) {
        EditableField field = RequireField(objects, fields, name);
        int previousFlags = ReadInheritedFieldFlags(objects, field.Dictionary);
        if (string.Equals(field.FieldType, "Btn", StringComparison.Ordinal) &&
            (previousFlags & FieldFlagPushButton) != 0 &&
            (flags & FieldFlagPushButton) == 0) {
            throw new NotSupportedException("Clearing the push-button flag is not supported because it changes the field's button semantics.");
        }
        if (string.Equals(field.FieldType, "Btn", StringComparison.Ordinal) &&
            (previousFlags & FieldFlagPushButton) == 0 &&
            (flags & FieldFlagPushButton) != 0) {
            throw new NotSupportedException("Converting a check box or radio button to a push button is not supported because it requires rebuilding the complete button state.");
        }
        if (string.Equals(field.FieldType, "Btn", StringComparison.Ordinal) &&
            ((previousFlags ^ flags) & FieldFlagRadio) != 0) {
            throw new NotSupportedException("Changing between check-box and radio-button semantics is not supported because it requires rebuilding the complete button state.");
        }
        if (string.Equals(field.FieldType, "Tx", StringComparison.Ordinal) &&
            (flags & FieldFlagComb) != 0 &&
            (!HasInheritedPositiveInteger(objects, field.Dictionary, "MaxLen") ||
             (flags & (FieldFlagMultiline | FieldFlagPassword | FieldFlagFileSelect)) != 0)) {
            throw new NotSupportedException("PDF comb text fields require MaxLen and cannot also be multiline, password, or file-select fields.");
        }
        if (string.Equals(field.FieldType, "Ch", StringComparison.Ordinal)) {
            bool isComboBox = (flags & FieldFlagCombo) != 0;
            bool isEditable = (flags & FieldFlagEdit) != 0;
            bool isMultiSelect = (flags & FieldFlagMultiSelect) != 0;
            if (isEditable && !isComboBox) {
                throw new NotSupportedException("Editable choice fields must be combo boxes.");
            }
            if (isComboBox && isMultiSelect) {
                throw new NotSupportedException("Combo-box choice fields cannot also be multi-select fields.");
            }
            if (!isMultiSelect &&
                (ReadInheritedFieldObject(objects, field.Dictionary, "V") is PdfArray ||
                 ReadInheritedFieldObject(objects, field.Dictionary, "DV") is PdfArray)) {
                throw new NotSupportedException("Clearing multi-select is not supported while the choice field stores an array value as its current or default value.");
            }
        }
        field.Dictionary.Items["Ff"] = new PdfNumber(flags);
        if (string.Equals(field.FieldType, "Btn", StringComparison.Ordinal) && (flags & FieldFlagPushButton) != 0) {
            refillValues.Remove(name);
            return;
        }
        QueueRefillValue(refillValues, name, field.FieldType, ReadInheritedSimpleValue(objects, field.Dictionary), includeEmptyChoice: true);
    }

    private static bool HasInheritedPositiveInteger(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary field,
        string key) {
        var visited = new HashSet<PdfDictionary>();
        PdfDictionary? current = field;
        while (current is not null && visited.Add(current)) {
            if (current.Items.TryGetValue(key, out PdfObject? value) &&
                PdfObjectLookup.Resolve(objects, value) is PdfNumber number) {
                return number.Value > 0D && number.Value <= int.MaxValue && Math.Truncate(number.Value) == number.Value;
            }
            current = current.Items.TryGetValue("Parent", out PdfObject? parentObject)
                ? ResolveDictionary(objects, parentObject)
                : null;
        }
        return false;
    }

    private static PdfObject? ReadInheritedFieldValue(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary field) =>
        ReadInheritedFieldObject(objects, field, "V");

    private static PdfObject? ReadInheritedFieldObject(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary field,
        string key) {
        var visited = new HashSet<PdfDictionary>();
        PdfDictionary? current = field;
        while (current is not null && visited.Add(current)) {
            if (current.Items.TryGetValue(key, out PdfObject? value)) {
                return PdfObjectLookup.Resolve(objects, value);
            }
            current = current.Items.TryGetValue("Parent", out PdfObject? parentObject)
                ? ResolveDictionary(objects, parentObject)
                : null;
        }
        return null;
    }

    private static void QueueRefillValue(Dictionary<string, string> refillValues, string name, string? fieldType, string? value, bool includeEmptyChoice = false) {
        if (value is null || string.Equals(fieldType, "Sig", StringComparison.Ordinal) ||
            !includeEmptyChoice && string.Equals(fieldType, "Ch", StringComparison.Ordinal) && value.Length == 0) {
            return;
        }
        refillValues[name] = value;
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
