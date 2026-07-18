namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private static void CollectFlattenWidgets(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject fieldObject,
        string? inheritedFieldType,
        int inheritedFlags,
        int? inheritedMaxLength,
        PdfDictionary? inheritedDefaultResources,
        string? inheritedDefaultAppearance,
        string? inheritedDisplayValue,
        IReadOnlyList<PdfFreeTextRichTextRun>? inheritedRichAppearanceRuns,
        string? inheritedName,
        PdfArray? inheritedChoiceOptions,
        PdfFormFillerOptions? options,
        Dictionary<int, FlattenWidgetState> widgets,
        HashSet<int> removableObjects,
        HashSet<int> visited,
        ref int nextObjectNumber) {
        int? fieldObjectNumber = null;
        if (fieldObject is PdfReference reference) {
            fieldObjectNumber = reference.ObjectNumber;
            if (!visited.Add(reference.ObjectNumber)) {
                return;
            }
        }

        if (ResolveObject(objects, fieldObject) is not PdfDictionary field) {
            return;
        }

        if (fieldObjectNumber.HasValue) {
            removableObjects.Add(fieldObjectNumber.Value);
        }

        string? partialName = TryReadText(objects, field, "T");
        string? fullName = CombineFieldName(inheritedName, partialName);
        string? fieldType = TryReadName(objects, field, "FT") ?? inheritedFieldType;
        int fieldFlags = ReadFieldFlags(objects, field, inheritedFlags);
        int? fieldMaxLength = ReadFieldMaxLength(objects, field, inheritedMaxLength);
        PdfDictionary? defaultResources = TryReadDefaultResources(objects, field) ?? inheritedDefaultResources;
        string? defaultAppearance = TryReadText(objects, field, "DA") ?? inheritedDefaultAppearance;
        PdfArray? choiceOptions = TryReadChoiceOptions(objects, field) ?? inheritedChoiceOptions;
        IReadOnlyList<string>? values = TryReadSimpleValues(objects, field, "V");
        string? richValue = TryReadText(objects, field, "RV");
        string? richPlainText = PdfFreeTextStyleParser.ExtractPlainText(richValue);
        IReadOnlyList<PdfFreeTextRichTextRun>? richAppearanceRuns = PdfFreeTextStyleParser.ExtractRichTextRuns(richValue) ?? inheritedRichAppearanceRuns;
        string? value = values is { Count: > 0 } ? values[0] : richPlainText ?? inheritedDisplayValue;
        bool isButtonField = string.Equals(fieldType, "Btn", StringComparison.Ordinal);
        bool isRadioButtonGroup = isButtonField && (fieldFlags & RadioButtonFlag) != 0;
        bool isMultiSelectChoice = string.Equals(fieldType, "Ch", StringComparison.Ordinal) && (fieldFlags & MultiSelectChoiceFlag) != 0;
        string choiceDisplaySeparator = isMultiSelectChoice ? "\n" : ", ";
        string? appearanceValue = string.Equals(fieldType, "Ch", StringComparison.Ordinal)
            ? TryResolveChoiceDisplayValue(objects, choiceOptions, values, choiceDisplaySeparator) ?? JoinSimpleValues(values, choiceDisplaySeparator) ?? inheritedDisplayValue
            : value;

        if (IsWidget(field)) {
            if (!fieldObjectNumber.HasValue ||
                !TryReadRectCoordinates(field, out double x, out double y, out double width, out double height)) {
                throw new NotSupportedException(UnsupportedFlattenWidgetMessage);
            }

            int appearanceObjectNumber;
            if (isButtonField) {
                string appearanceState = GetButtonWidgetFlattenAppearanceState(objects, field, value);
                if (!TryGetButtonAppearanceReference(objects, field, appearanceState, out PdfReference? appearanceReference)) {
                    EnsureButtonWidgetAppearances(objects, field, appearanceState, isRadioButtonGroup, ref nextObjectNumber);
                    if (!TryGetButtonAppearanceReference(objects, field, appearanceState, out appearanceReference)) {
                        throw new NotSupportedException(UnsupportedFlattenWidgetMessage);
                    }
                }

                appearanceObjectNumber = appearanceReference!.ObjectNumber;
            } else if (TryGetNormalAppearanceReference(objects, field, out PdfReference? appearanceReference)) {
                appearanceObjectNumber = appearanceReference!.ObjectNumber;
            } else {
                PdfDictionary? widgetAppearanceResources = TryReadNormalAppearanceResources(objects, field);
                PdfDictionary? widgetPageResources = TryReadWidgetPageResources(objects, field);
                PdfFormFieldStyle widgetStyle = ReadWidgetAppearanceStyle(objects, field, fieldFlags, inheritedMaxLength: fieldMaxLength, inheritedDefaultAppearance: defaultAppearance);
                if (isMultiSelectChoice) {
                    widgetStyle.IsMultiline = true;
                }

                appearanceObjectNumber = nextObjectNumber++;
                objects[appearanceObjectNumber] = new PdfIndirectObject(appearanceObjectNumber, 0, CreateTextAppearanceStream(objects, defaultResources, widgetAppearanceResources, widgetPageResources, appearanceValue ?? string.Empty, width, height, widgetStyle, defaultAppearance, ReadWidgetAppearanceFontSize(defaultAppearance, height), options, fullName, ref nextObjectNumber, richAppearanceRuns));
            }

            widgets[fieldObjectNumber.Value] = new FlattenWidgetState(fieldObjectNumber.Value, x, y, width, height, appearanceObjectNumber);
            return;
        }

        if (!field.Items.TryGetValue("Kids", out var kidsObject) ||
            ResolveObject(objects, kidsObject) is not PdfArray kids) {
            throw new NotSupportedException(UnsupportedFlattenWidgetMessage);
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            CollectFlattenWidgets(objects, kids.Items[i], fieldType, fieldFlags, fieldMaxLength, defaultResources, defaultAppearance, appearanceValue, richAppearanceRuns, fullName, choiceOptions, options, widgets, removableObjects, visited, ref nextObjectNumber);
        }
    }

    private static int FlattenPageWidgets(Dictionary<int, PdfIndirectObject> objects, Dictionary<int, FlattenWidgetState> widgets, ref int nextObjectNumber) {
        int flattenedWidgetCount = 0;
        foreach (var entry in objects.OrderBy(pair => pair.Key).ToArray()) {
            if (entry.Value.Value is not PdfDictionary page ||
                page.Get<PdfName>("Type")?.Name != "Page" ||
                !page.Items.TryGetValue("Annots", out var annotsObject) ||
                ResolveObject(objects, annotsObject) is not PdfArray annots) {
                continue;
            }

            var pageWidgets = new List<FlattenWidgetState>();
            var remainingAnnots = new PdfArray();
            for (int i = 0; i < annots.Items.Count; i++) {
                PdfObject annot = annots.Items[i];
                if (annot is PdfReference annotReference && widgets.TryGetValue(annotReference.ObjectNumber, out var widget)) {
                    pageWidgets.Add(widget);
                    flattenedWidgetCount++;
                } else {
                    remainingAnnots.Items.Add(annot);
                }
            }

            if (pageWidgets.Count == 0) {
                continue;
            }

            if (remainingAnnots.Items.Count == 0) {
                page.Items.Remove("Annots");
            } else {
                page.Items["Annots"] = remainingAnnots;
            }

            string content = BuildFlattenContent(objects, page, pageWidgets);
            int contentObjectNumber = nextObjectNumber++;
            objects[contentObjectNumber] = new PdfIndirectObject(contentObjectNumber, 0, CreateContentStream(content));
            AppendPageContent(objects, page, contentObjectNumber);
        }

        return flattenedWidgetCount;
    }

    private static string BuildFlattenContent(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page, List<FlattenWidgetState> widgets) {
        PdfDictionary xObjects = EnsurePageXObjects(objects, page);
        var builder = new StringBuilder();
        for (int i = 0; i < widgets.Count; i++) {
            FlattenWidgetState widget = widgets[i];
            string xObjectName = CreateUniqueXObjectName(xObjects);
            xObjects.Items[xObjectName] = new PdfReference(widget.AppearanceObjectNumber, 0);
            builder.Append("q\n");
            builder.Append(FormatNumber(widget.Width)).Append(" 0 0 ").Append(FormatNumber(widget.Height)).Append(' ')
                .Append(FormatNumber(widget.X)).Append(' ').Append(FormatNumber(widget.Y)).Append(" cm\n");
            builder.Append('/').Append(xObjectName).Append(" Do\n");
            builder.Append("Q\n");
        }

        return builder.ToString();
    }

    private static PdfDictionary EnsurePageXObjects(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page) {
        return PdfPageResourceHelper.EnsurePageXObjects(objects, page, "form flattening");
    }

    private static string CreateUniqueXObjectName(PdfDictionary xObjects) {
        int index = 1;
        string name;
        do {
            name = "OfficeIMOForm" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
            index++;
        } while (xObjects.Items.ContainsKey(name));

        return name;
    }

    private static PdfStream CreateContentStream(string content) {
        var dictionary = new PdfDictionary();
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static void AppendPageContent(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page, int contentObjectNumber) {
        var newReference = new PdfReference(contentObjectNumber, 0);
        if (!page.Items.TryGetValue("Contents", out var contents)) {
            page.Items["Contents"] = newReference;
            return;
        }

        if (contents is PdfArray contentsArray) {
            contentsArray.Items.Add(newReference);
            return;
        }

        var array = new PdfArray();
        AppendContentEntries(objects, array, contents);
        array.Items.Add(newReference);
        page.Items["Contents"] = array;
    }

    private static void AppendContentEntries(Dictionary<int, PdfIndirectObject> objects, PdfArray target, PdfObject contents) {
        if (contents is PdfArray directArray) {
            foreach (var item in directArray.Items) {
                target.Items.Add(item);
            }

            return;
        }

        if (contents is PdfReference reference &&
            PdfObjectLookup.TryGet(objects, reference, out var indirect) &&
            indirect.Value is PdfArray referencedArray) {
            foreach (var item in referencedArray.Items) {
                target.Items.Add(item);
            }

            return;
        }

        target.Items.Add(contents);
    }
}
