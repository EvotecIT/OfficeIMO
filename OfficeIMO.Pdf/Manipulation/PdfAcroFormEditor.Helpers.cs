namespace OfficeIMO.Pdf;

internal static partial class PdfAcroFormEditor {
    private sealed class EditableField {
        internal EditableField(string fullName, string? fieldType, PdfDictionary dictionary, PdfArray owner, PdfObject reference, IReadOnlyList<int> widgetObjectNumbers, IReadOnlyList<int> objectNumbers) {
            FullName = fullName; FieldType = fieldType; Dictionary = dictionary; Owner = owner; Reference = reference; WidgetObjectNumbers = widgetObjectNumbers; ObjectNumbers = objectNumbers;
        }
        internal string FullName { get; }
        internal string? FieldType { get; }
        internal PdfDictionary Dictionary { get; }
        internal PdfArray Owner { get; }
        internal PdfObject Reference { get; }
        internal IReadOnlyList<int> WidgetObjectNumbers { get; }
        internal IReadOnlyList<int> ObjectNumbers { get; }
    }

    private static PdfDictionary RequireCatalog(Dictionary<int, PdfIndirectObject> objects, PdfDocumentSecurityInfo security) {
        if (!security.RootObjectNumber.HasValue || !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? root) || root.Value is not PdfDictionary catalog) throw new InvalidOperationException("PDF catalog is not readable.");
        return catalog;
    }

    private static PdfDictionary EnsureAcroForm(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog, out PdfArray fields) {
        PdfDictionary? acroForm = catalog.Items.TryGetValue("AcroForm", out PdfObject? value) ? ResolveDictionary(objects, value) : null;
        if (acroForm is null) {
            int objectNumber = objects.Count == 0 ? 1 : objects.Keys.Max() + 1;
            acroForm = new PdfDictionary(); objects[objectNumber] = new PdfIndirectObject(objectNumber, 0, acroForm); catalog.Items["AcroForm"] = new PdfReference(objectNumber, 0);
        }
        bool hasFields = acroForm.Items.TryGetValue("Fields", out PdfObject? fieldsObject);
        fields = hasFields ? ResolveArray(objects, fieldsObject) ?? new PdfArray() : new PdfArray();
        if (!hasFields || ResolveArray(objects, fieldsObject) is null) acroForm.Items["Fields"] = fields;
        return acroForm;
    }

    private static EditableField RequireField(Dictionary<int, PdfIndirectObject> objects, PdfArray fields, string name) => FindField(objects, fields, name) ?? throw new ArgumentException("PDF form field was not found: " + name, nameof(name));

    private static EditableField? FindField(Dictionary<int, PdfIndirectObject> objects, PdfArray fields, string name) {
        var found = new List<EditableField>();
        CollectFields(objects, fields, null, null, found, new HashSet<int>());
        EditableField? match = null;
        for (int i = 0; i < found.Count; i++) {
            if (!string.Equals(found[i].FullName, name, StringComparison.Ordinal)) continue;
            if (match is not null) throw new InvalidOperationException("PDF contains duplicate fully qualified form field names: " + name);
            match = found[i];
        }
        return match;
    }

    private static IReadOnlyList<string> ReadCalculationOrder(byte[] pdf) {
        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf);
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(pdf).Map;
        PdfDictionary catalog = RequireCatalog(objects, security);
        if (!catalog.Items.TryGetValue("AcroForm", out PdfObject? acroFormObject) || ResolveDictionary(objects, acroFormObject) is not PdfDictionary acroForm ||
            !acroForm.Items.TryGetValue("Fields", out PdfObject? fieldsObject) || ResolveArray(objects, fieldsObject) is not PdfArray fields ||
            !acroForm.Items.TryGetValue("CO", out PdfObject? orderObject) || ResolveArray(objects, orderObject) is not PdfArray order) return Array.Empty<string>();
        var editable = new List<EditableField>();
        CollectFields(objects, fields, null, null, editable, new HashSet<int>());
        var names = editable.Where(static field => field.Reference is PdfReference).ToDictionary(static field => ((PdfReference)field.Reference).ObjectNumber, static field => field.FullName);
        var result = new List<string>(order.Items.Count);
        for (int i = 0; i < order.Items.Count; i++) {
            if (order.Items[i] is not PdfReference reference || !names.TryGetValue(reference.ObjectNumber, out string? name)) throw new InvalidOperationException("AcroForm calculation order references an unreadable field.");
            result.Add(name);
        }
        return result.AsReadOnly();
    }

    private static void CollectFields(Dictionary<int, PdfIndirectObject> objects, PdfArray owner, string? parentName, string? inheritedType, List<EditableField> result, HashSet<int> visited) {
        for (int i = 0; i < owner.Items.Count; i++) {
            PdfObject fieldObject = owner.Items[i];
            if (fieldObject is not PdfReference reference || !visited.Add(reference.ObjectNumber)) throw new NotSupportedException("Transactional AcroForm editing requires an acyclic indirect field tree.");
            PdfDictionary field = RequireDictionary(objects, reference.ObjectNumber);
            string? partialName = ReadText(field, "T"); string? fullName = CombineName(parentName, partialName);
            string? fieldType = ReadName(field, "FT") ?? inheritedType;
            PdfArray? kids = field.Items.TryGetValue("Kids", out PdfObject? kidsObject) ? ResolveArray(objects, kidsObject) : null;
            bool hasNamedFieldKids = false;
            if (kids is not null) for (int k = 0; k < kids.Items.Count; k++) {
                PdfDictionary? kid = ResolveDictionary(objects, kids.Items[k]);
                if (kid is not null && !string.IsNullOrEmpty(ReadText(kid, "T"))) { hasNamedFieldKids = true; break; }
            }
            if (kids is not null && hasNamedFieldKids) {
                CollectFields(objects, kids, fullName, fieldType, result, visited);
                continue;
            }
            if (string.IsNullOrEmpty(fullName)) continue;
            var widgetNumbers = new List<int>(); var objectNumbers = new List<int>();
            CollectSubtreeObjects(objects, fieldObject, objectNumbers, widgetNumbers, new HashSet<int>());
            result.Add(new EditableField(fullName!, fieldType, field, owner, fieldObject, widgetNumbers.AsReadOnly(), objectNumbers.AsReadOnly()));
        }
    }

    private static void CollectSubtreeObjects(Dictionary<int, PdfIndirectObject> objects, PdfObject value, List<int> objectNumbers, List<int> widgetNumbers, HashSet<int> visited) {
        if (value is not PdfReference reference || !visited.Add(reference.ObjectNumber)) return;
        objectNumbers.Add(reference.ObjectNumber);
        PdfDictionary field = RequireDictionary(objects, reference.ObjectNumber);
        if (string.Equals(ReadName(field, "Subtype"), "Widget", StringComparison.Ordinal)) widgetNumbers.Add(reference.ObjectNumber);
        if (field.Items.TryGetValue("Kids", out PdfObject? kidsObject) && ResolveArray(objects, kidsObject) is PdfArray kids)
            for (int i = 0; i < kids.Items.Count; i++) CollectSubtreeObjects(objects, kids.Items[i], objectNumbers, widgetNumbers, visited);
    }

    private static PdfDictionary RequirePage(Dictionary<int, PdfIndirectObject> objects, int[] pageObjectNumbers, int pageNumber) {
        if (pageNumber < 1 || pageNumber > pageObjectNumbers.Length) throw new ArgumentOutOfRangeException(nameof(pageNumber), "Page number is outside the document.");
        return RequireDictionary(objects, pageObjectNumbers[pageNumber - 1]);
    }

    private static PdfDictionary RequireDictionary(Dictionary<int, PdfIndirectObject> objects, int objectNumber) {
        if (!objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect) || indirect.Value is not PdfDictionary dictionary) throw new InvalidOperationException("Required PDF dictionary object is missing: " + objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture));
        return dictionary;
    }

    private static PdfReference CreateReference(Dictionary<int, PdfIndirectObject> objects, int objectNumber) {
        if (!objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect)) throw new InvalidOperationException("Required PDF object is missing.");
        return new PdfReference(indirect.ObjectNumber, indirect.Generation);
    }

    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) => PdfObjectLookup.Resolve(objects, value) as PdfDictionary;
    private static PdfArray? ResolveArray(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) => PdfObjectLookup.Resolve(objects, value) as PdfArray;

    private static PdfArray EnsureAnnotationArray(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page) {
        PdfArray? annots = page.Items.TryGetValue("Annots", out PdfObject? value) ? ResolveArray(objects, value) : null;
        if (annots is null) { annots = new PdfArray(); page.Items["Annots"] = annots; }
        return annots;
    }

    private static void RemoveWidgetReferences(Dictionary<int, PdfIndirectObject> objects, HashSet<int> widgetNumbers) {
        foreach (PdfIndirectObject indirect in objects.Values) {
            if (indirect.Value is not PdfDictionary page || !string.Equals(ReadName(page, "Type"), "Page", StringComparison.Ordinal) || !page.Items.TryGetValue("Annots", out PdfObject? annotsObject) || ResolveArray(objects, annotsObject) is not PdfArray annots) continue;
            for (int i = annots.Items.Count - 1; i >= 0; i--) if (annots.Items[i] is PdfReference reference && widgetNumbers.Contains(reference.ObjectNumber)) annots.Items.RemoveAt(i);
            if (annots.Items.Count == 0) page.Items.Remove("Annots");
        }
    }

    private static void FilterReferenceArray(Dictionary<int, PdfIndirectObject> objects, PdfDictionary owner, string key, HashSet<int> removed) {
        if (!owner.Items.TryGetValue(key, out PdfObject? value) || ResolveArray(objects, value) is not PdfArray array) return;
        for (int i = array.Items.Count - 1; i >= 0; i--) if (array.Items[i] is PdfReference reference && removed.Contains(reference.ObjectNumber)) array.Items.RemoveAt(i);
        if (array.Items.Count == 0) owner.Items.Remove(key);
    }

    private static void RemoveEmptyParents(Dictionary<int, PdfIndirectObject> objects, PdfArray owner) {
        for (int i = owner.Items.Count - 1; i >= 0; i--) {
            PdfDictionary? field = ResolveDictionary(objects, owner.Items[i]);
            if (field is null || !field.Items.TryGetValue("Kids", out PdfObject? kidsObject) || ResolveArray(objects, kidsObject) is not PdfArray kids) continue;
            RemoveEmptyParents(objects, kids);
            if (kids.Items.Count == 0 && !string.Equals(ReadName(field, "Subtype"), "Widget", StringComparison.Ordinal)) owner.Items.RemoveAt(i);
        }
    }

    private static PdfArray CreateRectangle(double x1, double y1, double x2, double y2) { var result = new PdfArray(); result.Items.Add(new PdfNumber(x1)); result.Items.Add(new PdfNumber(y1)); result.Items.Add(new PdfNumber(x2)); result.Items.Add(new PdfNumber(y2)); return result; }
    private static PdfArray CreateStringArray(IReadOnlyList<string> values) { var result = new PdfArray(); for (int i = 0; i < values.Count; i++) result.Items.Add(new PdfStringObj(values[i], true)); return result; }
    private static string? ReadText(PdfDictionary dictionary, string key) => dictionary.Items.TryGetValue(key, out PdfObject? value) && value is PdfStringObj text ? text.Value : null;
    private static string? ReadName(PdfDictionary dictionary, string key) => dictionary.Items.TryGetValue(key, out PdfObject? value) && value is PdfName name ? name.Name : null;
    private static string? ReadSimpleValue(PdfDictionary dictionary) => dictionary.Items.TryGetValue("V", out PdfObject? value) ? value is PdfStringObj text ? text.Value : value is PdfName name ? name.Name : null : null;
    private static string? CombineName(string? parent, string? partial) => string.IsNullOrEmpty(partial) ? parent : string.IsNullOrEmpty(parent) ? partial : parent + "." + partial;
    private static string ParentName(string name) { int index = name.LastIndexOf('.'); return index < 0 ? string.Empty : name.Substring(0, index); }
    private static string LeafName(string name) { int index = name.LastIndexOf('.'); return index < 0 ? name : name.Substring(index + 1); }
    private static string GetFieldType(PdfFormFieldCreationKind kind) => kind == PdfFormFieldCreationKind.Text ? "Tx" : kind == PdfFormFieldCreationKind.Choice ? "Ch" : kind == PdfFormFieldCreationKind.Signature ? "Sig" : "Btn";
    private static string GetTabOrderName(PdfPageTabOrder order) => order == PdfPageTabOrder.Row ? "R" : order == PdfPageTabOrder.Column ? "C" : order == PdfPageTabOrder.Structure ? "S" : "A";

    private static void SetFieldValue(PdfDictionary field, string fieldType, string? value, string checkedName, bool setAppearanceState) {
        if (string.Equals(fieldType, "Sig", StringComparison.Ordinal)) return;
        if (string.Equals(fieldType, "Btn", StringComparison.Ordinal)) {
            string state = IsChecked(value, checkedName) ? checkedName : "Off"; field.Items["V"] = new PdfName(state); if (setAppearanceState) field.Items["AS"] = new PdfName(state); return;
        }
        field.Items["V"] = new PdfStringObj(value ?? string.Empty, true);
    }

    private static void SetDefaultValue(PdfDictionary field, string? fieldType, string? value, string checkedName, bool normalizeButtonValue) {
        if (value is null) { field.Items.Remove("DV"); return; }
        field.Items["DV"] = string.Equals(fieldType, "Btn", StringComparison.Ordinal) ? new PdfName(normalizeButtonValue ? IsChecked(value, checkedName) ? checkedName : "Off" : value) : new PdfStringObj(value, true);
    }

    private static bool IsChecked(string? value, string checkedName) => string.Equals(value, checkedName, StringComparison.Ordinal) || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) || string.Equals(value, "1", StringComparison.Ordinal);

    private static void ValidateCreateOptions(PdfFormFieldCreateOptions options, int pageCount) {
        Guard.NotNullOrWhiteSpace(options.Name, nameof(options.Name));
        if (options.PageNumber < 1 || options.PageNumber > pageCount) throw new ArgumentOutOfRangeException(nameof(options), "Page number is outside the document.");
        if (!IsFinite(options.X) || !IsFinite(options.Y) || !IsFinite(options.Width) || !IsFinite(options.Height) || options.Width <= 0D || options.Height <= 0D) throw new ArgumentOutOfRangeException(nameof(options), "Field rectangle must contain finite coordinates and positive dimensions.");
        if (options.Name[0] == '.' || options.Name[options.Name.Length - 1] == '.') throw new ArgumentException("Field name cannot start or end with a period.", nameof(options));
        if (options.Kind == PdfFormFieldCreationKind.Choice && options.ChoiceOptions.Any(string.IsNullOrEmpty)) throw new ArgumentException("Choice options cannot be empty.", nameof(options));
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
