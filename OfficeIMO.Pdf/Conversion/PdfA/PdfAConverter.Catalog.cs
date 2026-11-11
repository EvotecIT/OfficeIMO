using System.Globalization;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfAConverter {
    private static readonly HashSet<string> AllowedActions = new(StringComparer.Ordinal) {
        "GoTo", "GoToR", "GoToE", "Thread", "URI", "Named", "SubmitForm"
    };

    private static readonly HashSet<string> AllowedNamedActions = new(StringComparer.Ordinal) {
        "NextPage", "PrevPage", "FirstPage", "LastPage"
    };

    private void ProcessCatalog(PdfDictionary catalog) {
        if (catalog.Items.ContainsKey("AA") &&
            Change(PdfAConversionChangeKind.ActionRemoved, "catalog", "Removed the document additional-actions (AA), which PDF/A forbids.", catalog, OfficeConversionLossKind.Omission)) {
            catalog.Items.Remove("AA");
        }
        if (catalog.Items.TryGetValue("OpenAction", out PdfObject? openAction) && ResolveDictionary(openAction) is not null) {
            PdfObject? kept = SanitizeAction(openAction, "catalog open action", 0);
            if (kept is null) catalog.Items.Remove("OpenAction");
            else catalog.Items["OpenAction"] = kept;
        }

        if (ResolveDictionary(catalog.Items.TryGetValue("Names", out PdfObject? names) ? names : null) is PdfDictionary nameDictionary) {
            if (nameDictionary.Items.ContainsKey("JavaScript") &&
                Change(PdfAConversionChangeKind.ActionRemoved, "catalog names", "Removed the document JavaScript name tree, which PDF/A forbids.", nameDictionary, OfficeConversionLossKind.Omission)) {
                nameDictionary.Items.Remove("JavaScript");
            }
            if (nameDictionary.Items.ContainsKey("AlternatePresentations") &&
                Change(PdfAConversionChangeKind.StructureUpdated, "catalog names", "Removed alternate presentations (slide shows), which PDF/A forbids.", nameDictionary, OfficeConversionLossKind.Omission)) {
                nameDictionary.Items.Remove("AlternatePresentations");
            }
        }

        if (catalog.Items.ContainsKey("Perms") &&
            Change(PdfAConversionChangeKind.StructureUpdated, "catalog", "Removed the permissions dictionary (usage rights), which a rewritten file cannot keep valid.", catalog, OfficeConversionLossKind.Omission)) {
            catalog.Items.Remove("Perms");
        }

        if (GetName(catalog, "Version") is string version && string.CompareOrdinal(version, "1.7") > 0 &&
            Change(PdfAConversionChangeKind.StructureUpdated, "catalog", "Removed the PDF " + version + " version override; PDF/A-2 and PDF/A-3 files are PDF 1.7.", catalog)) {
            catalog.Items.Remove("Version");
        }
        ProcessAcroForm(catalog);
        ProcessOptionalContent(catalog);
        ProcessOutlines(catalog);
    }

    /// <summary>Returns the action with forbidden links removed from its chain, or <see langword="null"/> when nothing remains.</summary>
    private PdfObject? SanitizeAction(PdfObject actionObject, string location, int depth) {
        if (depth > 64 || ResolveDictionary(actionObject) is not PdfDictionary action) return actionObject;
        if (action.Items.TryGetValue("Next", out PdfObject? next)) {
            PdfObject? resolvedNext = Resolve(next);
            var kept = new List<PdfObject>();
            IEnumerable<PdfObject> chain = resolvedNext is PdfArray array ? array.Items : new[] { next };
            foreach (PdfObject item in chain) {
                PdfObject? sanitized = SanitizeAction(item, location, depth + 1);
                if (sanitized is not null) kept.Add(sanitized);
            }
            if (!_verifyOnly) {
                if (kept.Count == 0) action.Items.Remove("Next");
                else if (kept.Count == 1) action.Items["Next"] = kept[0];
                else {
                    var nextArray = new PdfArray();
                    nextArray.Items.AddRange(kept);
                    action.Items["Next"] = nextArray;
                }
            }
        }

        string type = GetName(action, "S") ?? string.Empty;
        bool allowed = AllowedActions.Contains(type) && (type != "Named" || AllowedNamedActions.Contains(GetName(action, "N") ?? string.Empty));
        if (allowed) return actionObject;
        string description = type == "Named" ? "named action /" + (GetName(action, "N") ?? "?") : (type.Length == 0 ? "untyped" : type) + " action";
        if (!Change(PdfAConversionChangeKind.ActionRemoved, location, "Removed a " + description + ", which PDF/A forbids.", action, OfficeConversionLossKind.Omission)) {
            return actionObject;
        }
        if (action.Items.TryGetValue("Next", out PdfObject? remaining)) {
            PdfObject? resolvedRemaining = Resolve(remaining);
            if (resolvedRemaining is PdfArray remainingArray) {
                if (remainingArray.Items.Count == 0) return null;
                PdfObject first = remainingArray.Items[0];
                if (remainingArray.Items.Count > 1 && ResolveDictionary(first) is PdfDictionary firstAction) {
                    var rest = new PdfArray();
                    rest.Items.AddRange(remainingArray.Items.Skip(1));
                    if (firstAction.Items.TryGetValue("Next", out PdfObject? firstNext)) {
                        if (Resolve(firstNext) is PdfArray firstNextArray) rest.Items.InsertRange(0, firstNextArray.Items);
                        else rest.Items.Insert(0, firstNext);
                    }
                    firstAction.Items["Next"] = rest;
                }
                return first;
            }
            return remaining;
        }
        return null;
    }

    private void ProcessAcroForm(PdfDictionary catalog) {
        if (ResolveDictionary(catalog.Items.TryGetValue("AcroForm", out PdfObject? form) ? form : null) is not PdfDictionary acroForm) return;
        _acroForm = acroForm;
        _needAppearances = Resolve(acroForm.Items.TryGetValue("NeedAppearances", out PdfObject? needs) ? needs : null) is PdfBoolean { Value: true };
        if (acroForm.Items.ContainsKey("XFA")) {
            if (Resolve(catalog.Items.TryGetValue("NeedsRendering", out PdfObject? needsRendering) ? needsRendering : null) is PdfBoolean { Value: true }) {
                AddViolation(PdfAConversionRefusalReason.DynamicXfaForm, "interactive form",
                    "The form is a dynamic XFA form; its pages exist only in XFA, which PDF/A forbids.", acroForm);
            } else if (Change(PdfAConversionChangeKind.FormUpdated, "interactive form",
                    "Removed the XFA form data, which PDF/A forbids; the AcroForm fields and their appearances are kept.", acroForm, OfficeConversionLossKind.Omission)) {
                acroForm.Items.Remove("XFA");
            }
        }
        if (Resolve(acroForm.Items.TryGetValue("NeedAppearances", out PdfObject? needAppearances) ? needAppearances : null) is PdfBoolean { Value: true } &&
            Change(PdfAConversionChangeKind.FormUpdated, "interactive form",
                "Removed NeedAppearances so viewers show the stored field appearances, as PDF/A requires; field appearances are stored as viewers generate them.", acroForm)) {
            acroForm.Items.Remove("NeedAppearances");
        }
        if (Resolve(acroForm.Items.TryGetValue("Fields", out PdfObject? fields) ? fields : null) is PdfArray fieldArray) {
            var visited = new HashSet<PdfObject>(ReferenceComparer.Instance);
            foreach (PdfObject field in fieldArray.Items) ProcessField(field, visited, 0);
        }
    }

    private void ProcessField(PdfObject fieldObject, HashSet<PdfObject> visited, int depth) {
        if (depth > 64 || ResolveDictionary(fieldObject) is not PdfDictionary field || !visited.Add(field)) return;
        string name = Resolve(field.Items.TryGetValue("T", out PdfObject? title) ? title : null) is PdfStringObj text ? text.Value : "(unnamed)";
        if (field.Items.ContainsKey("AA") &&
            Change(PdfAConversionChangeKind.ActionRemoved, "form field " + name, "Removed field additional-actions (AA), which PDF/A forbids.", field, OfficeConversionLossKind.Omission)) {
            field.Items.Remove("AA");
        }
        if (Resolve(field.Items.TryGetValue("Kids", out PdfObject? kids) ? kids : null) is PdfArray children) {
            foreach (PdfObject child in children.Items) ProcessField(child, visited, depth + 1);
        }
    }

    private void ProcessOptionalContent(PdfDictionary catalog) {
        if (ResolveDictionary(catalog.Items.TryGetValue("OCProperties", out PdfObject? properties) ? properties : null) is not PdfDictionary ocProperties) return;
        var allGroups = new List<PdfObject>();
        if (Resolve(ocProperties.Items.TryGetValue("OCGs", out PdfObject? groups) ? groups : null) is PdfArray groupArray) allGroups.AddRange(groupArray.Items);
        var configurations = new List<PdfDictionary>();
        if (ResolveDictionary(ocProperties.Items.TryGetValue("D", out PdfObject? defaultConfig) ? defaultConfig : null) is PdfDictionary defaults) configurations.Add(defaults);
        if (Resolve(ocProperties.Items.TryGetValue("Configs", out PdfObject? configs) ? configs : null) is PdfArray configArray) {
            foreach (PdfObject config in configArray.Items) {
                if (ResolveDictionary(config) is PdfDictionary dictionary) configurations.Add(dictionary);
            }
        }
        var usedNames = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 0; index < configurations.Count; index++) {
            PdfDictionary config = configurations[index];
            string location = index == 0 ? "optional content default configuration" : "optional content configuration " + index.ToString(CultureInfo.InvariantCulture);
            string? name = Resolve(config.Items.TryGetValue("Name", out PdfObject? nameValue) ? nameValue : null) is PdfStringObj text ? text.Value : null;
            if (string.IsNullOrEmpty(name) || !usedNames.Add(name!)) {
                string unique = index == 0 ? "Default" : "Configuration " + (index + 1).ToString(CultureInfo.InvariantCulture);
                while (!usedNames.Add(unique)) unique += " ";
                if (Change(PdfAConversionChangeKind.OptionalContentUpdated, location, "Named the configuration '" + unique + "', as PDF/A requires unique names.", config)) {
                    config.Items["Name"] = new PdfStringObj(unique, useTextStringEncoding: true);
                }
            }
            if (config.Items.ContainsKey("AS") &&
                Change(PdfAConversionChangeKind.OptionalContentUpdated, location, "Removed automatic layer-state changes (AS), which PDF/A forbids.", config)) {
                config.Items.Remove("AS");
            }
            if (Resolve(config.Items.TryGetValue("Order", out PdfObject? order) ? order : null) is PdfArray orderArray) {
                var listed = new HashSet<int>();
                CollectOrderReferences(orderArray, listed, 0);
                var missing = allGroups.OfType<PdfReference>().Where(group => !listed.Contains(group.ObjectNumber)).ToList();
                if (missing.Count > 0 &&
                    Change(PdfAConversionChangeKind.OptionalContentUpdated, location,
                        "Listed " + missing.Count.ToString(CultureInfo.InvariantCulture) + " layer(s) missing from the layer order, as PDF/A requires.", config)) {
                    orderArray.Items.AddRange(missing);
                }
            }
        }
    }

    private void CollectOrderReferences(PdfArray order, HashSet<int> listed, int depth) {
        if (depth > 32) return;
        foreach (PdfObject item in order.Items) {
            if (item is PdfReference reference) {
                if (Resolve(reference) is PdfArray nested) CollectOrderReferences(nested, listed, depth + 1);
                else listed.Add(reference.ObjectNumber);
            } else if (item is PdfArray nestedArray) {
                CollectOrderReferences(nestedArray, listed, depth + 1);
            }
        }
    }

    private void ProcessOutlines(PdfDictionary catalog) {
        if (ResolveDictionary(catalog.Items.TryGetValue("Outlines", out PdfObject? outlines) ? outlines : null) is not PdfDictionary root) return;
        var visited = new HashSet<PdfObject>(ReferenceComparer.Instance);
        var pending = new Stack<PdfObject>();
        if (root.Items.TryGetValue("First", out PdfObject? first)) pending.Push(first);
        int processed = 0;
        while (pending.Count > 0 && processed++ < 1_000_000) {
            if (ResolveDictionary(pending.Pop()) is not PdfDictionary item || !visited.Add(item)) continue;
            if (item.Items.TryGetValue("A", out PdfObject? action)) {
                PdfObject? kept = SanitizeAction(action, "bookmark", 0);
                if (kept is null) item.Items.Remove("A");
                else item.Items["A"] = kept;
            }
            if (item.Items.TryGetValue("Next", out PdfObject? next)) pending.Push(next);
            if (item.Items.TryGetValue("First", out PdfObject? child)) pending.Push(child);
        }
    }
}
