namespace OfficeIMO.Pdf;

internal sealed partial class PdfAccessibilityInspector {
    // ISO 32000-1:2008, 14.8.4 standard structure types. H7 and deeper are not standard and must be role mapped
    // (Matterhorn 02-001), but they keep their heading level for the heading checks (ISO 14289-1 7.4.3).
    private static readonly HashSet<string> StandardTypes = new(StringComparer.Ordinal) {
        "Document", "Part", "Art", "Sect", "Div", "BlockQuote", "Caption", "TOC", "TOCI", "Index", "NonStruct", "Private",
        "P", "H", "H1", "H2", "H3", "H4", "H5", "H6", "L", "LI", "Lbl", "LBody", "Table", "TR", "TH", "TD", "THead", "TBody", "TFoot",
        "Span", "Quote", "Note", "Reference", "BibEntry", "Code", "Link", "Annot", "Ruby", "RB", "RT", "RP", "Warichu", "WT", "WP",
        "Figure", "Formula", "Form"
    };

    private sealed class StructNode {
        internal StructNode(int? objectNumber, PdfDictionary dictionary) { ObjectNumber = objectNumber; Dictionary = dictionary; }
        internal int? ObjectNumber { get; }
        internal PdfDictionary Dictionary { get; }
        internal string? Type { get; set; }
        internal string? StandardType { get; set; }
        internal StructNode? Parent { get; set; }
        internal List<StructNode> Kids { get; } = new();
        internal List<PdfObject> ObjectReferences { get; } = new();
        internal int ContentKidCount { get; set; }
        internal int? PageObjectNumber { get; set; }
        internal string? Language { get; set; }
        internal bool HasParentEntry { get; set; }
    }

    private sealed class StructureModel {
        internal PdfDictionary? Root { get; set; }
        internal Dictionary<string, string> RoleMap { get; } = new(StringComparer.Ordinal);
        internal List<StructNode> Nodes { get; } = new();
        internal List<StructNode> TopLevel { get; } = new();
        internal Dictionary<(int? Page, int Mcid), StructNode> PageMcids { get; } = new();
        internal Dictionary<(int Stream, int Mcid), StructNode> StreamMcids { get; } = new();
        internal Dictionary<int, StructNode> AnnotationOwners { get; } = new();
        internal Dictionary<PdfDictionary, StructNode> NodesByDictionary { get; } = new(PdfAccessibilityPageContent.ReferenceEqualityComparer<PdfDictionary>.Instance);
        internal Dictionary<int, PdfObject> ParentTree { get; } = new();
        internal bool Truncated { get; set; }
        /// <summary>Resolved role mappings by structure type, so long mapping chains are walked once per type.</summary>
        internal Dictionary<string, (string? Standard, bool Circular, bool Unmapped)> MappedRoles { get; } = new(StringComparer.Ordinal);

        internal bool IsLinked(int? pageObject, int? stream, int mcid) => stream is int streamObject
            ? StreamMcids.ContainsKey((streamObject, mcid))
            : PageMcids.ContainsKey((pageObject, mcid)) || PageMcids.ContainsKey((null, mcid)) && pageObject is null;

        internal StructNode? Owner(int? pageObject, int? stream, int mcid) => stream is int streamObject
            ? StreamMcids.TryGetValue((streamObject, mcid), out StructNode? node) ? node : null
            : PageMcids.TryGetValue((pageObject, mcid), out StructNode? owner) ? owner : null;
    }

    private StructureModel ReadStructure() {
        var model = new StructureModel();
        model.Root = _catalog is null ? null : ResolveDict(_catalog.Items.TryGetValue("StructTreeRoot", out PdfObject? rootObject) ? rootObject : null);
        if (model.Root is null) return model;
        if (ResolveDict(model.Root.Items.TryGetValue("RoleMap", out PdfObject? roleMapObject) ? roleMapObject : null) is PdfDictionary roleMap) {
            foreach (KeyValuePair<string, PdfObject> entry in roleMap.Items) {
                if (Resolve(entry.Value) is PdfName mapped) model.RoleMap[entry.Key] = mapped.Name;
            }
        }
        var visited = new HashSet<PdfDictionary>(PdfAccessibilityPageContent.ReferenceEqualityComparer<PdfDictionary>.Instance);
        int maxNodes = Math.Max(1, _document.ReadOptions.Limits.MaxIndirectObjects);
        ReadKids(model.Root.Items.TryGetValue("K", out PdfObject? rootKids) ? rootKids : null, null, null, 0);
        ReadParentTree(model);
        return model;

        void ReadKids(PdfObject? kids, StructNode? parent, int? inheritedPage, int depth) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (kids is null) return;
            if (depth > MaxStructureDepth) { model.Truncated = true; return; }
            PdfObject? resolved = Resolve(kids);
            if (resolved is PdfArray array) {
                foreach (PdfObject item in array.Items) ReadKid(item, parent, inheritedPage, depth);
                return;
            }
            ReadKid(kids, parent, inheritedPage, depth);
        }

        void ReadKid(PdfObject item, StructNode? parent, int? inheritedPage, int depth) {
            PdfObject? value = Resolve(item);
            if (value is PdfNumber number) {
                if (parent is null) return;
                parent.ContentKidCount++;
                if (number.Value >= 0 && number.Value <= int.MaxValue) model.PageMcids[(inheritedPage, (int)number.Value)] = parent;
                return;
            }
            if (value is not PdfDictionary dictionary) return;
            string? type = ReadName(dictionary, "Type");
            if (string.Equals(type, "MCR", StringComparison.Ordinal) || (type is null && dictionary.Items.ContainsKey("MCID") && !dictionary.Items.ContainsKey("S"))) {
                if (parent is null) return;
                parent.ContentKidCount++;
                int? mcid = ReadInteger(dictionary, "MCID");
                if (mcid is not int id || id < 0) return;
                int? stream = ObjectNumberOf(dictionary.Items.TryGetValue("Stm", out PdfObject? stm) ? stm : null);
                int? page = ObjectNumberOf(dictionary.Items.TryGetValue("Pg", out PdfObject? pg) ? pg : null) ?? inheritedPage;
                if (stream is int streamObject) model.StreamMcids[(streamObject, id)] = parent;
                else model.PageMcids[(page, id)] = parent;
                return;
            }
            if (string.Equals(type, "OBJR", StringComparison.Ordinal)) {
                if (parent is null) return;
                parent.ContentKidCount++;
                if (dictionary.Items.TryGetValue("Obj", out PdfObject? target)) {
                    parent.ObjectReferences.Add(target);
                    if (target is PdfReference targetReference) model.AnnotationOwners[targetReference.ObjectNumber] = parent;
                }
                return;
            }
            if (!dictionary.Items.ContainsKey("S")) return;
            if (!visited.Add(dictionary)) return;
            if (model.Nodes.Count >= maxNodes) { model.Truncated = true; return; }
            var node = new StructNode(ObjectNumberOf(item), dictionary) {
                Type = ReadName(dictionary, "S"),
                Parent = parent,
                PageObjectNumber = ObjectNumberOf(dictionary.Items.TryGetValue("Pg", out PdfObject? pageObject) ? pageObject : null) ?? inheritedPage,
                Language = ReadText(dictionary, "Lang"),
                HasParentEntry = dictionary.Items.ContainsKey("P")
            };
            // A role mapping wins; an unmapped H7 or deeper still counts as a heading of its own level.
            node.StandardType = MapToStandard(node.Type, model, out _, out _) ??
                (node.Type is not null && TryGetHeadingLevel(node.Type, out _) ? node.Type : null);
            model.Nodes.Add(node);
            model.NodesByDictionary[dictionary] = node;
            if (parent is null) model.TopLevel.Add(node); else parent.Kids.Add(node);
            ReadKids(dictionary.Items.TryGetValue("K", out PdfObject? kids) ? kids : null, node, node.PageObjectNumber, depth + 1);
        }
    }

    private void ReadParentTree(StructureModel model) {
        PdfObject? tree = model.Root!.Items.TryGetValue("ParentTree", out PdfObject? value) ? value : null;
        var visited = new HashSet<PdfObject>();
        Visit(tree, 0);

        void Visit(PdfObject? node, int depth) {
            if (ResolveDict(node) is not PdfDictionary dictionary || !visited.Add(dictionary) || depth > MaxStructureDepth) return;
            if (ResolveArray(dictionary.Items.TryGetValue("Nums", out PdfObject? nums) ? nums : null) is PdfArray pairs) {
                for (int i = 1; i < pairs.Items.Count; i += 2) {
                    if (Resolve(pairs.Items[i - 1]) is PdfNumber key) model.ParentTree[(int)key.Value] = pairs.Items[i];
                }
            }
            if (ResolveArray(dictionary.Items.TryGetValue("Kids", out PdfObject? kids) ? kids : null) is PdfArray kidArray) {
                foreach (PdfObject kid in kidArray.Items) Visit(kid, depth + 1);
            }
        }
    }

    // Conditions whose evidence comes from walking the structure tree.
    private static readonly string[] StructureConditions = {
        "UA1:7.1-12", "02-001", "02-003", "02-004", "09-004", "09-005", "09-006", "09-007", "09-008", "11-002", "13-004",
        "14-002", "14-003", "14-006", "14-007", "15-003", "17-002", "19-003", "19-004", "UA1:7.18.4-2"
    };

    private static bool IsStandard(string type) => StandardTypes.Contains(type);

    private static string? MapToStandard(string? type, StructureModel model, out bool circular, out bool unmapped) {
        circular = false;
        unmapped = false;
        if (type is null) return null;
        if (IsStandard(type)) return type;
        if (!model.MappedRoles.TryGetValue(type, out var mapped)) {
            mapped = WalkRoleMap(type, model);
            model.MappedRoles[type] = mapped;
        }
        circular = mapped.Circular;
        unmapped = mapped.Unmapped;
        return mapped.Standard;
    }

    private static (string? Standard, bool Circular, bool Unmapped) WalkRoleMap(string type, StructureModel model) {
        var seen = new HashSet<string>(StringComparer.Ordinal) { type };
        string current = type;
        while (model.RoleMap.TryGetValue(current, out string? next)) {
            if (IsStandard(next)) return (next, false, false);
            if (!seen.Add(next)) return (null, true, false);
            current = next;
        }
        return (null, false, true);
    }

    private void CheckStructure(StructureModel model) {
        Applies("UA1:7.1-11");
        if (model.Root is null) {
            Fail("UA1:7.1-11", "7.1-11", "The catalog has no StructTreeRoot. Tag the document so assistive technology can read its logical structure.");
            return;
        }
        if (model.Truncated) {
            foreach (string id in StructureConditions) {
                Unknown(id, null, "The structure tree exceeds the inspection limits; deeper elements were not inspected.");
            }
        }

        bool usesH = false, usesHn = false;
        int previousLevel = 0;
        var noteIds = new Dictionary<string, StructNode>(StringComparer.Ordinal);
        foreach (StructNode node in model.Nodes) {
            _cancellationToken.ThrowIfCancellationRequested();
            Applies("UA1:7.1-12");
            if (!node.HasParentEntry) Fail("UA1:7.1-12", "7.1-12", "The structure element has no P (parent) entry.", element: node);
            if (node.Language is not null) CheckLanguageValue(node.Language, null, node, "a structure element", node.ObjectNumber);
            CheckRoleMapping(node, model);
            string? type = node.StandardType;
            if (type is null) continue;
            CheckGrouping(node, type);
            CheckAlternateTextLanguage(node);

            if (type == "Figure") {
                Applies("13-004");
                if (string.IsNullOrEmpty(ReadText(node.Dictionary, "Alt")) && ReadText(node.Dictionary, "ActualText") is null) {
                    Fail("13-004", "7.3-1", "The Figure has no Alt or ActualText. Describe the image, or mark it as an artifact if it is decorative.", element: node);
                }
            } else if (type == "Formula") {
                Applies("17-002");
                if (string.IsNullOrEmpty(ReadText(node.Dictionary, "Alt")) && ReadText(node.Dictionary, "ActualText") is null) {
                    Fail("17-002", "7.7-1", "The Formula has no Alt or ActualText describing the expression.", element: node);
                }
            } else if (type == "Note") {
                Applies("19-003");
                string? id = node.Dictionary.Items.TryGetValue("ID", out PdfObject? idObject) && Resolve(idObject) is PdfStringObj idText ? idText.Value : null;
                if (string.IsNullOrEmpty(id)) {
                    Fail("19-003", "7.9-1", "The Note has no ID entry.", element: node);
                } else {
                    Applies("19-004");
                    if (noteIds.ContainsKey(id!)) Fail("19-004", "7.9-2", "The Note ID '" + id + "' is already used by another Note.", element: node);
                    else noteIds[id!] = node;
                }
            } else if (type == "H") {
                usesH = true;
                Applies("14-006");
            } else if (TryGetHeadingLevel(type, out int level)) {
                usesHn = true;
                Applies("14-002");
                Applies("14-003");
                if (previousLevel == 0 && level != 1) {
                    Fail("14-002", "7.4.2-1", "The first numbered heading is " + type + "; it must be H1.", element: node);
                } else if (previousLevel != 0 && level > previousLevel + 1) {
                    Fail("14-003", "7.4.2-1", type + " follows H" + previousLevel.ToString(System.Globalization.CultureInfo.InvariantCulture) + ", skipping a level.", element: node);
                }
                previousLevel = level;
            }
            if (node.Kids.Count(static kid => kid.StandardType == "H") > 1) {
                Fail("14-006", "7.4.4-1", "The element has more than one H child; give each H its own section.", element: node);
            }
            if (type == "Table") CheckTableHeaders(node);
            if (type == "Form") CheckFormElement(node);
        }
        if (usesH || usesHn) Applies("14-007");
        if (usesH && usesHn) {
            Fail("14-007", "7.4.4-2", "The document uses H headings together with numbered headings; use one scheme.");
            Fail("14-007", "7.4.4-3", "The document uses numbered headings together with H headings; use one scheme.");
        }
    }

    private void CheckRoleMapping(StructNode node, StructureModel model) {
        if (node.Type is null) return;
        if (IsStandard(node.Type)) {
            Applies("02-004");
            if (model.RoleMap.TryGetValue(node.Type, out string? remapped)) {
                if (string.Equals(remapped, node.Type, StringComparison.Ordinal)) {
                    Applies("02-003");
                    Fail("02-003", "7.1-6", "The RoleMap maps " + node.Type + " to itself.", element: node);
                } else {
                    Fail("02-004", "7.1-7", "The standard type " + node.Type + " is remapped to " + remapped + " in the RoleMap.", element: node);
                }
            }
            return;
        }
        Applies("02-001");
        Applies("02-003");
        MapToStandard(node.Type, model, out bool circular, out bool unmapped);
        if (circular) {
            Fail("02-003", "7.1-6", "The role mapping for " + node.Type + " is circular.", element: node);
            Fail("02-001", "7.1-5", "The role mapping for " + node.Type + " never reaches a standard type.", element: node);
        } else if (unmapped) Fail("02-001", "7.1-5", "The non-standard type " + node.Type + " is not role mapped to a standard type.", element: node);
    }

    private static bool TryGetHeadingLevel(string type, out int level) {
        level = 0;
        if (type.Length < 2 || type.Length > 6 || type[0] != 'H') return false;
        for (int i = 1; i < type.Length; i++) {
            if (type[i] < '0' || type[i] > '9') return false;
            level = level * 10 + (type[i] - '0');
        }
        return level > 0;
    }

    private void CheckGrouping(StructNode node, string type) {
        string? parent = node.Parent?.StandardType;
        string[] kids = node.Kids.Select(static kid => kid.StandardType ?? string.Empty).ToArray();
        // ISO 32000-1 Tables 336 to 338 allow these grouping elements only element children, never content. LI is left
        // out: producers commonly put item content in LI directly, and veraPDF accepts it.
        string? groupingCondition = type switch {
            "Table" or "TR" or "THead" or "TBody" or "TFoot" => "09-004",
            "L" => "09-005",
            "TOC" => "09-006",
            "Ruby" => "09-007",
            "Warichu" => "09-008",
            _ => null
        };
        if (groupingCondition is not null && node.ContentKidCount > 0) {
            Applies(groupingCondition);
            Fail(groupingCondition, null, type + " contains content or an object reference directly; it may contain only its permitted child elements.", element: node);
        }
        switch (type) {
            case "Table": {
                Applies("09-004");
                if (kids.Any(static kid => kid is not ("TR" or "THead" or "TBody" or "TFoot" or "Caption"))) TableFail("7.2-3", "A Table may contain only TR, THead, TBody, TFoot and Caption.");
                int heads = kids.Count(static kid => kid == "THead"), feet = kids.Count(static kid => kid == "TFoot"), bodies = kids.Count(static kid => kid == "TBody");
                if (heads > 1) TableFail("7.2-11", "A Table may contain at most one THead.");
                if (feet > 1) TableFail("7.2-12", "A Table may contain at most one TFoot.");
                if (feet > 0 && bodies == 0) TableFail("7.2-13", "A Table with a TFoot needs at least one TBody.");
                if (heads > 0 && bodies == 0) TableFail("7.2-14", "A Table with a THead needs at least one TBody.");
                int captions = kids.Count(static kid => kid == "Caption");
                if (captions > 1) TableFail("7.2-39", "A Table may contain only one Caption.");
                for (int i = 1; i < kids.Length - 1; i++) {
                    if (kids[i] == "Caption") { TableFail("7.2-16", "A Table Caption must be its first or last child."); break; }
                }
                break;
            }
            case "TR":
                Applies("09-004");
                if (parent is not ("Table" or "THead" or "TBody" or "TFoot")) TableFail("7.2-4", "TR must be inside Table, THead, TBody or TFoot.");
                if (kids.Any(static kid => kid is not ("TH" or "TD"))) TableFail("7.2-10", "TR may contain only TH and TD.");
                break;
            case "THead":
            case "TBody":
            case "TFoot":
                Applies("09-004");
                if (parent != "Table") TableFail(type == "THead" ? "7.2-5" : type == "TBody" ? "7.2-6" : "7.2-7", type + " must be inside a Table.");
                if (kids.Any(static kid => kid != "TR")) TableFail(type == "THead" ? "7.2-36" : type == "TBody" ? "7.2-37" : "7.2-38", type + " may contain only TR.");
                break;
            case "TH":
            case "TD":
                Applies("09-004");
                if (parent != "TR") TableFail(type == "TH" ? "7.2-8" : "7.2-9", type + " must be inside a TR.");
                break;
            case "L":
                Applies("09-005");
                if (kids.Any(static kid => kid is not ("L" or "LI" or "Caption"))) ListFail("7.2-19", "L may contain only L, LI and Caption.");
                for (int i = 1; i < kids.Length; i++) {
                    if (kids[i] == "Caption") { ListFail("7.2-40", "A list Caption must be the first child of L."); break; }
                }
                break;
            case "LI":
                Applies("09-005");
                if (parent != "L") ListFail("7.2-17", "LI must be inside an L.");
                if (kids.Any(static kid => kid is not ("Lbl" or "LBody"))) ListFail("7.2-20", "LI may contain only Lbl and LBody.");
                break;
            case "LBody":
                Applies("09-005");
                if (parent != "LI") ListFail("7.2-18", "LBody must be inside an LI.");
                break;
            case "TOC":
                Applies("09-006");
                if (kids.Any(static kid => kid is not ("TOC" or "TOCI" or "Caption"))) TocFail("7.2-27", "TOC may contain only TOC, TOCI and Caption.");
                for (int i = 1; i < kids.Length; i++) {
                    if (kids[i] == "Caption") { TocFail("7.2-28", "A TOC Caption must be its first child."); break; }
                }
                break;
            case "TOCI":
                Applies("09-006");
                if (parent != "TOC") TocFail("7.2-26", "TOCI must be inside a TOC.");
                break;
            case "Ruby":
                Applies("09-007");
                if (kids.Any(static kid => kid is not ("RB" or "RT" or "RP"))) Fail("09-007", null, "Ruby may contain only RB, RT and RP.", element: node);
                break;
            case "RB":
            case "RT":
            case "RP":
                Applies("09-007");
                if (parent != "Ruby") Fail("09-007", null, type + " must be inside a Ruby.", element: node);
                break;
            case "Warichu":
                Applies("09-008");
                if (kids.Any(static kid => kid is not ("WT" or "WP"))) Fail("09-008", null, "Warichu may contain only WT and WP.", element: node);
                break;
            case "WT":
            case "WP":
                Applies("09-008");
                if (parent != "Warichu") Fail("09-008", null, type + " must be inside a Warichu.", element: node);
                break;
        }

        void TableFail(string rule, string message) => Fail("09-004", rule, message, element: node);
        void ListFail(string rule, string message) => Fail("09-005", rule, message, element: node);
        void TocFail(string rule, string message) => Fail("09-006", rule, message, element: node);
    }

    private void CheckTableHeaders(StructNode table) {
        var cells = new List<StructNode>();
        var pending = new Stack<StructNode>(table.Kids);
        while (pending.Count > 0) {
            StructNode current = pending.Pop();
            if (current.StandardType is "TH" or "TD") cells.Add(current);
            else if (current.StandardType is "TR" or "THead" or "TBody" or "TFoot") foreach (StructNode kid in current.Kids) pending.Push(kid);
        }
        if (!cells.Any(static cell => cell.StandardType == "TH")) return;
        Applies("15-003");
        // A table is organised with Headers and IDs only when every data cell names headers that are TH cells of this table.
        var headerIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (StructNode cell in cells) {
            if (cell.StandardType == "TH" && cell.Dictionary.Items.TryGetValue("ID", out PdfObject? idObject) &&
                Resolve(idObject) is PdfStringObj id && !string.IsNullOrEmpty(id.Value)) {
                headerIds.Add(id.Value);
            }
        }
        StructNode[] dataCells = cells.Where(static cell => cell.StandardType == "TD").ToArray();
        bool usesHeaders = dataCells.Length > 0 && dataCells.All(cell =>
            ResolveArray(ReadTableAttribute(cell, "Headers")) is PdfArray headers && headers.Items.Count > 0 &&
            headers.Items.All(header => Resolve(header) is PdfStringObj reference && headerIds.Contains(reference.Value)));
        if (usesHeaders) return;
        foreach (StructNode cell in cells) {
            if (cell.StandardType != "TH") continue;
            string? scope = Resolve(ReadTableAttribute(cell, "Scope")) is PdfName scopeName ? scopeName.Name : null;
            if (scope is not ("Row" or "Column" or "Both")) {
                Fail("15-003", scope is null ? "7.5-1" : "7.5-2", scope is null
                    ? "The TH cell has no Scope attribute and the table does not identify headers with Headers and IDs."
                    : "The TH cell's Scope '" + scope + "' is not Row, Column or Both.", element: cell);
            }
        }
    }

    private PdfObject? ReadTableAttribute(StructNode cell, string key) {
        foreach (PdfDictionary attributes in ReadAttributeDictionaries(cell.Dictionary)) {
            if (string.Equals(ReadName(attributes, "O"), "Table", StringComparison.Ordinal) && attributes.Items.TryGetValue(key, out PdfObject? value)) {
                return value;
            }
        }
        return null;
    }

    private IEnumerable<PdfDictionary> ReadAttributeDictionaries(PdfDictionary element) {
        if (element.Items.TryGetValue("A", out PdfObject? attributes)) {
            PdfObject? resolved = Resolve(attributes);
            if (resolved is PdfArray array) {
                foreach (PdfObject item in array.Items) if (ResolveDict(item) is PdfDictionary dictionary) yield return dictionary;
            } else if (ResolveDict(resolved) is PdfDictionary single) {
                yield return single;
            }
        }
        if (!element.Items.TryGetValue("C", out PdfObject? classes) || _catalog is null) yield break;
        PdfDictionary? root = ResolveDict(_catalog.Items.TryGetValue("StructTreeRoot", out PdfObject? rootObject) ? rootObject : null);
        PdfDictionary? classMap = root is null ? null : ResolveDict(root.Items.TryGetValue("ClassMap", out PdfObject? classMapObject) ? classMapObject : null);
        if (classMap is null) yield break;
        PdfObject? resolvedClasses = Resolve(classes);
        IEnumerable<PdfObject> names = resolvedClasses is PdfArray classArray ? classArray.Items : new[] { resolvedClasses ?? PdfNull.Instance };
        foreach (PdfObject name in names) {
            if (Resolve(name) is not PdfName className || !classMap.Items.TryGetValue(className.Name, out PdfObject? mapped)) continue;
            PdfObject? mappedValue = Resolve(mapped);
            if (mappedValue is PdfArray mappedArray) {
                foreach (PdfObject item in mappedArray.Items) if (ResolveDict(item) is PdfDictionary dictionary) yield return dictionary;
            } else if (ResolveDict(mappedValue) is PdfDictionary dictionary) {
                yield return dictionary;
            }
        }
    }

    private void CheckFormElement(StructNode form) {
        Applies("UA1:7.18.4-2");
        bool hasRole = ReadAttributeDictionaries(form.Dictionary).Any(attributes =>
            string.Equals(ReadName(attributes, "O"), "PrintField", StringComparison.Ordinal) && attributes.Items.ContainsKey("Role"));
        if (hasRole) return;
        if (form.Kids.Count != 0 || form.ContentKidCount != 1 || form.ObjectReferences.Count != 1 ||
            ResolveDict(form.ObjectReferences[0]) is not PdfDictionary target || ReadName(target, "Subtype") != "Widget") {
            Fail("UA1:7.18.4-2", "7.18.4-2", "A Form element without a Role attribute must contain exactly one object reference to its widget.", element: form);
        }
    }

    private void CheckAlternateTextLanguage(StructNode node) {
        foreach ((string key, string rule) in new[] { ("ActualText", "7.2-21"), ("Alt", "7.2-22"), ("E", "7.2-23") }) {
            if (ReadText(node.Dictionary, key) is null) continue;
            Applies("11-002");
            if (!HasLanguage(node)) {
                Fail("11-002", rule, "The " + key + " text has no determinable language: set Lang on the element, an ancestor, or the catalog.", element: node);
            }
        }
    }

    private bool HasLanguage(StructNode? node) {
        if (_catalogHasLanguage) return true;
        for (StructNode? current = node; current is not null; current = current.Parent) {
            if (!string.IsNullOrEmpty(current.Language)) return true;
        }
        return false;
    }
}
