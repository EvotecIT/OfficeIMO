namespace OfficeIMO.Pdf;

internal sealed class PdfPageOptionalContentVisibility {
    private readonly Dictionary<string, bool> _hiddenProperties;
    private readonly HashSet<string> _knownProperties;
    private readonly HashSet<string> _invalidProperties;
    private readonly HashSet<int> _hiddenObjectNumbers;
    private readonly Dictionary<int, bool> _groupVisibility;
    private readonly Dictionary<int, PdfIndirectObject> _objects;
    private readonly int _maxExpressionDepth;
    private readonly System.Collections.Concurrent.ConcurrentDictionary<string, (bool Success, bool Visible)> _inlineVisibilityExpressionCache =
        new System.Collections.Concurrent.ConcurrentDictionary<string, (bool Success, bool Visible)>(StringComparer.Ordinal);

    private PdfPageOptionalContentVisibility(Dictionary<string, bool> hiddenProperties, HashSet<string> knownProperties, HashSet<string> invalidProperties, HashSet<int> hiddenObjectNumbers, Dictionary<int, bool> groupVisibility, Dictionary<int, PdfIndirectObject> objects, int maxExpressionDepth, bool hasUnsupportedViewUsageApplications) {
        _hiddenProperties = hiddenProperties;
        _knownProperties = knownProperties;
        _invalidProperties = invalidProperties;
        _hiddenObjectNumbers = hiddenObjectNumbers;
        _groupVisibility = groupVisibility;
        _objects = objects;
        _maxExpressionDepth = maxExpressionDepth;
        HasUnsupportedViewUsageApplications = hasUnsupportedViewUsageApplications;
    }

    internal bool HasUnsupportedViewUsageApplications { get; }

    public static PdfPageOptionalContentVisibility? Create(
        PdfDictionary? resources,
        Dictionary<int, PdfIndirectObject> objects,
        int maxExpressionDepth) {
        int effectiveMaxExpressionDepth = System.Math.Min(maxExpressionDepth, PdfReadLimits.DefaultMaxContentNestingDepth);
        Dictionary<int, bool> groupVisibility = ReadGroupVisibility(objects, out bool hasUnsupportedViewUsageApplications);
        bool hasPropertiesDeclaration = resources != null && resources.Items.ContainsKey("Properties");
        PdfDictionary? properties = resources != null &&
            resources.Items.TryGetValue("Properties", out PdfObject? propertiesObject)
                ? ResolveObject(propertiesObject, objects) as PdfDictionary
                : null;
        if (groupVisibility.Count == 0 && !hasUnsupportedViewUsageApplications && !hasPropertiesDeclaration) {
            return null;
        }

        var hiddenObjectNumbers = new HashSet<int>();
        foreach (KeyValuePair<int, bool> entry in groupVisibility) {
            if (!entry.Value) {
                hiddenObjectNumbers.Add(entry.Key);
            }
        }

        foreach (KeyValuePair<int, PdfIndirectObject> entry in objects) {
            if (hiddenObjectNumbers.Contains(entry.Key)) {
                continue;
            }

            if (IsOptionalContentObjectHidden(entry.Value.Value, groupVisibility, objects, new HashSet<int>(), effectiveMaxExpressionDepth, depth: 0)) {
                hiddenObjectNumbers.Add(entry.Key);
            }
        }

        var hiddenProperties = new Dictionary<string, bool>(StringComparer.Ordinal);
        var knownProperties = new HashSet<string>(StringComparer.Ordinal);
        var invalidProperties = new HashSet<string>(StringComparer.Ordinal);
        if (properties != null) {
            foreach (KeyValuePair<string, PdfObject> entry in properties.Items) {
                knownProperties.Add(entry.Key);
                if (IsOptionalContentObjectInvalid(entry.Value, groupVisibility, objects, new HashSet<int>(), effectiveMaxExpressionDepth, depth: 0)) {
                    invalidProperties.Add(entry.Key);
                }
                if (IsOptionalContentObjectHidden(entry.Value, groupVisibility, objects, new HashSet<int>(), effectiveMaxExpressionDepth, depth: 0)) {
                    hiddenProperties[entry.Key] = true;
                }
            }
        }

        return new PdfPageOptionalContentVisibility(hiddenProperties, knownProperties, invalidProperties, hiddenObjectNumbers, groupVisibility, objects, effectiveMaxExpressionDepth, hasUnsupportedViewUsageApplications);
    }

    public bool IsHidden(string propertyName) =>
        _hiddenProperties.TryGetValue(propertyName, out bool hidden) && hidden;

    internal bool HasInvalidProperty(string propertyName) =>
        !_knownProperties.Contains(propertyName) || _invalidProperties.Contains(propertyName);

    public bool IsHiddenAny(IReadOnlyList<int> objectNumbers) {
        for (int i = 0; i < objectNumbers.Count; i++) {
            if (_hiddenObjectNumbers.Contains(objectNumbers[i])) {
                return true;
            }
        }

        return false;
    }

    public bool IsHidden(PdfInlineOptionalContentReferences references) {
        if (references.IsMembershipDictionary) {
            if (!string.IsNullOrWhiteSpace(references.VisibilityExpression)) {
                string expression = references.VisibilityExpression!;
                if (TryEvaluateInlineOrIndirectVisibilityExpression(expression, out bool expressionVisible)) return !expressionVisible;
            }

            return IsMembershipHidden(references.ObjectReferences, references.Policy);
        }

        return IsHiddenAny(references.ObjectNumbers);
    }

    internal bool IsHidden(PdfObject optionalContentObject) =>
        IsOptionalContentObjectHidden(
            optionalContentObject,
            _groupVisibility,
            _objects,
            new HashSet<int>(),
            _maxExpressionDepth,
            depth: 0);

    internal bool HasInvalidMembershipReferences(PdfInlineOptionalContentReferences references) {
        if (!references.IsMembershipDictionary) return false;
        if (references.HasInvalidPolicy || references.HasInvalidGroupContainer) return true;
        if (!string.IsNullOrWhiteSpace(references.VisibilityExpression) &&
            !TryEvaluateInlineOrIndirectVisibilityExpression(references.VisibilityExpression!, out _)) return true;
        for (int index = 0; index < references.ObjectReferences.Count; index++) {
            PdfReference reference = references.ObjectReferences[index];
            if (!PdfObjectLookup.TryGet(_objects, reference, out PdfIndirectObject groupObject) ||
                !_groupVisibility.ContainsKey(reference.ObjectNumber) ||
                ResolveObject(groupObject.Value, _objects) is not PdfDictionary group ||
                ResolveObject(group.Items.TryGetValue("Type", out PdfObject? typeObject) ? typeObject : null, _objects) is not PdfName { Name: "OCG" }) {
                return true;
            }
        }
        return false;
    }

    private bool TryEvaluateInlineVisibilityExpression(string expression, out bool visible) {
        visible = false;
        int index = 0;
        SkipInlineWhitespace(expression, ref index);
        if (index >= expression.Length || expression[index] != '[') {
            return false;
        }

        if (!TryEvaluateInlineVisibilityExpression(expression, ref index, depth: 0, out visible)) {
            return false;
        }

        SkipInlineWhitespace(expression, ref index);
        return index == expression.Length;
    }

    private bool TryEvaluateInlineOrIndirectVisibilityExpression(string expression, out bool visible) {
        if (_inlineVisibilityExpressionCache.TryGetValue(expression, out (bool Success, bool Visible) cached)) {
            visible = cached.Visible;
            return cached.Success;
        }

        bool success = TryEvaluateInlineOrIndirectVisibilityExpressionUncached(expression, out visible);
        _inlineVisibilityExpressionCache[expression] = (success, visible);
        return success;
    }

    private bool TryEvaluateInlineOrIndirectVisibilityExpressionUncached(string expression, out bool visible) {
        if (TryEvaluateInlineVisibilityExpression(expression, out visible)) return true;

        visible = false;
        int index = 0;
        SkipInlineWhitespace(expression, ref index);
        if (!TryReadInlineReference(expression, ref index, out PdfReference reference)) return false;
        SkipInlineWhitespace(expression, ref index);
        return index == expression.Length &&
            PdfObjectLookup.TryGet(_objects, reference, out PdfIndirectObject indirect) &&
            TryEvaluateVisibilityExpression(
                indirect.Value,
                _groupVisibility,
                _objects,
                new HashSet<int> { reference.ObjectNumber },
                _maxExpressionDepth,
                depth: 0,
                out visible);
    }

    private bool TryEvaluateInlineVisibilityExpression(string expression, ref int index, int depth, out bool visible) {
        visible = false;
        if (depth > _maxExpressionDepth) {
            return false;
        }
        SkipInlineWhitespace(expression, ref index);
        if (index >= expression.Length) {
            return false;
        }

        if (expression[index] == '[') {
            return TryEvaluateInlineVisibilityArray(expression, ref index, depth, out visible);
        }

        if (TryReadInlineReference(expression, ref index, out PdfReference reference)) {
            if (!PdfObjectLookup.TryGet(_objects, reference, out PdfIndirectObject indirect)) {
                return false;
            }
            if (_groupVisibility.TryGetValue(reference.ObjectNumber, out visible)) {
                return true;
            }

            return TryEvaluateVisibilityExpression(
                    indirect.Value,
                    _groupVisibility,
                    _objects,
                    new HashSet<int> { reference.ObjectNumber },
                    _maxExpressionDepth,
                    depth + 1,
                    out visible);
        }

        return false;
    }

    private bool TryEvaluateInlineVisibilityArray(string expression, ref int index, int depth, out bool visible) {
        visible = false;
        if (index >= expression.Length || expression[index] != '[') {
            return false;
        }

        index++;
        SkipInlineWhitespace(expression, ref index);
        if (!TryReadInlineName(expression, ref index, out string? operatorName)) {
            return false;
        }

        switch (operatorName) {
            case "And":
                visible = true;
                int andOperandCount = 0;
                while (TryReadInlineExpressionOperand(expression, ref index, depth + 1, out bool operandVisible)) {
                    visible &= operandVisible;
                    andOperandCount++;
                }

                return andOperandCount > 0 && TryCloseInlineArray(expression, ref index);
            case "Or":
                visible = false;
                int orOperandCount = 0;
                while (TryReadInlineExpressionOperand(expression, ref index, depth + 1, out bool operandVisible)) {
                    visible |= operandVisible;
                    orOperandCount++;
                }

                return orOperandCount > 0 && TryCloseInlineArray(expression, ref index);
            case "Not":
                if (!TryReadInlineExpressionOperand(expression, ref index, depth + 1, out bool nestedVisible)) {
                    return false;
                }

                visible = !nestedVisible;
                return TryCloseInlineArray(expression, ref index);
            default:
                return false;
        }
    }

    private bool TryReadInlineExpressionOperand(string expression, ref int index, int depth, out bool visible) {
        visible = false;
        SkipInlineWhitespace(expression, ref index);
        if (index >= expression.Length || expression[index] == ']') {
            return false;
        }

        return TryEvaluateInlineVisibilityExpression(expression, ref index, depth, out visible);
    }

    private static bool TryCloseInlineArray(string expression, ref int index) {
        SkipInlineWhitespace(expression, ref index);
        if (index >= expression.Length || expression[index] != ']') {
            return false;
        }

        index++;
        return true;
    }

    private static void SkipInlineWhitespace(string text, ref int index) {
        while (index < text.Length) {
            while (index < text.Length && IsInlineWhitespace(text[index])) {
                index++;
            }

            if (index >= text.Length || text[index] != '%') {
                return;
            }

            while (index < text.Length && text[index] != '\r' && text[index] != '\n') {
                index++;
            }
        }
    }

    private static bool IsInlineWhitespace(char ch) =>
        ch == '\0' || ch == '\t' || ch == '\n' || ch == '\f' || ch == '\r' || ch == ' ';

    private static bool TryReadInlineName(string text, ref int index, out string? name) {
        name = null;
        SkipInlineWhitespace(text, ref index);
        if (index >= text.Length || text[index] != '/') {
            return false;
        }

        index++;
        int start = index;
        while (index < text.Length &&
               !IsInlineWhitespace(text[index]) &&
               text[index] != '[' && text[index] != ']' && text[index] != '%' && text[index] != '/' &&
               text[index] != '(' && text[index] != ')' && text[index] != '<' && text[index] != '>') {
            index++;
        }

        if (index == start) {
            return false;
        }

        name = PdfSyntax.DecodeName(text.Substring(start, index - start));
        return true;
    }

    private static bool TryReadInlineReference(string text, ref int index, out PdfReference reference) {
        reference = null!;
        SkipInlineWhitespace(text, ref index);
        int start = index;
        if (!TryReadInlineInteger(text, ref index, out int objectNumber)) {
            return false;
        }

        SkipInlineWhitespace(text, ref index);
        if (!TryReadInlineInteger(text, ref index, out int generation)) {
            index = start;
            return false;
        }

        SkipInlineWhitespace(text, ref index);
        if (index >= text.Length || text[index] != 'R' || !IsInlineTokenBoundary(text, index + 1)) {
            index = start;
            return false;
        }

        index++;
        reference = new PdfReference(objectNumber, generation);
        return true;
    }

    private static bool IsInlineTokenBoundary(string text, int index) {
        if (index >= text.Length) {
            return true;
        }

        char ch = text[index];
        return IsInlineWhitespace(ch) || ch == '%' || ch == '(' || ch == ')' || ch == '<' || ch == '>' ||
            ch == '[' || ch == ']' || ch == '{' || ch == '}' || ch == '/';
    }

    private static bool TryReadInlineInteger(string text, ref int index, out int value) {
        value = 0;
        SkipInlineWhitespace(text, ref index);
        int start = index;
        if (index < text.Length && (text[index] == '+' || text[index] == '-')) {
            index++;
        }

        int digitStart = index;
        while (index < text.Length && char.IsDigit(text[index])) {
            index++;
        }

        if (index == digitStart ||
#pragma warning disable CA1846
            !int.TryParse(text.Substring(start, index - start), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out value)) {
#pragma warning restore CA1846
            index = start;
            return false;
        }

        return true;
    }

    private bool IsMembershipHidden(IReadOnlyList<PdfReference> objectReferences, string? policy) {
        bool anyVisible = false;
        bool anyHidden = false;
        bool hasResolvedGroup = false;
        for (int i = 0; i < objectReferences.Count; i++) {
            PdfReference reference = objectReferences[i];
            if (!PdfObjectLookup.TryGet(_objects, reference, out _)) continue;
            hasResolvedGroup = true;
            bool visible = !_hiddenObjectNumbers.Contains(reference.ObjectNumber);
            anyVisible |= visible;
            anyHidden |= !visible;
        }
        if (!hasResolvedGroup && objectReferences.Count > 0) return false;

        bool visibleByPolicy = policy switch {
            "AllOn" => !anyHidden,
            "AnyOff" => anyHidden,
            "AllOff" => !anyVisible,
            _ => anyVisible
        };
        return !visibleByPolicy;
    }

    private static Dictionary<int, bool> ReadGroupVisibility(Dictionary<int, PdfIndirectObject> objects, out bool hasUnsupportedViewUsageApplications) {
        hasUnsupportedViewUsageApplications = false;
        var result = new Dictionary<int, bool>();
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects);
        if (catalog == null ||
            !catalog.Items.TryGetValue("OCProperties", out PdfObject? optionalContentObject) ||
            ResolveObject(optionalContentObject, objects) is not PdfDictionary optionalContent ||
            ResolveObject(optionalContent.Items.TryGetValue("OCGs", out PdfObject? groupsObject) ? groupsObject : null, objects) is not PdfArray groups) {
            return result;
        }

        PdfDictionary? defaultConfiguration = null;
        bool invalidDefaultConfiguration = false;
        if (optionalContent.Items.TryGetValue("D", out PdfObject? defaultConfigurationObject)) {
            if (defaultConfigurationObject is not PdfNull) {
                PdfObject? resolvedDefaultConfiguration = ResolveObject(defaultConfigurationObject, objects);
                if (resolvedDefaultConfiguration is PdfDictionary dictionary) defaultConfiguration = dictionary;
                else if (resolvedDefaultConfiguration is not PdfNull) invalidDefaultConfiguration = true;
            }
        }
        bool validBaseState = TryReadBaseState(defaultConfiguration, objects, out string? baseState);
        HashSet<int> onGroups = ReadReferenceSet(defaultConfiguration, "ON", objects, out bool invalidOnGroups);
        HashSet<int> offGroups = ReadReferenceSet(defaultConfiguration, "OFF", objects, out bool invalidOffGroups);
        hasUnsupportedViewUsageApplications = invalidDefaultConfiguration || !validBaseState || invalidOnGroups || invalidOffGroups;

        for (int i = 0; i < groups.Items.Count; i++) {
            if (groups.Items[i] is not PdfReference reference) {
                hasUnsupportedViewUsageApplications = true;
                continue;
            }
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject groupObject) ||
                ResolveObject(groupObject.Value, objects) is not PdfDictionary group ||
                ResolveObject(group.Items.TryGetValue("Type", out PdfObject? groupTypeObject) ? groupTypeObject : null, objects) is not PdfName { Name: "OCG" }) {
                hasUnsupportedViewUsageApplications = true;
                continue;
            }

            bool isVisible = true;
            if (string.Equals(baseState, "OFF", StringComparison.Ordinal)) {
                isVisible = onGroups.Contains(reference.ObjectNumber);
            } else if (offGroups.Contains(reference.ObjectNumber)) {
                isVisible = false;
            } else if (onGroups.Contains(reference.ObjectNumber)) {
                isVisible = true;
            }

            result[reference.ObjectNumber] = isVisible;
        }

        hasUnsupportedViewUsageApplications |=
            HasUnsupportedOptionalContentIntent(defaultConfiguration, groups, objects) ||
            ApplyViewUsageApplications(defaultConfiguration, groups, result, objects);

        return result;
    }

    private static bool TryReadBaseState(
        PdfDictionary? defaultConfiguration,
        Dictionary<int, PdfIndirectObject> objects,
        out string? baseState) {
        baseState = null;
        if (defaultConfiguration == null ||
            !defaultConfiguration.Items.TryGetValue("BaseState", out PdfObject? baseStateObject)) {
            return true;
        }

        PdfObject? resolved = ResolveObject(baseStateObject, objects);
        if (resolved is null or PdfNull) return true;
        if (resolved is not PdfName name ||
            name.Name is not ("ON" or "OFF" or "Unchanged")) {
            return false;
        }

        baseState = name.Name;
        return true;
    }

    private static bool HasUnsupportedOptionalContentIntent(
        PdfDictionary? defaultConfiguration,
        PdfArray groups,
        Dictionary<int, PdfIndirectObject> objects) {
        if (!HasDefaultViewIntent(defaultConfiguration, objects)) return true;

        for (int index = 0; index < groups.Items.Count; index++) {
            if (ResolveObject(groups.Items[index], objects) is PdfDictionary group &&
                !HasDefaultViewIntent(group, objects)) return true;
        }

        return false;
    }

    private static bool HasDefaultViewIntent(PdfDictionary? dictionary, Dictionary<int, PdfIndirectObject> objects) {
        if (dictionary == null || !dictionary.Items.TryGetValue("Intent", out PdfObject? intentObject)) return true;
        PdfObject? intent = ResolveObject(intentObject, objects);
        if (intent is PdfNull or PdfName { Name: "View" }) return true;
        return intent is PdfArray { Items.Count: 1 } names &&
            ResolveObject(names.Items[0], objects) is PdfName { Name: "View" };
    }

    private static bool ApplyViewUsageApplications(
        PdfDictionary? defaultConfiguration,
        PdfArray groups,
        Dictionary<int, bool> visibility,
        Dictionary<int, PdfIndirectObject> objects) {
        if (defaultConfiguration == null ||
            !defaultConfiguration.Items.TryGetValue("AS", out PdfObject? applicationsObject)) {
            return false;
        }
        PdfObject? resolvedApplications = ResolveObject(applicationsObject, objects);
        if (resolvedApplications is null or PdfNull) return false;
        if (resolvedApplications is not PdfArray applications) return true;

        var declaredGroups = new HashSet<long>();
        for (int groupIndex = 0; groupIndex < groups.Items.Count; groupIndex++) {
            if (groups.Items[groupIndex] is PdfReference declaredReference) {
                declaredGroups.Add(GetReferenceKey(declaredReference));
            }
        }

        bool hasUnsupportedViewUsageApplications = false;
        for (int applicationIndex = 0; applicationIndex < applications.Items.Count; applicationIndex++) {
            if (ResolveObject(applications.Items[applicationIndex], objects) is not PdfDictionary application) {
                hasUnsupportedViewUsageApplications = true;
                continue;
            }
            if (!application.Items.TryGetValue("Event", out PdfObject? eventObject) ||
                ResolveObject(eventObject, objects) is not PdfName eventName) {
                hasUnsupportedViewUsageApplications = true;
                continue;
            }
            if (!string.Equals(eventName.Name, "View", StringComparison.Ordinal)) {
                continue;
            }
            if (!HasExactViewCategory(application, objects)) {
                hasUnsupportedViewUsageApplications = true;
                continue;
            }

            PdfArray targets;
            if (!application.Items.TryGetValue("OCGs", out PdfObject? targetObject)) {
                targets = groups;
            } else {
                PdfObject? resolvedTargets = ResolveObject(targetObject, objects);
                if (resolvedTargets is PdfNull) targets = groups;
                else if (resolvedTargets is PdfArray targetArray) targets = targetArray;
                else {
                    hasUnsupportedViewUsageApplications = true;
                    continue;
                }
            }
            for (int targetIndex = 0; targetIndex < targets.Items.Count; targetIndex++) {
                if (targets.Items[targetIndex] is not PdfReference reference ||
                    !declaredGroups.Contains(GetReferenceKey(reference))) {
                    hasUnsupportedViewUsageApplications = true;
                    continue;
                }
                if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject groupObject) ||
                    ResolveObject(groupObject.Value, objects) is not PdfDictionary group) {
                    hasUnsupportedViewUsageApplications = true;
                    continue;
                }

                if (!group.Items.TryGetValue("Usage", out PdfObject? usageObject)) continue;
                PdfObject? resolvedUsage = ResolveObject(usageObject, objects);
                if (resolvedUsage is PdfNull) continue;
                if (resolvedUsage is not PdfDictionary usage) {
                    hasUnsupportedViewUsageApplications = true;
                    continue;
                }
                if (!usage.Items.TryGetValue("View", out PdfObject? viewObject)) continue;
                PdfObject? resolvedView = ResolveObject(viewObject, objects);
                if (resolvedView is PdfNull) continue;
                if (resolvedView is not PdfDictionary view) {
                    hasUnsupportedViewUsageApplications = true;
                    continue;
                }

                if (!view.Items.TryGetValue("ViewState", out PdfObject? viewStateObject)) {
                    visibility[reference.ObjectNumber] = true;
                    continue;
                }
                PdfObject? resolvedViewState = ResolveObject(viewStateObject, objects);
                if (resolvedViewState is PdfNull) visibility[reference.ObjectNumber] = true;
                else if (resolvedViewState is PdfName { Name: "ON" }) visibility[reference.ObjectNumber] = true;
                else if (resolvedViewState is PdfName { Name: "OFF" }) visibility[reference.ObjectNumber] = false;
                else hasUnsupportedViewUsageApplications = true;
            }
        }
        return hasUnsupportedViewUsageApplications;
    }

    private static long GetReferenceKey(PdfReference reference) =>
        ((long)reference.ObjectNumber << 32) | (uint)reference.Generation;

    private static bool HasExactViewCategory(PdfDictionary application, Dictionary<int, PdfIndirectObject> objects) =>
        ResolveObject(application.Items.TryGetValue("Category", out PdfObject? value) ? value : null, objects) is PdfArray { Items.Count: 1 } names &&
        ResolveObject(names.Items[0], objects) is PdfName { Name: "View" };

    private static HashSet<int> ReadReferenceSet(
        PdfDictionary? dictionary,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        out bool invalid) {
        var result = new HashSet<int>();
        invalid = false;
        if (dictionary == null || !dictionary.Items.TryGetValue(key, out PdfObject? value)) {
            return result;
        }
        PdfObject? resolved = ResolveObject(value, objects);
        if (resolved is PdfNull) {
            return result;
        }
        if (resolved is not PdfArray array) {
            invalid = true;
            return result;
        }

        for (int i = 0; i < array.Items.Count; i++) {
            if (array.Items[i] is not PdfReference reference ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject groupObject) ||
                ResolveObject(groupObject.Value, objects) is not PdfDictionary group ||
                ResolveObject(group.Items.TryGetValue("Type", out PdfObject? groupTypeObject) ? groupTypeObject : null, objects) is not PdfName { Name: "OCG" }) {
                invalid = true;
                continue;
            }
            result.Add(reference.ObjectNumber);
        }

        return result;
    }

    private static string? ReadName(PdfDictionary? dictionary, string key, Dictionary<int, PdfIndirectObject> objects) {
        if (dictionary == null ||
            ResolveObject(dictionary.Items.TryGetValue(key, out PdfObject? value) ? value : null, objects) is not PdfName name ||
            string.IsNullOrEmpty(name.Name)) {
            return null;
        }

        return name.Name;
    }

    private static bool IsOptionalContentObjectHidden(
        PdfObject value,
        Dictionary<int, bool> groupVisibility,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<int> visited,
        int maxExpressionDepth,
        int depth) {
        if (depth > maxExpressionDepth) {
            return false;
        }
        if (value is PdfReference reference) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) {
                return false;
            }
            if (groupVisibility.TryGetValue(reference.ObjectNumber, out bool groupVisible)) {
                return !groupVisible;
            }

            if (!visited.Add(reference.ObjectNumber)) {
                return false;
            }
            try {
                return IsOptionalContentObjectHidden(indirect.Value, groupVisibility, objects, visited, maxExpressionDepth, depth + 1);
            } finally {
                visited.Remove(reference.ObjectNumber);
            }
        }

        if (ResolveObject(value, objects) is not PdfDictionary dictionary) {
            return false;
        }

        string? type = ReadName(dictionary, "Type", objects);
        if (!string.Equals(type, "OCMD", StringComparison.Ordinal)) {
            return false;
        }

        if (dictionary.Items.TryGetValue("VE", out PdfObject? expressionObject) &&
            TryEvaluateVisibilityExpression(expressionObject, groupVisibility, objects, new HashSet<int>(), maxExpressionDepth, depth + 1, out bool expressionVisible)) {
            return !expressionVisible;
        }

        List<bool> visibilities = ReadOptionalContentMembershipGroupVisibilities(dictionary, groupVisibility, objects);
        string policy = ReadName(dictionary, "P", objects) ?? "AnyOn";
        bool visible = policy switch {
            "AllOn" => visibilities.TrueForAll(static visible => visible),
            "AnyOff" => visibilities.Exists(static visible => !visible),
            "AllOff" => visibilities.TrueForAll(static visible => !visible),
            _ => visibilities.Exists(static visible => visible)
        };
        return !visible;
    }

    private static bool IsOptionalContentObjectInvalid(
        PdfObject value,
        Dictionary<int, bool> groupVisibility,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<int> visited,
        int maxExpressionDepth,
        int depth) {
        if (depth > maxExpressionDepth) return true;
        if (value is PdfReference reference) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject indirect)) return true;
            if (groupVisibility.ContainsKey(reference.ObjectNumber)) {
                return ResolveObject(indirect.Value, objects) is not PdfDictionary group ||
                    ResolveObject(group.Items.TryGetValue("Type", out PdfObject? typeObject) ? typeObject : null, objects) is not PdfName { Name: "OCG" };
            }
            if (!visited.Add(reference.ObjectNumber)) return true;
            try {
                return IsOptionalContentObjectInvalid(indirect.Value, groupVisibility, objects, visited, maxExpressionDepth, depth + 1);
            } finally {
                visited.Remove(reference.ObjectNumber);
            }
        }

        if (ResolveObject(value, objects) is not PdfDictionary dictionary ||
            !string.Equals(ReadName(dictionary, "Type", objects), "OCMD", StringComparison.Ordinal)) return true;
        if (dictionary.Items.TryGetValue("P", out PdfObject? policyObject)) {
            PdfObject? policy = ResolveObject(policyObject, objects);
            if (policy is null || policy is not PdfNull and not PdfName { Name: "AnyOn" or "AllOn" or "AnyOff" or "AllOff" }) return true;
        }
        if (dictionary.Items.TryGetValue("VE", out PdfObject? expressionObject)) {
            PdfObject? expression = ResolveObject(expressionObject, objects);
            if (expression is null ||
                expression is not PdfNull &&
                !TryEvaluateVisibilityExpression(expressionObject, groupVisibility, objects, new HashSet<int>(), maxExpressionDepth, depth + 1, out _)) return true;
        }
        return dictionary.Items.TryGetValue("OCGs", out PdfObject? groupsObject) &&
            HasInvalidOptionalContentGroups(groupsObject, groupVisibility, objects, new HashSet<int>(), maxExpressionDepth, depth + 1, allowArray: true);
    }

    private static bool HasInvalidOptionalContentGroups(
        PdfObject value,
        Dictionary<int, bool> groupVisibility,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<int> visited,
        int maxExpressionDepth,
        int depth,
        bool allowArray) {
        if (depth > maxExpressionDepth) return true;
        if (value is PdfReference reference) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject indirect)) return true;
            if (groupVisibility.ContainsKey(reference.ObjectNumber)) {
                return ResolveObject(indirect.Value, objects) is not PdfDictionary group ||
                    ResolveObject(group.Items.TryGetValue("Type", out PdfObject? typeObject) ? typeObject : null, objects) is not PdfName { Name: "OCG" };
            }
            if (!visited.Add(reference.ObjectNumber)) return true;
            try {
                return HasInvalidOptionalContentGroups(indirect.Value, groupVisibility, objects, visited, maxExpressionDepth, depth + 1, allowArray);
            } finally {
                visited.Remove(reference.ObjectNumber);
            }
        }

        PdfObject? resolved = ResolveObject(value, objects);
        if (resolved is PdfNull) return false;
        if (!allowArray || resolved is not PdfArray groups) return true;
        for (int index = 0; index < groups.Items.Count; index++) {
            if (HasInvalidOptionalContentGroups(groups.Items[index], groupVisibility, objects, visited, maxExpressionDepth, depth + 1, allowArray: false)) return true;
        }
        return false;
    }

    private static List<bool> ReadOptionalContentMembershipGroupVisibilities(
        PdfDictionary dictionary,
        Dictionary<int, bool> groupVisibility,
        Dictionary<int, PdfIndirectObject> objects) {
        var visibilities = new List<bool>();
        if (!dictionary.Items.TryGetValue("OCGs", out PdfObject? groupsObject)) {
            return visibilities;
        }

        PdfObject? resolved = ResolveObject(groupsObject, objects);
        if (resolved is PdfArray groups) {
            for (int i = 0; i < groups.Items.Count; i++) {
                AddOptionalContentGroupVisibility(groups.Items[i], groupVisibility, objects, visibilities);
            }

            return visibilities;
        }

        AddOptionalContentGroupVisibility(groupsObject, groupVisibility, objects, visibilities);
        return visibilities;
    }

    private static void AddOptionalContentGroupVisibility(
        PdfObject value,
        Dictionary<int, bool> groupVisibility,
        Dictionary<int, PdfIndirectObject> objects,
        List<bool> visibilities) {
        if (value is PdfReference reference) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) return;
            if (groupVisibility.TryGetValue(reference.ObjectNumber, out bool visible)) {
                visibilities.Add(visible);
                return;
            }
            value = indirect.Value;
        }

        if (ResolveObject(value, objects) is PdfArray nested) {
            for (int i = 0; i < nested.Items.Count; i++) {
                AddOptionalContentGroupVisibility(nested.Items[i], groupVisibility, objects, visibilities);
            }
        }
    }

    private static bool TryEvaluateVisibilityExpression(
        PdfObject value,
        Dictionary<int, bool> groupVisibility,
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<int> visited,
        int maxExpressionDepth,
        int depth,
        out bool visible) {
        visible = false;
        if (depth > maxExpressionDepth) {
            return false;
        }
        if (value is PdfReference reference) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) {
                return false;
            }
            if (groupVisibility.TryGetValue(reference.ObjectNumber, out visible)) {
                return true;
            }

            if (!visited.Add(reference.ObjectNumber)) {
                return false;
            }
            try {
                return TryEvaluateVisibilityExpression(indirect.Value, groupVisibility, objects, visited, maxExpressionDepth, depth + 1, out visible);
            } finally {
                visited.Remove(reference.ObjectNumber);
            }
        }

        PdfObject? resolved = ResolveObject(value, objects);
        if (resolved is PdfDictionary dictionary) {
            if (string.Equals(ReadName(dictionary, "Type", objects), "OCMD", StringComparison.Ordinal)) {
                visible = !IsOptionalContentObjectHidden(dictionary, groupVisibility, objects, visited, maxExpressionDepth, depth + 1);
                return true;
            }

            return false;
        }

        if (resolved is not PdfArray expression ||
            expression.Items.Count == 0 ||
            ResolveObject(expression.Items[0], objects) is not PdfName operatorName) {
            return false;
        }

        switch (operatorName.Name) {
            case "And":
                if (expression.Items.Count < 2) return false;
                visible = true;
                for (int i = 1; i < expression.Items.Count; i++) {
                    if (!TryEvaluateVisibilityExpression(expression.Items[i], groupVisibility, objects, visited, maxExpressionDepth, depth + 1, out bool operandVisible)) {
                        return false;
                    }

                    visible &= operandVisible;
                }

                return true;
            case "Or":
                if (expression.Items.Count < 2) return false;
                visible = false;
                for (int i = 1; i < expression.Items.Count; i++) {
                    if (!TryEvaluateVisibilityExpression(expression.Items[i], groupVisibility, objects, visited, maxExpressionDepth, depth + 1, out bool operandVisible)) {
                        return false;
                    }

                    visible |= operandVisible;
                }

                return true;
            case "Not":
                if (expression.Items.Count != 2 ||
                    !TryEvaluateVisibilityExpression(expression.Items[1], groupVisibility, objects, visited, maxExpressionDepth, depth + 1, out bool nestedVisible)) {
                    return false;
                }

                visible = !nestedVisible;
                return true;
            default:
                return false;
        }
    }

    private static PdfObject? ResolveObject(PdfObject? value, Dictionary<int, PdfIndirectObject> objects) =>
        PdfObjectLookup.Resolve(objects, value);
}
