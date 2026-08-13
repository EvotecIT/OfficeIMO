namespace OfficeIMO.Pdf;

internal sealed class PdfPageOptionalContentVisibility {
    private readonly Dictionary<string, bool> _hiddenProperties;
    private readonly HashSet<int> _hiddenObjectNumbers;
    private readonly Dictionary<int, bool> _groupVisibility;
    private readonly Dictionary<int, PdfIndirectObject> _objects;
    private readonly int _maxExpressionDepth;

    private PdfPageOptionalContentVisibility(Dictionary<string, bool> hiddenProperties, HashSet<int> hiddenObjectNumbers, Dictionary<int, bool> groupVisibility, Dictionary<int, PdfIndirectObject> objects, int maxExpressionDepth) {
        _hiddenProperties = hiddenProperties;
        _hiddenObjectNumbers = hiddenObjectNumbers;
        _groupVisibility = groupVisibility;
        _objects = objects;
        _maxExpressionDepth = maxExpressionDepth;
    }

    public static PdfPageOptionalContentVisibility? Create(
        PdfDictionary? resources,
        Dictionary<int, PdfIndirectObject> objects,
        int maxExpressionDepth) {
        int effectiveMaxExpressionDepth = System.Math.Min(maxExpressionDepth, PdfReadLimits.DefaultMaxContentNestingDepth);
        Dictionary<int, bool> groupVisibility = ReadGroupVisibility(objects);
        if (groupVisibility.Count == 0) {
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
        if (resources != null &&
            resources.Items.TryGetValue("Properties", out PdfObject? propertiesObject) &&
            ResolveObject(propertiesObject, objects) is PdfDictionary properties) {
            foreach (KeyValuePair<string, PdfObject> entry in properties.Items) {
                if (IsOptionalContentObjectHidden(entry.Value, groupVisibility, objects, new HashSet<int>(), effectiveMaxExpressionDepth, depth: 0)) {
                    hiddenProperties[entry.Key] = true;
                }
            }
        }

        return new PdfPageOptionalContentVisibility(hiddenProperties, hiddenObjectNumbers, groupVisibility, objects, effectiveMaxExpressionDepth);
    }

    public bool IsHidden(string propertyName) =>
        _hiddenProperties.TryGetValue(propertyName, out bool hidden) && hidden;

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
                if (TryEvaluateInlineVisibilityExpression(expression, out bool expressionVisible)) return !expressionVisible;
                int index = 0;
                SkipInlineWhitespace(expression, ref index);
                if (TryReadInlineReference(expression, ref index, out int objectNumber)) {
                    SkipInlineWhitespace(expression, ref index);
                    if (index == expression.Length &&
                        _objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect) &&
                        TryEvaluateVisibilityExpression(indirect.Value, _groupVisibility, _objects, new HashSet<int>(), _maxExpressionDepth, depth: 0, out expressionVisible)) {
                        return !expressionVisible;
                    }
                }
            }

            return IsMembershipHidden(references.ObjectNumbers, references.Policy);
        }

        return IsHiddenAny(references.ObjectNumbers);
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

        if (TryReadInlineReference(expression, ref index, out int objectNumber)) {
            if (_groupVisibility.TryGetValue(objectNumber, out visible)) {
                return true;
            }

            return _objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect) &&
                TryEvaluateVisibilityExpression(
                    indirect.Value,
                    _groupVisibility,
                    _objects,
                    new HashSet<int> { objectNumber },
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

    private static bool TryReadInlineReference(string text, ref int index, out int objectNumber) {
        objectNumber = 0;
        SkipInlineWhitespace(text, ref index);
        int start = index;
        if (!TryReadInlineInteger(text, ref index, out objectNumber)) {
            return false;
        }

        SkipInlineWhitespace(text, ref index);
        if (!TryReadInlineInteger(text, ref index, out _)) {
            index = start;
            return false;
        }

        SkipInlineWhitespace(text, ref index);
        if (index >= text.Length || text[index] != 'R' || !IsInlineTokenBoundary(text, index + 1)) {
            index = start;
            return false;
        }

        index++;
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

    private bool IsMembershipHidden(IReadOnlyList<int> objectNumbers, string? policy) {
        if (objectNumbers.Count == 0) {
            return false;
        }

        bool anyVisible = false;
        bool anyHidden = false;
        for (int i = 0; i < objectNumbers.Count; i++) {
            bool visible = !_hiddenObjectNumbers.Contains(objectNumbers[i]);
            anyVisible |= visible;
            anyHidden |= !visible;
        }

        bool visibleByPolicy = policy switch {
            "AllOn" => !anyHidden,
            "AnyOff" => anyHidden,
            "AllOff" => !anyVisible,
            _ => anyVisible
        };
        return !visibleByPolicy;
    }

    private static Dictionary<int, bool> ReadGroupVisibility(Dictionary<int, PdfIndirectObject> objects) {
        var result = new Dictionary<int, bool>();
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects);
        if (catalog == null ||
            !catalog.Items.TryGetValue("OCProperties", out PdfObject? optionalContentObject) ||
            ResolveObject(optionalContentObject, objects) is not PdfDictionary optionalContent ||
            ResolveObject(optionalContent.Items.TryGetValue("OCGs", out PdfObject? groupsObject) ? groupsObject : null, objects) is not PdfArray groups) {
            return result;
        }

        PdfDictionary? defaultConfiguration = ResolveObject(
            optionalContent.Items.TryGetValue("D", out PdfObject? defaultConfigurationObject) ? defaultConfigurationObject : null,
            objects) as PdfDictionary;
        string? baseState = ReadName(defaultConfiguration, "BaseState", objects);
        HashSet<int> onGroups = ReadReferenceSet(defaultConfiguration, "ON", objects);
        HashSet<int> offGroups = ReadReferenceSet(defaultConfiguration, "OFF", objects);

        for (int i = 0; i < groups.Items.Count; i++) {
            if (groups.Items[i] is not PdfReference reference) {
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

        return result;
    }

    private static HashSet<int> ReadReferenceSet(PdfDictionary? dictionary, string key, Dictionary<int, PdfIndirectObject> objects) {
        var result = new HashSet<int>();
        if (dictionary == null ||
            ResolveObject(dictionary.Items.TryGetValue(key, out PdfObject? value) ? value : null, objects) is not PdfArray array) {
            return result;
        }

        for (int i = 0; i < array.Items.Count; i++) {
            AddReferenceObjectNumbers(array.Items[i], objects, result);
        }

        return result;
    }

    private static void AddReferenceObjectNumbers(PdfObject value, Dictionary<int, PdfIndirectObject> objects, HashSet<int> result) {
        if (value is PdfReference reference) {
            result.Add(reference.ObjectNumber);
            return;
        }

        if (ResolveObject(value, objects) is PdfArray nested) {
            for (int i = 0; i < nested.Items.Count; i++) {
                AddReferenceObjectNumbers(nested.Items[i], objects, result);
            }
        }
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
            if (groupVisibility.TryGetValue(reference.ObjectNumber, out bool groupVisible)) {
                return !groupVisible;
            }

            if (!visited.Add(reference.ObjectNumber) ||
                !objects.TryGetValue(reference.ObjectNumber, out PdfIndirectObject? indirect)) {
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

        List<bool> visibilities = ReadOptionalContentMembershipGroupVisibilities(dictionary, groupVisibility, objects);
        if (visibilities.Count == 0) {
            return false;
        }

        if (dictionary.Items.TryGetValue("VE", out PdfObject? expressionObject) &&
            TryEvaluateVisibilityExpression(expressionObject, groupVisibility, objects, new HashSet<int>(), maxExpressionDepth, depth + 1, out bool expressionVisible)) {
            return !expressionVisible;
        }

        string policy = ReadName(dictionary, "P", objects) ?? "AnyOn";
        bool visible = policy switch {
            "AllOn" => visibilities.TrueForAll(static visible => visible),
            "AnyOff" => visibilities.Exists(static visible => !visible),
            "AllOff" => visibilities.TrueForAll(static visible => !visible),
            _ => visibilities.Exists(static visible => visible)
        };
        return !visible;
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
        if (value is PdfReference reference &&
            groupVisibility.TryGetValue(reference.ObjectNumber, out bool visible)) {
            visibilities.Add(visible);
            return;
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
            if (groupVisibility.TryGetValue(reference.ObjectNumber, out visible)) {
                return true;
            }

            if (!visited.Add(reference.ObjectNumber)) {
                return false;
            }
            try {
                return objects.TryGetValue(reference.ObjectNumber, out PdfIndirectObject? indirect) &&
                    TryEvaluateVisibilityExpression(indirect.Value, groupVisibility, objects, visited, maxExpressionDepth, depth + 1, out visible);
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
