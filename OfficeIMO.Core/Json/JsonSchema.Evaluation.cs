using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Json;

public sealed partial class JsonSchema {
    private const int MaxListedValues = 8;

    /// <summary>One schema resource on the dynamic scope, innermost first.</summary>
    private sealed class ScopeEntry {
        public ScopeEntry(ScopeEntry? parent, JsonSchemaResource resource) {
            Parent = parent;
            Resource = resource;
        }

        public ScopeEntry? Parent { get; }
        public JsonSchemaResource Resource { get; }
    }

    private sealed class EvaluationContext {
        private readonly JsonSchemaOptions _options;
        private readonly List<JsonSchemaFinding> _findings = new List<JsonSchemaFinding>();
        private List<string>? _unchecked;

        public EvaluationContext(JsonSchema schema, bool report) {
            _options = schema._options;
            Report = report;
        }

        public bool Report { get; }
        public bool Aborted { get; private set; }
        public bool Truncated { get; private set; }
        public long Steps;

        public JsonSchemaOptions Options => _options;

        /// <summary>Records a finding; returns false when validation must stop.</summary>
        public bool Add(JsonPointerPath? instancePath, JsonPointerPath? schemaPath, string keyword, string message, JsonData instance) {
            if (Aborted) return false;
            if (_findings.Count >= _options.MaxErrors) {
                Truncated = true;
                Aborted = true;
                return false;
            }
            _findings.Add(new JsonSchemaFinding(JsonPointerPath.Format(instancePath), JsonPointerPath.Format(schemaPath), keyword, message, instance));
            return true;
        }

        /// <summary>Stops validation at a limit; the finding is recorded even while probing, so the result is never valid.</summary>
        public void Abort(JsonPointerPath? instancePath, JsonPointerPath? schemaPath, string keyword, string message, JsonData instance) {
            if (Aborted) return;
            _findings.Add(new JsonSchemaFinding(JsonPointerPath.Format(instancePath), JsonPointerPath.Format(schemaPath), keyword, message, instance));
            Aborted = true;
        }

        public void NoteUncheckedFormat(string format) {
            _unchecked ??= new List<string>();
            if (!_unchecked.Contains(format)) _unchecked.Add(format);
        }

        public JsonSchemaValidationResult ToResult() =>
            new JsonSchemaValidationResult(_findings.AsReadOnly(), Truncated, (_unchecked ?? new List<string>()).AsReadOnly());
    }

    private bool Evaluate(JsonSchemaNode node, JsonData instance, JsonPointerPath? ip, JsonPointerPath? sp, bool report,
        JsonSchemaAnnotations? annotations, ScopeEntry? scope, int depth, EvaluationContext cx) {
        if (cx.Aborted) return false;
        JsonSchemaOptions options = cx.Options;
        if (++cx.Steps > options.MaxEvaluationSteps) {
            cx.Abort(ip, sp, "$schema", "Validation stopped after " + options.MaxEvaluationSteps.ToString(CultureInfo.InvariantCulture) + " schema evaluations (JsonSchemaOptions.MaxEvaluationSteps).", instance);
            return false;
        }
        if (depth > options.MaxDepth) {
            cx.Abort(ip, sp, "$ref", "Validation stopped: schemas nest deeper than " + options.MaxDepth.ToString(CultureInfo.InvariantCulture) + " levels here (JsonSchemaOptions.MaxDepth), which usually means a reference cycle.", instance);
            return false;
        }
        if (node.Boolean is bool allowed) {
            if (allowed) return true;
            if (report) cx.Add(ip, sp, "false", "No value is allowed here.", instance);
            return false;
        }
        if (scope is null || !ReferenceEquals(scope.Resource, node.Resource)) scope = new ScopeEntry(scope, node.Resource);
        JsonSchemaAnnotations? local = instance.Kind is JsonDataKind.Object or JsonDataKind.Array ? new JsonSchemaAnnotations() : null;
        bool valid = true;

        if (node.Ref is not null) {
            JsonSchemaNode? target = GetRefTarget(node, out string? error);
            if (target is null) {
                valid = false;
                if (!report || !cx.Add(ip, Child(sp, "$ref"), "$ref", error!, instance)) return false;
            } else if (!Evaluate(target, instance, ip, Child(sp, "$ref"), report, local, scope, depth + 1, cx)) {
                valid = false;
                if (!report || cx.Aborted) return false;
            }
            if (node.RefOnly) return Finish(valid, annotations, local);
        }
        if (node.DynamicRef is not null) {
            JsonSchemaNode? target = GetDynamicRefTarget(node, scope, out string? error);
            if (target is null) {
                valid = false;
                if (!report || !cx.Add(ip, Child(sp, "$dynamicRef"), "$dynamicRef", error!, instance)) return false;
            } else if (!Evaluate(target, instance, ip, Child(sp, "$dynamicRef"), report, local, scope, depth + 1, cx)) {
                valid = false;
                if (!report || cx.Aborted) return false;
            }
        }

        if (node.Types != JsonSchemaTypes.None && !MatchesType(node.Types, instance)) {
            valid = false;
            if (!report || !cx.Add(ip, Child(sp, "type"), "type", "Expected " + DescribeTypes(node.Types) + " but found " + DescribeValueType(instance) + ".", instance)) return false;
        }
        if (node.Enum is not null && !node.Enum.Any(item => JsonData.DeepEquals(item, instance))) {
            valid = false;
            if (!report || !cx.Add(ip, Child(sp, "enum"), "enum", "Value must be one of " + ListValues(node.Enum) + ".", instance)) return false;
        }
        if (node.Const is not null && !JsonData.DeepEquals(node.Const, instance)) {
            valid = false;
            if (!report || !cx.Add(ip, Child(sp, "const"), "const", "Value must be " + Preview(node.Const) + ".", instance)) return false;
        }

        switch (instance.Kind) {
            case JsonDataKind.Number:
                if (!EvaluateNumber(node, instance, ip, sp, report, cx)) {
                    valid = false;
                    if (!report || cx.Aborted) return false;
                }
                break;
            case JsonDataKind.String:
                if (!EvaluateString(node, instance, ip, sp, report, cx)) {
                    valid = false;
                    if (!report || cx.Aborted) return false;
                }
                break;
            case JsonDataKind.Array:
                if (!EvaluateArray(node, instance, ip, sp, report, local!, scope, depth, cx)) {
                    valid = false;
                    if (!report || cx.Aborted) return false;
                }
                break;
            case JsonDataKind.Object:
                if (!EvaluateObject(node, instance, ip, sp, report, local!, scope, depth, cx)) {
                    valid = false;
                    if (!report || cx.Aborted) return false;
                }
                break;
        }

        if (!EvaluateApplicators(node, instance, ip, sp, report, local, scope, depth, cx)) {
            valid = false;
            if (!report || cx.Aborted) return false;
        }

        if (node.UnevaluatedItems is not null && instance.Kind == JsonDataKind.Array) {
            IReadOnlyList<JsonData> items = instance.Items;
            for (int i = 0; i < items.Count; i++) {
                if (local!.IsItemEvaluated(i)) continue;
                if (node.UnevaluatedItems.Boolean == false) {
                    valid = false;
                    if (!report || !cx.Add(Child(ip, i), Child(sp, "unevaluatedItems"), "unevaluatedItems", "Item " + i.ToString(CultureInfo.InvariantCulture) + " is not allowed: no schema in 'prefixItems', 'items' or 'contains' covers it.", items[i])) return false;
                } else if (!Evaluate(node.UnevaluatedItems, items[i], Child(ip, i), Child(sp, "unevaluatedItems"), report, null, scope, depth + 1, cx)) {
                    valid = false;
                    if (!report || cx.Aborted) return false;
                }
            }
            local!.AllItems = true;
        }
        if (node.UnevaluatedProperties is not null && instance.Kind == JsonDataKind.Object) {
            foreach (KeyValuePair<string, JsonData> property in instance.Properties) {
                if (local!.IsPropertyEvaluated(property.Key)) continue;
                if (node.UnevaluatedProperties.Boolean == false) {
                    valid = false;
                    if (!report || !cx.Add(Child(ip, property.Key), Child(sp, "unevaluatedProperties"), "unevaluatedProperties", "Property '" + property.Key + "' is not allowed.", property.Value)) return false;
                } else if (!Evaluate(node.UnevaluatedProperties, property.Value, Child(ip, property.Key), Child(sp, "unevaluatedProperties"), report, null, scope, depth + 1, cx)) {
                    valid = false;
                    if (!report || cx.Aborted) return false;
                }
            }
            local!.AllProperties = true;
        }

        return Finish(valid, annotations, local);
    }

    private static bool Finish(bool valid, JsonSchemaAnnotations? annotations, JsonSchemaAnnotations? local) {
        if (valid && annotations is not null && local is not null) annotations.Merge(local);
        return valid;
    }

    private static JsonPointerPath? Child(JsonPointerPath? parent, string name) => JsonPointerPath.Child(parent, name);

    private static JsonPointerPath? Child(JsonPointerPath? parent, int index) => JsonPointerPath.Child(parent, index);

    private JsonSchemaNode? GetRefTarget(JsonSchemaNode node, out string? error) {
        lock (_sync) {
            if (!node.RefResolved) {
                node.RefTarget = ResolveReference(node, node.Ref!, out node.RefError);
                node.RefResolved = true;
            }
            error = node.RefError;
            return node.RefTarget;
        }
    }

    private JsonSchemaNode? GetDynamicRefTarget(JsonSchemaNode node, ScopeEntry scope, out string? error) {
        lock (_sync) {
            if (!node.DynamicRefResolved) {
                node.DynamicRefTarget = ResolveReference(node, node.DynamicRef!, out node.DynamicRefError);
                // The reference is dynamic only when it names an anchor that its initial target declares as dynamic.
                string fragment = JsonSchemaUri.Fragment(JsonSchemaUri.Resolve(node.Resource.Uri, node.DynamicRef!));
                if (node.DynamicRefTarget is not null && fragment.Length != 0 && fragment[0] != '/'
                    && node.DynamicRefTarget.Resource.DynamicAnchors is not null
                    && node.DynamicRefTarget.Resource.DynamicAnchors.TryGetValue(fragment, out JsonData? declared)
                    && ReferenceEquals(declared, node.DynamicRefTarget.Data)) {
                    node.DynamicAnchorName = fragment;
                }
                node.DynamicRefResolved = true;
            }
            error = node.DynamicRefError;
            if (node.DynamicRefTarget is null || node.DynamicAnchorName is null) return node.DynamicRefTarget;
        }
        return ResolveDynamicAnchor(node.DynamicAnchorName, scope) ?? node.DynamicRefTarget;
    }

    #region Scalars

    private static bool MatchesType(JsonSchemaTypes types, JsonData instance) => instance.Kind switch {
        JsonDataKind.Null => (types & JsonSchemaTypes.Null) != 0,
        JsonDataKind.Boolean => (types & JsonSchemaTypes.Boolean) != 0,
        JsonDataKind.String => (types & JsonSchemaTypes.String) != 0,
        JsonDataKind.Array => (types & JsonSchemaTypes.Array) != 0,
        JsonDataKind.Object => (types & JsonSchemaTypes.Object) != 0,
        _ => (types & JsonSchemaTypes.Number) != 0 || ((types & JsonSchemaTypes.Integer) != 0 && instance.IsInteger)
    };

    private bool EvaluateNumber(JsonSchemaNode node, JsonData instance, JsonPointerPath? ip, JsonPointerPath? sp, bool report, EvaluationContext cx) {
        bool valid = true;
        if (node.MultipleOf is not null && !IsMultipleOf(instance, node.MultipleOf)) {
            valid = false;
            if (!report || !cx.Add(ip, Child(sp, "multipleOf"), "multipleOf", "Value must be a multiple of " + Preview(node.MultipleOf) + ".", instance)) return false;
        }
        if (node.Maximum is not null && JsonData.CompareNumbers(instance, node.Maximum) > 0) {
            valid = false;
            if (!report || !cx.Add(ip, Child(sp, "maximum"), "maximum", "Value must be at most " + Preview(node.Maximum) + ".", instance)) return false;
        }
        if (node.ExclusiveMaximum is not null && JsonData.CompareNumbers(instance, node.ExclusiveMaximum) >= 0) {
            valid = false;
            if (!report || !cx.Add(ip, Child(sp, "exclusiveMaximum"), "exclusiveMaximum", "Value must be less than " + Preview(node.ExclusiveMaximum) + ".", instance)) return false;
        }
        if (node.Minimum is not null && JsonData.CompareNumbers(instance, node.Minimum) < 0) {
            valid = false;
            if (!report || !cx.Add(ip, Child(sp, "minimum"), "minimum", "Value must be at least " + Preview(node.Minimum) + ".", instance)) return false;
        }
        if (node.ExclusiveMinimum is not null && JsonData.CompareNumbers(instance, node.ExclusiveMinimum) <= 0) {
            valid = false;
            if (!report || !cx.Add(ip, Child(sp, "exclusiveMinimum"), "exclusiveMinimum", "Value must be greater than " + Preview(node.ExclusiveMinimum) + ".", instance)) return false;
        }
        return valid;
    }

    private static bool IsMultipleOf(JsonData value, JsonData divisor) {
        if (value.HasExact && divisor.HasExact) {
            bool? exact = JsonExactNumber.IsMultipleOf(value.Exact, divisor.Exact);
            if (exact is not null) return exact.Value;
        }
        double quotient = value.DoubleValue / divisor.DoubleValue;
        if (double.IsInfinity(quotient) || double.IsNaN(quotient)) return false;
        return Math.Floor(quotient) == quotient;
    }

    private bool EvaluateString(JsonSchemaNode node, JsonData instance, JsonPointerPath? ip, JsonPointerPath? sp, bool report, EvaluationContext cx) {
        bool valid = true;
        string text = instance.GetString();
        if (node.MaxLength >= 0 || node.MinLength >= 0) {
            long length = CountCodePoints(text);
            if (node.MaxLength >= 0 && length > node.MaxLength) {
                valid = false;
                if (!report || !cx.Add(ip, Child(sp, "maxLength"), "maxLength", "Text must be at most " + node.MaxLength.ToString(CultureInfo.InvariantCulture) + " characters long; it has " + length.ToString(CultureInfo.InvariantCulture) + ".", instance)) return false;
            }
            if (node.MinLength >= 0 && length < node.MinLength) {
                valid = false;
                if (!report || !cx.Add(ip, Child(sp, "minLength"), "minLength", "Text must be at least " + node.MinLength.ToString(CultureInfo.InvariantCulture) + " characters long; it has " + length.ToString(CultureInfo.InvariantCulture) + ".", instance)) return false;
            }
        }
        if (node.Pattern is not null) {
            bool? matches = TryMatch(node.Pattern, text);
            if (matches is null) {
                cx.Abort(ip, Child(sp, "pattern"), "pattern", "Matching the pattern '" + node.PatternText + "' timed out (JsonSchemaOptions.RegexTimeout).", instance);
                return false;
            }
            if (matches == false) {
                valid = false;
                if (!report || !cx.Add(ip, Child(sp, "pattern"), "pattern", "Text must match the pattern '" + node.PatternText + "'.", instance)) return false;
            }
        }
        if (node.Format is not null) {
            bool? formatValid = cx.Options.ValidateFormats ? JsonSchemaFormats.Check(node.Format, text, cx.Options.RegexTimeout) : null;
            if (formatValid is null) {
                cx.NoteUncheckedFormat(node.Format);
            } else if (formatValid == false) {
                valid = false;
                if (!report || !cx.Add(ip, Child(sp, "format"), "format", "Text is not a valid " + node.Format + ".", instance)) return false;
            }
        }
        return valid;
    }

    private static long CountCodePoints(string text) {
        long count = text.Length;
        for (int i = 0; i + 1 < text.Length; i++) {
            if (char.IsHighSurrogate(text[i]) && char.IsLowSurrogate(text[i + 1])) {
                count--;
                i++;
            }
        }
        return count;
    }

    private static bool? TryMatch(Regex regex, string text) {
        try {
            return regex.IsMatch(text);
        } catch (RegexMatchTimeoutException) {
            return null;
        }
    }

    #endregion

    #region Arrays and objects

    private bool EvaluateArray(JsonSchemaNode node, JsonData instance, JsonPointerPath? ip, JsonPointerPath? sp, bool report,
        JsonSchemaAnnotations local, ScopeEntry scope, int depth, EvaluationContext cx) {
        bool valid = true;
        IReadOnlyList<JsonData> items = instance.Items;
        int count = items.Count;
        if (node.MaxItems >= 0 && count > node.MaxItems) {
            valid = false;
            if (!report || !cx.Add(ip, Child(sp, "maxItems"), "maxItems", "Array must have at most " + node.MaxItems.ToString(CultureInfo.InvariantCulture) + " items; it has " + count.ToString(CultureInfo.InvariantCulture) + ".", instance)) return false;
        }
        if (node.MinItems >= 0 && count < node.MinItems) {
            valid = false;
            if (!report || !cx.Add(ip, Child(sp, "minItems"), "minItems", "Array must have at least " + node.MinItems.ToString(CultureInfo.InvariantCulture) + " items; it has " + count.ToString(CultureInfo.InvariantCulture) + ".", instance)) return false;
        }
        if (node.UniqueItems && count > 1 && FindDuplicate(items, out int first, out int second)) {
            valid = false;
            if (!report || !cx.Add(ip, Child(sp, "uniqueItems"), "uniqueItems", "Items " + first.ToString(CultureInfo.InvariantCulture) + " and " + second.ToString(CultureInfo.InvariantCulture) + " are equal; items must be unique.", instance)) return false;
        }
        int prefix = 0;
        if (node.PrefixItems is not null) {
            prefix = Math.Min(count, node.PrefixItems.Length);
            for (int i = 0; i < prefix; i++) {
                if (!Evaluate(node.PrefixItems[i], items[i], Child(ip, i), Child(Child(sp, node.PrefixItemsKeyword), i), report, null, scope, depth + 1, cx)) {
                    valid = false;
                    if (!report || cx.Aborted) return false;
                }
            }
            if (prefix > local.PrefixItems) local.PrefixItems = prefix;
        }
        if (node.Items is not null && count > prefix) {
            for (int i = prefix; i < count; i++) {
                if (node.Items.Boolean == false) {
                    valid = false;
                    string message = prefix == 0 ? "Array must be empty." : "Array must have at most " + prefix.ToString(CultureInfo.InvariantCulture) + " items; item " + i.ToString(CultureInfo.InvariantCulture) + " is not allowed.";
                    if (!report || !cx.Add(Child(ip, i), Child(sp, node.ItemsKeyword), node.ItemsKeyword, message, items[i])) return false;
                    break;
                }
                if (!Evaluate(node.Items, items[i], Child(ip, i), Child(sp, node.ItemsKeyword), report, null, scope, depth + 1, cx)) {
                    valid = false;
                    if (!report || cx.Aborted) return false;
                }
            }
            local.AllItems = true;
        }
        if (node.Contains is not null) {
            int matches = 0;
            for (int i = 0; i < count; i++) {
                if (Evaluate(node.Contains, items[i], Child(ip, i), Child(sp, "contains"), false, null, scope, depth + 1, cx)) {
                    matches++;
                    (local.ContainedItems ??= new HashSet<int>()).Add(i);
                }
                if (cx.Aborted) return false;
            }
            long min = node.MinContains >= 0 ? node.MinContains : 1;
            if (matches < min) {
                valid = false;
                string keyword = node.MinContains >= 0 ? "minContains" : "contains";
                string message = min == 1 ? "Array must contain an item that matches 'contains'." : "Array must contain at least " + min.ToString(CultureInfo.InvariantCulture) + " items that match 'contains'; it has " + matches.ToString(CultureInfo.InvariantCulture) + ".";
                if (!report || !cx.Add(ip, Child(sp, keyword), keyword, message, instance)) return false;
            }
            if (node.MaxContains >= 0 && matches > node.MaxContains) {
                valid = false;
                if (!report || !cx.Add(ip, Child(sp, "maxContains"), "maxContains", "Array must contain at most " + node.MaxContains.ToString(CultureInfo.InvariantCulture) + " items that match 'contains'; it has " + matches.ToString(CultureInfo.InvariantCulture) + ".", instance)) return false;
            }
        }
        return valid;
    }

    private static bool FindDuplicate(IReadOnlyList<JsonData> items, out int first, out int second) {
        var buckets = new Dictionary<int, List<int>>();
        for (int i = 0; i < items.Count; i++) {
            int hash = JsonData.DeepHash(items[i]);
            if (buckets.TryGetValue(hash, out List<int>? candidates)) {
                foreach (int candidate in candidates) {
                    if (JsonData.DeepEquals(items[candidate], items[i])) {
                        first = candidate;
                        second = i;
                        return true;
                    }
                }
                candidates.Add(i);
            } else {
                buckets.Add(hash, new List<int> { i });
            }
        }
        first = second = -1;
        return false;
    }

    private bool EvaluateObject(JsonSchemaNode node, JsonData instance, JsonPointerPath? ip, JsonPointerPath? sp, bool report,
        JsonSchemaAnnotations local, ScopeEntry scope, int depth, EvaluationContext cx) {
        bool valid = true;
        IReadOnlyList<KeyValuePair<string, JsonData>> properties = instance.Properties;
        int count = properties.Count;
        if (node.MaxProperties >= 0 && count > node.MaxProperties) {
            valid = false;
            if (!report || !cx.Add(ip, Child(sp, "maxProperties"), "maxProperties", "Object must have at most " + node.MaxProperties.ToString(CultureInfo.InvariantCulture) + " properties; it has " + count.ToString(CultureInfo.InvariantCulture) + ".", instance)) return false;
        }
        if (node.MinProperties >= 0 && count < node.MinProperties) {
            valid = false;
            if (!report || !cx.Add(ip, Child(sp, "minProperties"), "minProperties", "Object must have at least " + node.MinProperties.ToString(CultureInfo.InvariantCulture) + " properties; it has " + count.ToString(CultureInfo.InvariantCulture) + ".", instance)) return false;
        }
        if (node.Required is not null) {
            foreach (string name in node.Required) {
                if (instance.IndexOf(name) >= 0) continue;
                valid = false;
                if (!report || !cx.Add(ip, Child(sp, "required"), "required", "Required property '" + name + "' is missing.", instance)) return false;
            }
        }
        if (node.DependentRequired is not null) {
            foreach (KeyValuePair<string, string[]> dependency in node.DependentRequired) {
                if (instance.IndexOf(dependency.Key) < 0) continue;
                foreach (string name in dependency.Value) {
                    if (instance.IndexOf(name) >= 0) continue;
                    valid = false;
                    if (!report || !cx.Add(ip, Child(Child(sp, node.DependentRequiredKeyword), dependency.Key), node.DependentRequiredKeyword, "Property '" + name + "' is required when '" + dependency.Key + "' is present.", instance)) return false;
                }
            }
        }
        if (node.Properties is not null) {
            foreach (KeyValuePair<string, JsonSchemaNode> entry in node.Properties) {
                if (!instance.TryGetProperty(entry.Key, out JsonData value)) continue;
                local.AddProperty(entry.Key);
                if (entry.Value.Boolean == false) {
                    valid = false;
                    if (!report || !cx.Add(Child(ip, entry.Key), Child(Child(sp, "properties"), entry.Key), "properties", "Property '" + entry.Key + "' is not allowed.", value)) return false;
                } else if (!Evaluate(entry.Value, value, Child(ip, entry.Key), Child(Child(sp, "properties"), entry.Key), report, null, scope, depth + 1, cx)) {
                    valid = false;
                    if (!report || cx.Aborted) return false;
                }
            }
        }
        if (node.PatternProperties is not null || node.AdditionalProperties is not null) {
            foreach (KeyValuePair<string, JsonData> property in properties) {
                bool matched = false;
                if (node.PatternProperties is not null) {
                    foreach (JsonSchemaNode.PatternSchema pattern in node.PatternProperties) {
                        bool? isMatch = TryMatch(pattern.Regex, property.Key);
                        if (isMatch is null) {
                            cx.Abort(Child(ip, property.Key), Child(Child(sp, "patternProperties"), pattern.Pattern), "patternProperties", "Matching the pattern '" + pattern.Pattern + "' timed out (JsonSchemaOptions.RegexTimeout).", property.Value);
                            return false;
                        }
                        if (isMatch == false) continue;
                        matched = true;
                        local.AddProperty(property.Key);
                        if (!Evaluate(pattern.Schema, property.Value, Child(ip, property.Key), Child(Child(sp, "patternProperties"), pattern.Pattern), report, null, scope, depth + 1, cx)) {
                            valid = false;
                            if (!report || cx.Aborted) return false;
                        }
                    }
                }
                if (matched || node.AdditionalProperties is null || HasProperty(node, property.Key)) continue;
                local.AddProperty(property.Key);
                if (node.AdditionalProperties.Boolean == false) {
                    valid = false;
                    if (!report || !cx.Add(Child(ip, property.Key), Child(sp, "additionalProperties"), "additionalProperties", "Property '" + property.Key + "' is not allowed.", property.Value)) return false;
                } else if (!Evaluate(node.AdditionalProperties, property.Value, Child(ip, property.Key), Child(sp, "additionalProperties"), report, null, scope, depth + 1, cx)) {
                    valid = false;
                    if (!report || cx.Aborted) return false;
                }
            }
        }
        if (node.PropertyNames is not null) {
            foreach (KeyValuePair<string, JsonData> property in properties) {
                JsonData name = JsonData.CreateString(property.Key, property.Value.Location);
                if (Evaluate(node.PropertyNames, name, Child(ip, property.Key), Child(sp, "propertyNames"), false, null, scope, depth + 1, cx)) continue;
                if (cx.Aborted) return false;
                valid = false;
                if (!report || !cx.Add(Child(ip, property.Key), Child(sp, "propertyNames"), "propertyNames", "Property name '" + property.Key + "' does not match 'propertyNames'.", property.Value)) return false;
            }
        }
        if (node.DependentSchemas is not null) {
            foreach (KeyValuePair<string, JsonSchemaNode> dependency in node.DependentSchemas) {
                if (instance.IndexOf(dependency.Key) < 0) continue;
                if (!Evaluate(dependency.Value, instance, ip, Child(Child(sp, node.DependentSchemasKeyword), dependency.Key), report, local, scope, depth + 1, cx)) {
                    valid = false;
                    if (!report || cx.Aborted) return false;
                }
            }
        }
        return valid;
    }

    private static bool HasProperty(JsonSchemaNode node, string name) {
        if (node.Properties is null) return false;
        if (node.PropertyLookup is not null) return node.PropertyLookup.ContainsKey(name);
        foreach (KeyValuePair<string, JsonSchemaNode> entry in node.Properties) {
            if (string.Equals(entry.Key, name, StringComparison.Ordinal)) return true;
        }
        return false;
    }

    #endregion

    #region Applicators

    private bool EvaluateApplicators(JsonSchemaNode node, JsonData instance, JsonPointerPath? ip, JsonPointerPath? sp, bool report,
        JsonSchemaAnnotations? local, ScopeEntry scope, int depth, EvaluationContext cx) {
        bool valid = true;
        if (node.AllOf is not null) {
            for (int i = 0; i < node.AllOf.Length; i++) {
                if (!Evaluate(node.AllOf[i], instance, ip, Child(Child(sp, "allOf"), i), report, local, scope, depth + 1, cx)) {
                    valid = false;
                    if (!report || cx.Aborted) return false;
                }
            }
        }
        if (node.AnyOf is not null) {
            bool any = false;
            for (int i = 0; i < node.AnyOf.Length; i++) {
                // Every matching branch contributes annotations, so keep going after a match when they are collected.
                JsonSchemaAnnotations? branch = local is null ? null : new JsonSchemaAnnotations();
                if (Evaluate(node.AnyOf[i], instance, ip, Child(Child(sp, "anyOf"), i), false, branch, scope, depth + 1, cx)) {
                    any = true;
                    if (branch is null) break;
                    local!.Merge(branch);
                }
                if (cx.Aborted) return false;
            }
            if (!any) {
                valid = false;
                if (!report || !cx.Add(ip, Child(sp, "anyOf"), "anyOf", "Value does not match any of the " + node.AnyOf.Length.ToString(CultureInfo.InvariantCulture) + " schemas in 'anyOf'.", instance)) return false;
            }
        }
        if (node.OneOf is not null) {
            int matched = -1, second = -1;
            JsonSchemaAnnotations? matchedAnnotations = null;
            for (int i = 0; i < node.OneOf.Length; i++) {
                JsonSchemaAnnotations? branch = local is null ? null : new JsonSchemaAnnotations();
                if (Evaluate(node.OneOf[i], instance, ip, Child(Child(sp, "oneOf"), i), false, branch, scope, depth + 1, cx)) {
                    if (matched < 0) {
                        matched = i;
                        matchedAnnotations = branch;
                    } else {
                        second = i;
                        break;
                    }
                }
                if (cx.Aborted) return false;
            }
            if (matched < 0 || second >= 0) {
                valid = false;
                string message = matched < 0
                    ? "Value does not match any of the " + node.OneOf.Length.ToString(CultureInfo.InvariantCulture) + " schemas in 'oneOf'."
                    : "Value matches more than one schema in 'oneOf' (" + matched.ToString(CultureInfo.InvariantCulture) + " and " + second.ToString(CultureInfo.InvariantCulture) + ").";
                if (!report || !cx.Add(ip, Child(sp, "oneOf"), "oneOf", message, instance)) return false;
            } else if (matchedAnnotations is not null) {
                local!.Merge(matchedAnnotations);
            }
        }
        if (node.Not is not null) {
            bool matches = Evaluate(node.Not, instance, ip, Child(sp, "not"), false, null, scope, depth + 1, cx);
            if (cx.Aborted) return false;
            if (matches) {
                valid = false;
                if (!report || !cx.Add(ip, Child(sp, "not"), "not", "Value must not match the schema in 'not'.", instance)) return false;
            }
        }
        if (node.If is not null && (node.Then is not null || node.Else is not null || local is not null)) {
            JsonSchemaAnnotations? condition = local is null ? null : new JsonSchemaAnnotations();
            bool matches = Evaluate(node.If, instance, ip, Child(sp, "if"), false, condition, scope, depth + 1, cx);
            if (cx.Aborted) return false;
            if (matches && condition is not null) local!.Merge(condition);
            JsonSchemaNode? branch = matches ? node.Then : node.Else;
            if (branch is not null && !Evaluate(branch, instance, ip, Child(sp, matches ? "then" : "else"), report, local, scope, depth + 1, cx)) {
                valid = false;
                if (!report || cx.Aborted) return false;
            }
        }
        return valid;
    }

    #endregion

    #region Messages

    private static string DescribeTypes(JsonSchemaTypes types) {
        var names = new List<string>();
        if ((types & JsonSchemaTypes.Object) != 0) names.Add("object");
        if ((types & JsonSchemaTypes.Array) != 0) names.Add("array");
        if ((types & JsonSchemaTypes.String) != 0) names.Add("string");
        if ((types & JsonSchemaTypes.Number) != 0) names.Add("number");
        if ((types & JsonSchemaTypes.Integer) != 0) names.Add("integer");
        if ((types & JsonSchemaTypes.Boolean) != 0) names.Add("boolean");
        if ((types & JsonSchemaTypes.Null) != 0) names.Add("null");
        if (names.Count == 1) return names[0];
        return string.Join(", ", names.Take(names.Count - 1)) + " or " + names[names.Count - 1];
    }

    private static string DescribeValueType(JsonData value) => value.Kind switch {
        JsonDataKind.Null => "null",
        JsonDataKind.Boolean => "boolean",
        JsonDataKind.Number => value.IsInteger ? "integer" : "number",
        JsonDataKind.String => "string",
        JsonDataKind.Array => "array",
        _ => "object"
    };

    private static string ListValues(JsonData[] values) {
        var text = new StringBuilder();
        int shown = Math.Min(values.Length, MaxListedValues);
        for (int i = 0; i < shown; i++) {
            if (i > 0) text.Append(i == values.Length - 1 ? " or " : ", ");
            text.Append(Preview(values[i]));
        }
        if (shown < values.Length) text.Append(", … (").Append(values.Length.ToString(CultureInfo.InvariantCulture)).Append(" values)");
        return text.ToString();
    }

    private static string Preview(JsonData value) {
        string text = value.ToString();
        return text.Length <= 80 ? text : text.Substring(0, 77) + "...";
    }

    #endregion
}
