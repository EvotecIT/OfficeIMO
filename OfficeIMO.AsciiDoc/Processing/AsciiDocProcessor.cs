namespace OfficeIMO.AsciiDoc;

/// <summary>Opt-in, bounded preprocessor for attributes, conditionals, and includes.</summary>
public static class AsciiDocProcessor {
    /// <summary>Processes a source string without changing the lossless original document.</summary>
    public static AsciiDocProcessingResult Process(string source, AsciiDocProcessorOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        options ??= new AsciiDocProcessorOptions();
        ValidateOptions(options);

        AsciiDocDocument sourceDocument = AsciiDocDocument.Parse(source).Document;
        var state = new PreprocessorState(options);
        string processed = state.ProcessSource(source, options.SourceName, 0);
        AsciiDocDocument document = AsciiDocDocument.Parse(processed).Document;
        return new AsciiDocProcessingResult(sourceDocument, document, processed, state.Attributes, state.Diagnostics);
    }

    private static void ValidateOptions(AsciiDocProcessorOptions options) {
        if (options.MaximumIncludeDepth < 0) throw new ArgumentOutOfRangeException(nameof(options), "MaximumIncludeDepth cannot be negative.");
        if (options.MaximumIncludeCount < 0) throw new ArgumentOutOfRangeException(nameof(options), "MaximumIncludeCount cannot be negative.");
        if (options.MaximumIncludedCharacters < 0) throw new ArgumentOutOfRangeException(nameof(options), "MaximumIncludedCharacters cannot be negative.");
        if (options.MaximumOutputLength < 1) throw new ArgumentOutOfRangeException(nameof(options), "MaximumOutputLength must be positive.");
        if (options.MaximumExtensionInvocations < 0) throw new ArgumentOutOfRangeException(nameof(options), "MaximumExtensionInvocations cannot be negative.");
    }

    private sealed class PreprocessorState {
        private readonly AsciiDocProcessorOptions _options;
        private readonly Dictionary<string, string> _attributes;
        private readonly List<AsciiDocProcessingDiagnostic> _diagnostics = new List<AsciiDocProcessingDiagnostic>();
        private readonly HashSet<string> _activeSources;
        private int _includeCount;
        private int _includedCharacters;
        private int _extensionInvocations;

        internal PreprocessorState(AsciiDocProcessorOptions options) {
            _options = options;
            _attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (options.Attributes != null) {
                foreach (KeyValuePair<string, string> value in options.Attributes) _attributes[value.Key] = value.Value;
            }
            _activeSources = new HashSet<string>(Path.DirectorySeparatorChar == '\\' ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
        }

        internal AsciiDocDocumentAttributes Attributes => new AsciiDocDocumentAttributes(_attributes);

        internal IReadOnlyList<AsciiDocProcessingDiagnostic> Diagnostics => _diagnostics;

        internal string ProcessSource(string source, string? sourceName, int includeDepth) {
            bool addedSource = sourceName != null && _activeSources.Add(sourceName);
            var output = new StringBuilder(source.Length);
            var conditions = new Stack<bool>();
            bool active = true;
            IReadOnlyList<AsciiDocSourceLine> lines = AsciiDocLineReader.Read(source);

            for (int index = 0; index < lines.Count; index++) {
                AsciiDocSourceLine line = lines[index];
                string content = line.Content;
                if (content.Length > 0 && content[0] != '\\' &&
                    AsciiDocLineClassifier.TryParseBlockMacro(content, out AsciiDocLineClassifier.BlockMacroParts directive)) {
                    if (IsConditional(directive.Name)) {
                        active = ProcessConditional(directive, conditions, active, output, line, sourceName, index + 1);
                        continue;
                    }
                    if (string.Equals(directive.Name, "endif", StringComparison.Ordinal)) {
                        if (conditions.Count == 0) {
                            Report("ADOCPROC005", AsciiDocDiagnosticSeverity.Error, "Unexpected endif directive.", sourceName, index + 1);
                        } else {
                            conditions.Pop();
                            active = conditions.Count == 0 || conditions.All(static condition => condition);
                        }
                        continue;
                    }
                    if (active && string.Equals(directive.Name, "include", StringComparison.Ordinal)) {
                        ProcessInclude(directive, output, line, sourceName, index + 1, includeDepth);
                        continue;
                    }
                    if (active && _options.Extensions != null &&
                        _options.Extensions.TryGetDirective(directive.Name, out IAsciiDocDirectiveProcessor extension)) {
                        ProcessExtension(extension, directive, output, line, sourceName, index + 1);
                        continue;
                    }
                }

                if (!active) continue;
                if (AsciiDocLineClassifier.TryParseAttribute(content, out AsciiDocLineClassifier.AttributeParts attribute)) {
                    if (attribute.IsUnset) _attributes.Remove(attribute.Name);
                    else _attributes[attribute.Name] = Substitute(attribute.Value, sourceName, index + 1);
                }
                output.Append(line.FullText);
                EnforceOutput(output);
            }

            if (conditions.Count > 0) {
                Report("ADOCPROC006", AsciiDocDiagnosticSeverity.Error, "Conditional directive is not terminated by endif.", sourceName, lines.Count);
            }
            if (addedSource && sourceName != null) _activeSources.Remove(sourceName);
            return output.ToString();
        }

        private bool ProcessConditional(
            AsciiDocLineClassifier.BlockMacroParts directive,
            Stack<bool> conditions,
            bool currentlyActive,
            StringBuilder output,
            AsciiDocSourceLine line,
            string? sourceName,
            int lineNumber) {
            bool result;
            if (string.Equals(directive.Name, "ifeval", StringComparison.Ordinal)) {
                result = EvaluateExpression(Substitute(directive.AttributeList, sourceName, lineNumber), sourceName, lineNumber);
                conditions.Push(currentlyActive && result);
                return conditions.All(static condition => condition);
            }

            result = EvaluateAttributeCondition(directive.Target);
            if (string.Equals(directive.Name, "ifndef", StringComparison.Ordinal)) result = !result;
            if (directive.AttributeList.Length == 0) {
                conditions.Push(currentlyActive && result);
                return conditions.All(static condition => condition);
            }
            if (currentlyActive && result) {
                output.Append(Substitute(directive.AttributeList, sourceName, lineNumber));
                output.Append(line.LineEnding);
                EnforceOutput(output);
            }
            return currentlyActive;
        }

        private bool EvaluateAttributeCondition(string expression) {
            if (expression.IndexOf('+') >= 0) {
                string[] all = expression.Split(new[] { '+' }, StringSplitOptions.RemoveEmptyEntries);
                return all.Length > 0 && all.All(name => _attributes.ContainsKey(name.Trim()));
            }
            if (expression.IndexOf(',') >= 0) {
                string[] any = expression.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                return any.Any(name => _attributes.ContainsKey(name.Trim()));
            }
            return _attributes.ContainsKey(expression.Trim());
        }

        private bool EvaluateExpression(string expression, string? sourceName, int lineNumber) {
            string[] operators = { "==", "!=", ">=", "<=", ">", "<" };
            for (int index = 0; index < operators.Length; index++) {
                string operation = operators[index];
                int position = expression.IndexOf(operation, StringComparison.Ordinal);
                if (position < 0) continue;
                string left = Unquote(expression.Substring(0, position).Trim());
                string right = Unquote(expression.Substring(position + operation.Length).Trim());
                if (decimal.TryParse(left, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out decimal leftNumber) &&
                    decimal.TryParse(right, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out decimal rightNumber)) {
                    int comparison = leftNumber.CompareTo(rightNumber);
                    return Compare(comparison, operation);
                }
                return Compare(string.Compare(left, right, StringComparison.Ordinal), operation);
            }
            Report("ADOCPROC007", AsciiDocDiagnosticSeverity.Error, "Unsupported ifeval expression.", sourceName, lineNumber);
            return false;
        }

        private void ProcessInclude(
            AsciiDocLineClassifier.BlockMacroParts directive,
            StringBuilder output,
            AsciiDocSourceLine line,
            string? sourceName,
            int lineNumber,
            int includeDepth) {
            AsciiDocElementAttributes elementAttributes = AsciiDocAttributeListParser.Parse(directive.AttributeList);
            bool optional = elementAttributes.Options.Any(static option => string.Equals(option, "optional", StringComparison.OrdinalIgnoreCase)) ||
                string.Equals(elementAttributes.GetNamedValue("opts"), "optional", StringComparison.OrdinalIgnoreCase);
            string target = Substitute(directive.Target, sourceName, lineNumber);
            if (_options.IncludeResolver == null) {
                if (!optional) Report("ADOCPROC001", AsciiDocDiagnosticSeverity.Warning, "Include resolution is disabled.", sourceName, lineNumber);
                output.Append(line.FullText);
                return;
            }
            if (includeDepth >= _options.MaximumIncludeDepth || _includeCount >= _options.MaximumIncludeCount) {
                Report("ADOCPROC002", AsciiDocDiagnosticSeverity.Error, "Include limit exceeded.", sourceName, lineNumber);
                output.Append(line.FullText);
                return;
            }

            var request = new AsciiDocIncludeRequest(target, sourceName, includeDepth + 1, Attributes);
            AsciiDocIncludeResult? resolved = _options.IncludeResolver.Resolve(request);
            if (resolved == null) {
                if (!optional) {
                    Report("ADOCPROC003", AsciiDocDiagnosticSeverity.Warning, "Include target was unavailable or denied: " + target, sourceName, lineNumber);
                    output.Append(line.FullText);
                }
                return;
            }
            if (resolved.SourceName != null && _activeSources.Contains(resolved.SourceName)) {
                Report("ADOCPROC004", AsciiDocDiagnosticSeverity.Error, "Include cycle detected: " + resolved.SourceName, sourceName, lineNumber);
                output.Append(line.FullText);
                return;
            }

            _includeCount++;
            _includedCharacters += resolved.Content.Length;
            if (_includedCharacters > _options.MaximumIncludedCharacters) throw new InvalidDataException("Includes exceed MaximumIncludedCharacters.");
            string selected = AsciiDocIncludeSelector.Apply(resolved.Content, elementAttributes);
            string expanded = ProcessSource(selected, resolved.SourceName, includeDepth + 1);
            output.Append(expanded);
            if (line.LineEnding.Length > 0 && expanded.Length > 0 && !AsciiDocText.EndsWithLineEnding(expanded)) output.Append(line.LineEnding);
            EnforceOutput(output);
        }

        private void ProcessExtension(
            IAsciiDocDirectiveProcessor extension,
            AsciiDocLineClassifier.BlockMacroParts directive,
            StringBuilder output,
            AsciiDocSourceLine line,
            string? sourceName,
            int lineNumber) {
            _extensionInvocations++;
            if (_extensionInvocations > _options.MaximumExtensionInvocations) {
                Report("ADOCPROC008", AsciiDocDiagnosticSeverity.Error, "Extension invocation limit exceeded.", sourceName, lineNumber);
                output.Append(line.FullText);
                return;
            }
            var context = new AsciiDocDirectiveContext(
                directive.Name,
                directive.Target,
                directive.AttributeList,
                line.FullText,
                sourceName,
                lineNumber,
                Attributes);
            AsciiDocDirectiveResult result = extension.Process(context)
                ?? throw new InvalidOperationException("AsciiDoc directive processors must return a result.");
            if (result.PreserveOriginal) output.Append(line.FullText);
            else output.Append(result.Replacement);
            EnforceOutput(output);
        }

        private string Substitute(string value, string? sourceName, int lineNumber) {
            var substitutionOptions = new AsciiDocAttributeSubstitutionOptions {
                UndefinedAttributeBehavior = _options.UndefinedAttributeBehavior,
                MaximumOutputLength = _options.MaximumOutputLength
            };
            AsciiDocAttributeSubstitutionResult result = AsciiDocAttributeSubstitutor.Substitute(value, Attributes, substitutionOptions);
            for (int index = 0; index < result.Diagnostics.Count; index++) {
                AsciiDocEvaluationDiagnostic diagnostic = result.Diagnostics[index];
                Report(diagnostic.Code, diagnostic.Severity, diagnostic.Message, sourceName, lineNumber);
            }
            return result.Value;
        }

        private void EnforceOutput(StringBuilder output) {
            if (output.Length > _options.MaximumOutputLength) throw new InvalidDataException("Processed AsciiDoc exceeds MaximumOutputLength.");
        }

        private void Report(string code, AsciiDocDiagnosticSeverity severity, string message, string? sourceName, int line) =>
            _diagnostics.Add(new AsciiDocProcessingDiagnostic(code, severity, message, sourceName, line));

        private static bool IsConditional(string name) =>
            string.Equals(name, "ifdef", StringComparison.Ordinal) ||
            string.Equals(name, "ifndef", StringComparison.Ordinal) ||
            string.Equals(name, "ifeval", StringComparison.Ordinal);

        private static string Unquote(string value) {
            if (value.Length >= 2 && ((value[0] == '"' && value[value.Length - 1] == '"') ||
                                      (value[0] == '\'' && value[value.Length - 1] == '\''))) {
                return value.Substring(1, value.Length - 2);
            }
            return value;
        }

        private static bool Compare(int comparison, string operation) {
            switch (operation) {
                case "==": return comparison == 0;
                case "!=": return comparison != 0;
                case ">=": return comparison >= 0;
                case "<=": return comparison <= 0;
                case ">": return comparison > 0;
                case "<": return comparison < 0;
                default: return false;
            }
        }
    }
}
