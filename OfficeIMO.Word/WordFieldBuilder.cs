namespace OfficeIMO.Word {
    /// <summary>
    /// Provides a fluent builder for constructing field codes including nested fields.
    /// </summary>
    public class WordFieldBuilder : WordFieldCode {
        private readonly List<object> _parts = new();

        /// <summary>
        /// Gets or sets the format switch applied to the field.
        /// </summary>
        public WordFieldFormat? Format { get; private set; }

        /// <summary>
        /// Gets or sets the custom date or time format.
        /// </summary>
        public string CustomFormat { get; private set; }

        internal override WordFieldType FieldType { get; }

        /// <summary>
        /// Initializes a new instance for the specified <see cref="WordFieldType"/>.
        /// </summary>
        public WordFieldBuilder(WordFieldType fieldType) {
            FieldType = fieldType;
        }

        /// <summary>
        /// Adds an instruction to the field code.
        /// </summary>
        /// <param name="instruction">Instruction to append.</param>
        /// <param name="quoted">When true, the instruction is wrapped in quotes.</param>
        /// <returns>The current <see cref="WordFieldBuilder"/>.</returns>
        public WordFieldBuilder AddInstruction(string instruction, bool quoted = false) {
            if (!string.IsNullOrWhiteSpace(instruction)) {
                _parts.Add(quoted ? $"\"{instruction}\"" : instruction);
            }
            return this;
        }

        /// <summary>
        /// Adds a nested field as an instruction.
        /// </summary>
        /// <param name="field">Nested field builder.</param>
        /// <returns>The current <see cref="WordFieldBuilder"/>.</returns>
        public WordFieldBuilder AddInstruction(WordFieldBuilder field) {
            if (field != null) {
                _parts.Add(field);
            }
            return this;
        }

        /// <summary>
        /// Adds a switch to the field code.
        /// </summary>
        /// <param name="value">Switch string including leading backslash.</param>
        /// <returns>The current <see cref="WordFieldBuilder"/>.</returns>
        public WordFieldBuilder AddSwitch(string value) {
            if (!string.IsNullOrWhiteSpace(value)) {
                _parts.Add(value.Trim());
            }
            return this;
        }

        /// <summary>
        /// Applies a format switch to the field code.
        /// </summary>
        /// <param name="format">Format switch to apply.</param>
        /// <returns>The current <see cref="WordFieldBuilder"/>.</returns>
        public WordFieldBuilder SetFormat(WordFieldFormat format) {
            Format = format;
            return this;
        }

        /// <summary>
        /// Sets a custom date or time format.
        /// </summary>
        /// <param name="format">Custom format string.</param>
        /// <returns>The current <see cref="WordFieldBuilder"/>.</returns>
        public WordFieldBuilder SetCustomFormat(string format) {
            CustomFormat = format;
            return this;
        }

        internal override List<string> GetParameters() {
            var parameters = new List<string>();
            foreach (var part in _parts) {
                if (part is WordFieldBuilder builder) {
                    parameters.Add($"{{ {builder.Build()} }}");
                } else {
                    parameters.Add(part.ToString());
                }
            }
            return parameters;
        }

        /// <summary>
        /// Builds the string representation of the field code.
        /// </summary>
        public string Build() {
            var sb = new System.Text.StringBuilder();
            sb.Append(' ').Append(FieldType.ToString().ToUpper()).Append(' ');
            var parameters = GetParameters();
            if (parameters.Count > 0) {
                sb.Append(string.Join(" ", parameters)).Append(' ');
            }
            if (Format != null) {
                sb.Append("\\* ").Append(Format).Append(' ');
            }
            if (!string.IsNullOrWhiteSpace(CustomFormat)) {
                sb.Append("\\@ \"").Append(CustomFormat).Append("\" ");
            }
            sb.Append("\\* MERGEFORMAT ");
            return sb.ToString();
        }
    }
}
