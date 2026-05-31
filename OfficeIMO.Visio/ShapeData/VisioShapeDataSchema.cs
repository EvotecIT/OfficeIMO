using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Reusable Shape Data schema for applying consistent Visio Prop rows to shapes and connectors.
    /// </summary>
    public sealed class VisioShapeDataSchema {
        private readonly List<VisioShapeDataField> _fields = new();

        /// <summary>
        /// Creates an empty Shape Data schema.
        /// </summary>
        public static VisioShapeDataSchema Create() {
            return new VisioShapeDataSchema();
        }

        /// <summary>
        /// Fields included in the schema.
        /// </summary>
        public IReadOnlyList<VisioShapeDataField> Fields => _fields;

        /// <summary>
        /// Adds a field to the schema.
        /// </summary>
        /// <param name="name">Shape Data row name.</param>
        /// <param name="label">Optional label shown in Visio's Shape Data window.</param>
        /// <param name="type">Optional Shape Data type.</param>
        /// <param name="defaultValue">Default value used when applying the schema to missing or overwritten fields.</param>
        /// <param name="prompt">Optional help prompt.</param>
        /// <param name="format">Optional format picture or list values.</param>
        /// <param name="sortKey">Optional sort key used by Visio's Shape Data window.</param>
        /// <param name="required">Whether validation should require a non-empty value.</param>
        /// <param name="invisible">Whether the row is hidden in Visio's Shape Data window.</param>
        /// <param name="verify">Whether Visio should ask for the value when the shape is created, copied, or duplicated.</param>
        /// <param name="dataLinked">Whether this row is linked to an external data recordset.</param>
        /// <param name="allowedValues">Optional allowed value list used by validation and list formats.</param>
        public VisioShapeDataSchema Field(
            string name,
            string? label = null,
            VisioShapeDataType? type = null,
            string? defaultValue = null,
            string? prompt = null,
            string? format = null,
            string? sortKey = null,
            bool required = false,
            bool? invisible = null,
            bool? verify = null,
            bool? dataLinked = null,
            IEnumerable<string>? allowedValues = null) {
            Add(new VisioShapeDataField(name) {
                Label = label,
                Type = type,
                DefaultValue = defaultValue,
                Prompt = prompt,
                Format = format,
                SortKey = sortKey,
                Required = required,
                Invisible = invisible,
                Verify = verify,
                DataLinked = dataLinked,
                AllowedValues = allowedValues?.ToArray() ?? Array.Empty<string>()
            });

            return this;
        }

        /// <summary>
        /// Adds a preconfigured field to the schema.
        /// </summary>
        /// <param name="field">Field to add.</param>
        public VisioShapeDataSchema Add(VisioShapeDataField field) {
            if (field == null) {
                throw new ArgumentNullException(nameof(field));
            }

            if (_fields.Any(existing => string.Equals(existing.Name, field.Name, StringComparison.OrdinalIgnoreCase))) {
                throw new InvalidOperationException("The schema already contains a Shape Data field named '" + field.Name + "'.");
            }

            _fields.Add(field);
            return this;
        }

        /// <summary>
        /// Applies the schema to a shape.
        /// </summary>
        /// <param name="shape">Shape to update.</param>
        /// <param name="overwriteValues">Whether schema defaults should replace existing values.</param>
        public VisioShape ApplyTo(VisioShape shape, bool overwriteValues = false) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            foreach (VisioShapeDataField field in _fields) {
                ApplyField(shape, field, overwriteValues);
            }

            return shape;
        }

        /// <summary>
        /// Applies the schema to a connector.
        /// </summary>
        /// <param name="connector">Connector to update.</param>
        /// <param name="overwriteValues">Whether schema defaults should replace existing values.</param>
        public VisioConnector ApplyTo(VisioConnector connector, bool overwriteValues = false) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            foreach (VisioShapeDataField field in _fields) {
                ApplyField(connector, field, overwriteValues);
            }

            return connector;
        }

        /// <summary>
        /// Applies the schema to a set of shapes.
        /// </summary>
        /// <param name="shapes">Shapes to update.</param>
        /// <param name="overwriteValues">Whether schema defaults should replace existing values.</param>
        public IReadOnlyList<VisioShape> ApplyTo(IEnumerable<VisioShape> shapes, bool overwriteValues = false) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            List<VisioShape> updated = new();
            foreach (VisioShape shape in shapes) {
                updated.Add(ApplyTo(shape, overwriteValues));
            }

            return updated;
        }

        /// <summary>
        /// Applies the schema to a set of connectors.
        /// </summary>
        /// <param name="connectors">Connectors to update.</param>
        /// <param name="overwriteValues">Whether schema defaults should replace existing values.</param>
        public IReadOnlyList<VisioConnector> ApplyToConnectors(IEnumerable<VisioConnector> connectors, bool overwriteValues = false) {
            if (connectors == null) {
                throw new ArgumentNullException(nameof(connectors));
            }

            List<VisioConnector> updated = new();
            foreach (VisioConnector connector in connectors) {
                updated.Add(ApplyTo(connector, overwriteValues));
            }

            return updated;
        }

        /// <summary>
        /// Validates a shape against the schema.
        /// </summary>
        /// <param name="shape">Shape to validate.</param>
        public IReadOnlyList<VisioShapeDataSchemaIssue> Validate(VisioShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            return Validate(
                "Shape",
                shape.Id,
                shape.Text,
                name => shape.FindShapeData(name),
                name => shape.GetShapeDataValue(name));
        }

        /// <summary>
        /// Validates a connector against the schema.
        /// </summary>
        /// <param name="connector">Connector to validate.</param>
        public IReadOnlyList<VisioShapeDataSchemaIssue> Validate(VisioConnector connector) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            return Validate(
                "Connector",
                connector.Id,
                connector.Label,
                name => connector.FindShapeData(name),
                name => connector.GetShapeDataValue(name));
        }

        private static void ApplyField(VisioShape shape, VisioShapeDataField field, bool overwriteValues) {
            VisioShapeDataRow? existing = shape.FindShapeData(field.Name);
            string? value = ResolveValue(existing, shape.GetShapeDataValue(field.Name), field, overwriteValues);
            VisioShapeDataRow row = shape.SetShapeData(field.Name, value, field.Label, field.Type, field.Prompt, field.GetEffectiveFormat());
            ApplyMetadata(row, field);
        }

        private static void ApplyField(VisioConnector connector, VisioShapeDataField field, bool overwriteValues) {
            VisioShapeDataRow? existing = connector.FindShapeData(field.Name);
            string? value = ResolveValue(existing, connector.GetShapeDataValue(field.Name), field, overwriteValues);
            VisioShapeDataRow row = connector.SetShapeData(field.Name, value, field.Label, field.Type, field.Prompt, field.GetEffectiveFormat());
            ApplyMetadata(row, field);
        }

        private static string? ResolveValue(VisioShapeDataRow? existing, string? currentValue, VisioShapeDataField field, bool overwriteValues) {
            if (overwriteValues || existing == null) {
                return field.DefaultValue;
            }

            return currentValue;
        }

        private static void ApplyMetadata(VisioShapeDataRow row, VisioShapeDataField field) {
            if (field.SortKey != null) row.SortKey = field.SortKey;
            if (field.Invisible.HasValue) row.Invisible = field.Invisible.Value;
            if (field.Verify.HasValue) row.Verify = field.Verify.Value;
            if (field.DataLinked.HasValue) row.DataLinked = field.DataLinked.Value;
            if (field.Calendar != null) row.Calendar = field.Calendar;
            if (field.LangId != null) row.LangId = field.LangId;
        }

        private IReadOnlyList<VisioShapeDataSchemaIssue> Validate(
            string targetKind,
            string targetId,
            string? targetText,
            Func<string, VisioShapeDataRow?> rowResolver,
            Func<string, string?> valueResolver) {
            List<VisioShapeDataSchemaIssue> issues = new();
            foreach (VisioShapeDataField field in _fields) {
                VisioShapeDataRow? row = rowResolver(field.Name);
                if (row == null) {
                    issues.Add(new VisioShapeDataSchemaIssue(
                        VisioShapeDataSchemaIssueKind.MissingField,
                        targetKind,
                        targetId,
                        targetText,
                        field.Name,
                        "Missing Shape Data field '" + field.Name + "'."));
                    continue;
                }

                string? value = valueResolver(field.Name);
                if (field.Required && string.IsNullOrWhiteSpace(value)) {
                    issues.Add(new VisioShapeDataSchemaIssue(
                        VisioShapeDataSchemaIssueKind.MissingValue,
                        targetKind,
                        targetId,
                        targetText,
                        field.Name,
                        "Shape Data field '" + field.Name + "' requires a value."));
                }

                if (field.Type.HasValue && row.Type != field.Type.Value) {
                    issues.Add(new VisioShapeDataSchemaIssue(
                        VisioShapeDataSchemaIssueKind.TypeMismatch,
                        targetKind,
                        targetId,
                        targetText,
                        field.Name,
                        "Shape Data field '" + field.Name + "' type is '" + (row.Type?.ToString() ?? "not set") + "' but expected '" + field.Type.Value + "'."));
                }

                if (!string.IsNullOrEmpty(value) &&
                    field.AllowedValues.Count > 0 &&
                    !field.AllowedValues.Any(allowed => string.Equals(allowed, value, StringComparison.OrdinalIgnoreCase))) {
                    issues.Add(new VisioShapeDataSchemaIssue(
                        VisioShapeDataSchemaIssueKind.ValueNotAllowed,
                        targetKind,
                        targetId,
                        targetText,
                        field.Name,
                        "Shape Data field '" + field.Name + "' value '" + value + "' is not in the allowed value list."));
                }
            }

            return issues;
        }
    }

    /// <summary>
    /// Describes one reusable Shape Data row in a <see cref="VisioShapeDataSchema"/>.
    /// </summary>
    public sealed class VisioShapeDataField {
        private IReadOnlyList<string> _allowedValues = Array.Empty<string>();

        /// <summary>
        /// Initializes a Shape Data schema field.
        /// </summary>
        /// <param name="name">Shape Data row name.</param>
        public VisioShapeDataField(string name) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Shape data name cannot be empty.", nameof(name));
            }

            Name = name;
        }

        /// <summary>ShapeSheet row name.</summary>
        public string Name { get; }

        /// <summary>Default value used when applying the schema to missing or overwritten fields.</summary>
        public string? DefaultValue { get; set; }

        /// <summary>Label shown in the Visio Shape Data window.</summary>
        public string? Label { get; set; }

        /// <summary>Prompt shown as help text in the Visio Shape Data window.</summary>
        public string? Prompt { get; set; }

        /// <summary>Shape data type.</summary>
        public VisioShapeDataType? Type { get; set; }

        /// <summary>Format picture or list items, depending on the shape data type.</summary>
        public string? Format { get; set; }

        /// <summary>Sort key used by Visio's Shape Data window.</summary>
        public string? SortKey { get; set; }

        /// <summary>Whether validation should require a non-empty value.</summary>
        public bool Required { get; set; }

        /// <summary>Whether the data row is hidden in the Shape Data window.</summary>
        public bool? Invisible { get; set; }

        /// <summary>Whether Visio asks for this value when the shape is created, copied, or duplicated.</summary>
        public bool? Verify { get; set; }

        /// <summary>Whether this row is linked to an external data recordset.</summary>
        public bool? DataLinked { get; set; }

        /// <summary>Calendar type used for date values.</summary>
        public string? Calendar { get; set; }

        /// <summary>Language identifier used for the value.</summary>
        public string? LangId { get; set; }

        /// <summary>Allowed values used by validation and list-style field formats.</summary>
        public IReadOnlyList<string> AllowedValues {
            get => _allowedValues;
            set => _allowedValues = value?.Where(item => !string.IsNullOrWhiteSpace(item)).ToArray() ?? Array.Empty<string>();
        }

        internal string? GetEffectiveFormat() {
            if (!string.IsNullOrEmpty(Format) || AllowedValues.Count == 0) {
                return Format;
            }

            if (Type == VisioShapeDataType.FixedList || Type == VisioShapeDataType.VariableList) {
                return string.Join(";", AllowedValues);
            }

            return Format;
        }
    }

    /// <summary>
    /// Type of Shape Data schema validation issue.
    /// </summary>
    public enum VisioShapeDataSchemaIssueKind {
        /// <summary>A required schema field is missing.</summary>
        MissingField,
        /// <summary>A required schema field exists but has no value.</summary>
        MissingValue,
        /// <summary>The actual Shape Data type does not match the schema.</summary>
        TypeMismatch,
        /// <summary>The actual value is outside the allowed value list.</summary>
        ValueNotAllowed
    }

    /// <summary>
    /// One validation issue produced by <see cref="VisioShapeDataSchema"/>.
    /// </summary>
    public sealed class VisioShapeDataSchemaIssue {
        internal VisioShapeDataSchemaIssue(
            VisioShapeDataSchemaIssueKind kind,
            string targetKind,
            string targetId,
            string? targetText,
            string fieldName,
            string message) {
            Kind = kind;
            TargetKind = targetKind;
            TargetId = targetId;
            TargetText = targetText;
            FieldName = fieldName;
            Message = message;
        }

        /// <summary>Issue kind.</summary>
        public VisioShapeDataSchemaIssueKind Kind { get; }

        /// <summary>Target kind, such as Shape or Connector.</summary>
        public string TargetKind { get; }

        /// <summary>Target identifier.</summary>
        public string TargetId { get; }

        /// <summary>Target visible text or label, when available.</summary>
        public string? TargetText { get; }

        /// <summary>Shape Data field name.</summary>
        public string FieldName { get; }

        /// <summary>Human-readable validation message.</summary>
        public string Message { get; }

        /// <inheritdoc />
        public override string ToString() {
            return TargetKind + " " + TargetId + " " + FieldName + ": " + Message;
        }
    }
}
