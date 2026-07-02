namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocDocumentPropertyValue {
        internal LegacyDocDocumentPropertyValue(object value, LegacyDocDocumentPropertyValueKind kind) {
            Value = value ?? throw new ArgumentNullException(nameof(value));
            Kind = kind;
        }

        internal object Value { get; }

        internal LegacyDocDocumentPropertyValueKind Kind { get; }
    }
}
