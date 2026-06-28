namespace OfficeIMO.Excel.LegacyXls.Model {
    internal sealed class LegacyXlsDocumentPropertyValue {
        internal LegacyXlsDocumentPropertyValue(object? value, LegacyXlsDocumentPropertyValueKind kind) {
            Value = value;
            Kind = kind;
        }

        internal object? Value { get; }

        internal LegacyXlsDocumentPropertyValueKind Kind { get; }
    }
}
