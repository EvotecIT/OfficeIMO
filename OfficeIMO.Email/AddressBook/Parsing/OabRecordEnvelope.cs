namespace OfficeIMO.Email.AddressBook;

internal sealed class OabRecordEnvelope {
    internal OabRecordEnvelope(int size, byte[] body) {
        Size = size;
        Body = body;
    }

    internal int Size { get; }
    internal byte[] Body { get; }
}

internal sealed class OabParsedRecord {
    internal OabParsedRecord(IReadOnlyList<OfficeIMO.Email.MapiProperty> properties,
        IReadOnlyList<OfficeIMO.Email.EmailDiagnostic> diagnostics) {
        Properties = properties;
        Diagnostics = diagnostics;
    }

    internal IReadOnlyList<OfficeIMO.Email.MapiProperty> Properties { get; }
    internal IReadOnlyList<OfficeIMO.Email.EmailDiagnostic> Diagnostics { get; }
}
