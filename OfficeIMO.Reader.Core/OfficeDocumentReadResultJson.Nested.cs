using System.Text.Json;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentReadResultJson {
    private static void EnsureAcyclicNestedResults(OfficeDocumentReadResult result, HashSet<OfficeDocumentReadResult> ancestors, int depth) {
        if (depth > 64 || !ancestors.Add(result)) throw new JsonException("Nested result graph is cyclic or too deep.");
        foreach (var child in result.NestedDocuments ?? Array.Empty<OfficeDocumentNestedResult>()) {
            if (child == null || child.Document == null) throw new JsonException("A nested document requires a document.");
            EnsureAcyclicNestedResults(child.Document, ancestors, depth + 1);
        }
        ancestors.Remove(result);
    }

    private static void EnsureNestedDeserializedContracts(OfficeDocumentReadResult result) {
        foreach (var nested in result.NestedDocuments ?? Array.Empty<OfficeDocumentNestedResult>()) {
            if (nested.Document.SchemaVersion != result.SchemaVersion) throw new JsonException("A nested transport document must use the parent schema version.");
            EnsureKindSupported(nested.Document.SchemaVersion, nested.Document.Kind);
            EnsureChunkKindsSupported(nested.Document.SchemaVersion, nested.Document.Chunks);
            EnsureDiagnosticContracts(nested.Document.Diagnostics);
            EnsureNestedDeserializedContracts(nested.Document);
        }
    }

    private static void ValidateTransportEnvelope(JsonElement root) {
        if (root.ValueKind != JsonValueKind.Object) throw new JsonException("A nested document must be an object.");
        int version = TryReadSchemaVersion(root);
        OfficeDocumentReadResultSchema.EnsureSupported(TryReadSchemaId(root), version);
        EnsureRequiredTopLevelProperties(root);
        EnsureKnownTopLevelProperties(root);
        EnsureNestedTransportContracts(root);
        if (!root.TryGetProperty("nestedDocuments", out var nested)) {
            if (version >= 9) throw new JsonException("Required document read result property 'nestedDocuments' is missing.");
            return;
        }
        if (version < 9) throw new JsonException("Nested documents require schema version 9.");
        EnsureObjectArray(nested, "nestedDocuments");
        foreach (var child in nested.EnumerateArray()) {
            foreach (var property in child.EnumerateObject())
                if (property.Name != "path" && property.Name != "document") throw new JsonException("Unknown nested document property.");
            if (!child.TryGetProperty("path", out var path) || path.ValueKind != JsonValueKind.String ||
                !child.TryGetProperty("document", out var document)) throw new JsonException("A nested document requires path and document.");
            ValidateTransportEnvelope(document);
        }
    }

    private static IReadOnlyList<OfficeDocumentNestedResult> NormalizeNestedForSerialization(OfficeDocumentReadResult result, int version) {
        var children = result.NestedDocuments ?? Array.Empty<OfficeDocumentNestedResult>();
        if (version < 9) {
            if (children.Count > 0) throw new JsonException("Nested documents require schema version 9.");
            return null!; // The versioned transport omits this property for schemas 5 through 8.
        }
        var normalized = new OfficeDocumentNestedResult[children.Count];
        for (int i = 0; i < children.Count; i++) {
            var child = children[i];
            if (child == null || child.Path == null || child.Document == null) throw new JsonException("A nested document requires path and document.");
            var document = child.Document;
            string id = string.IsNullOrWhiteSpace(document.SchemaId) ? OfficeDocumentReadResultSchema.Id : document.SchemaId;
            int childVersion = document.SchemaVersion == 0 ? OfficeDocumentReadResultSchema.CurrentVersion : document.SchemaVersion;
            OfficeDocumentReadResultSchema.EnsureSupported(id, childVersion);
            EnsureKindSupported(childVersion, document.Kind);
            EnsureChunkKindsSupported(childVersion, document.Chunks);
            EnsureKindSupported(version, document.Kind);
            EnsureChunkKindsSupported(version, document.Chunks);
            EnsureStringCollection(document.CapabilitiesUsed, "capabilitiesUsed");
            EnsureDiagnosticContracts(document.Diagnostics);
            normalized[i] = new OfficeDocumentNestedResult { Path = child.Path,
                Document = NormalizeForSerialization(document, id, version) };
        }
        return normalized;
    }
}
