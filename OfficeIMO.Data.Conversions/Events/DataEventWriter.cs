namespace OfficeIMO.Data.Conversions;

/// <summary>A forward-only sink for one target format. Each sink reports what its format cannot carry.</summary>
internal abstract class DataEventWriter : IDisposable {
    protected DataEventWriter(DataConversionContext context) {
        Context = context;
    }

    public DataConversionContext Context { get; }

    public abstract void WriteStartDocument();
    public abstract void WriteEndDocument();
    public abstract void WriteStartObject();
    public abstract void WriteEndObject();
    public abstract void WriteStartArray();
    public abstract void WriteEndArray();
    public abstract void WritePropertyName(ReadOnlySpan<char> name);
    public abstract void WriteScalar(DataScalarKind kind, ReadOnlySpan<char> text);

    /// <summary>Writes a comment. Formats without comments report the omission once per conversion.</summary>
    public virtual void WriteComment(ReadOnlySpan<char> text) =>
        Context.Report(OfficeDiagnosticCodes.CommentsOmitted, OfficeConversionLossKind.Omission,
            "Comments were dropped because the target format has no comments.");

    /// <summary>Completes the output after the last event.</summary>
    public abstract void Complete();

    public virtual void Dispose() {
    }
}

internal static class DataCopier {
    /// <summary>Pipes every event from the reader into the writer and completes the writer.</summary>
    public static long Copy(DataEventReader reader, DataEventWriter writer) {
        long documents = 0;
        while (reader.Read()) {
            switch (reader.TokenType) {
                case DataTokenType.StartDocument: writer.WriteStartDocument(); documents++; break;
                case DataTokenType.EndDocument: writer.WriteEndDocument(); break;
                case DataTokenType.StartObject: writer.WriteStartObject(); break;
                case DataTokenType.EndObject: writer.WriteEndObject(); break;
                case DataTokenType.StartArray: writer.WriteStartArray(); break;
                case DataTokenType.EndArray: writer.WriteEndArray(); break;
                case DataTokenType.PropertyName: writer.WritePropertyName(reader.Text); break;
                case DataTokenType.Scalar: writer.WriteScalar(reader.ScalarKind, reader.Text); break;
                case DataTokenType.Comment: writer.WriteComment(reader.Text); break;
            }
        }

        writer.Complete();
        return documents;
    }
}
