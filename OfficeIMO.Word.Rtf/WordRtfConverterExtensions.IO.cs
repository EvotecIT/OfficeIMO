namespace OfficeIMO.Word.Rtf;

/// <summary>
/// Extension methods for Word and RTF serialization input and output.
/// </summary>
public static partial class WordRtfConverterExtensions {
    /// <summary>Serializes a Word document to RTF text.</summary>
    public static string ToRtf(this WordDocument document, RtfWriteOptions? options = null) {
        return document.ToRtfDocument().ToRtf(options);
    }

    /// <summary>Serializes a Word document to encoded RTF bytes.</summary>
    public static byte[] ToRtfBytes(this WordDocument document, RtfWriteOptions? options = null, Encoding? encoding = null) {
        return document.ToRtfDocument().ToBytes(options, encoding);
    }

    /// <summary>Serializes a Word document to an encoded RTF memory stream.</summary>
    public static MemoryStream ToRtfMemoryStream(this WordDocument document, RtfWriteOptions? options = null, Encoding? encoding = null) {
        return document.ToRtfDocument().ToMemoryStream(options, encoding);
    }

    /// <summary>Saves a Word document as an RTF file.</summary>
    public static void SaveAsRtf(this WordDocument document, string path, RtfWriteOptions? options = null, Encoding? encoding = null) {
        document.ToRtfDocument().Save(path, options, encoding);
    }

    /// <summary>Saves a Word document as RTF to a stream without closing or rewinding the stream.</summary>
    public static void SaveAsRtf(this WordDocument document, Stream stream, RtfWriteOptions? options = null, Encoding? encoding = null) {
        document.ToRtfDocument().Save(stream, options, encoding);
    }

    /// <summary>Creates a Word document from RTF text.</summary>
    public static WordDocument LoadFromRtf(this string rtf, RtfReadOptions? readOptions = null) {
        RtfReadResult result = RtfDocument.Read(rtf, readOptions);
        return result.Document.ToWordDocument();
    }

    /// <summary>Creates a Word document from source RTF bytes using the core byte-preserving RTF reader.</summary>
    public static WordDocument LoadFromRtf(this byte[] rtfBytes, RtfReadOptions? readOptions = null) {
        RtfReadResult result = RtfDocument.Load(rtfBytes, readOptions);
        return result.Document.ToWordDocument();
    }

    /// <summary>Creates a Word document from an RTF stream, reading from the stream's current position.</summary>
    public static WordDocument LoadFromRtf(this Stream rtfStream, RtfReadOptions? readOptions = null, Encoding? encoding = null) {
        RtfReadResult result = RtfDocument.Load(rtfStream, readOptions, encoding);
        return result.Document.ToWordDocument();
    }

    /// <summary>Loads an RTF file and converts it to a Word document.</summary>
    public static WordDocument LoadFromRtfFile(string path, RtfReadOptions? readOptions = null, Encoding? encoding = null) {
        RtfReadResult result = RtfDocument.Load(path, readOptions, encoding);
        return result.Document.ToWordDocument();
    }
}
