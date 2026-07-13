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
    public static MemoryStream ToRtfStream(this WordDocument document, RtfWriteOptions? options = null, Encoding? encoding = null) {
        return document.ToRtfDocument().ToStream(options, encoding);
    }

    /// <summary>Saves a Word document as an RTF file.</summary>
    public static void SaveAsRtf(this WordDocument document, string path, RtfWriteOptions? options = null, Encoding? encoding = null) {
        document.ToRtfDocument().Save(path, options, encoding);
    }

    /// <summary>Saves a Word document as RTF to a stream without closing or rewinding the stream.</summary>
    public static void SaveAsRtf(this WordDocument document, Stream stream, RtfWriteOptions? options = null, Encoding? encoding = null) {
        document.ToRtfDocument().Save(stream, options, encoding);
    }

}
