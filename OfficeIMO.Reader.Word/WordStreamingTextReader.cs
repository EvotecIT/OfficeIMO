using System.Globalization;

namespace OfficeIMO.Reader.Word;

/// <summary>
/// The text projection's streamed read of a DOCX body: <see cref="global::OfficeIMO.Word.WordStreamedBodyText"/> reads
/// the lines without the Word model or the Open XML SDK, and this turns them into text chunks. When that pass declines,
/// the caller builds the model and the read scope reduces that result to the same text chunks.
/// </summary>
internal static class WordStreamingTextReader {
    /// <summary>Whether the read asked for the text projection and nothing this pass cannot honor.</summary>
    internal static bool Applies(ReaderOptions readerOptions) =>
        DocumentReaderEngine.IsTextProjection
        && string.IsNullOrEmpty(readerOptions.OpenPassword)
        && readerOptions.Inspection == null
        && FormatInternals.OpenXmlReaderAdmission.Policy(readerOptions, global::OfficeIMO.Word.WordLoadOptions.DefaultMaxInputBytes) == null;

    /// <summary>Returns the streamed result, or null to read the document through the Word model.</summary>
    internal static OfficeDocumentReadResult? TryRead(Stream stream, string sourceName, ReaderOptions readerOptions, ReaderWordOptions options,
        CancellationToken cancellationToken) {
        if (!stream.CanSeek) return null;
        long start = stream.Position;
        if (stream.Length - start > global::OfficeIMO.Word.WordLoadOptions.DefaultMaxInputBytes) return null;
        try {
            List<string>? lines = global::OfficeIMO.Word.WordStreamedBodyText.TryReadLines(stream, readerOptions.MaxChars,
                readerOptions.MaxTableRows, readerOptions.OpenXmlMaxCharactersInPart, options.IncludeFootnotes, cancellationToken, out _);
            return lines == null ? null : CreateResult(lines, sourceName);
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            // The Word model reports damaged or unusual packages with its own errors.
            return null;
        } finally {
            stream.Position = start;
        }
    }

    private static OfficeDocumentReadResult CreateResult(List<string> lines, string sourceName) {
        string fileName = Path.GetFileName(sourceName);
        var chunks = new List<ReaderChunk>();
        foreach (string line in lines) {
            if (line.Length == 0) continue;
            int index = chunks.Count;
            chunks.Add(new ReaderChunk {
                Id = "word:" + fileName + ":b" + index.ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Word,
                Location = new ReaderLocation { Path = sourceName, BlockIndex = index },
                Text = line
            });
        }
        return DocumentReaderEngine.CreateTextResult(chunks, ReaderInputKind.Word,
            new[] { OfficeDocumentReaderBuilderWordExtensions.HandlerId, "officeimo.reader.word.streamed-text" });
    }
}
