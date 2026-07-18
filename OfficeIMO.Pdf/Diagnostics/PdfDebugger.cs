using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

/// <summary>Creates bounded, read-only debugger projections without exposing a mutable PDF object model.</summary>
internal static class PdfDebugger {
    /// <summary>Dumps a PDF byte array into typed debugger records.</summary>
    public static PdfDebuggerReport Dump(byte[] pdf, PdfDebuggerOptions? options = null, PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfDebuggerOptions effectiveOptions = options ?? new PdfDebuggerOptions();
        effectiveOptions.Validate();
        PdfReadOptions effectiveReadOptions = PdfReadOptions.Resolve(readOptions);
        PdfReadDocument document = PdfReadDocument.Open(pdf, effectiveReadOptions);
        var (objects, _) = PdfSyntax.ParseObjects(pdf, effectiveReadOptions);
        HashSet<int> reachable = GetReachableObjectNumbers(objects, document.Security);
        var objectReports = objects.Values
            .OrderBy(static item => item.ObjectNumber)
            .Select(item => BuildObject(item, objects, reachable, effectiveOptions))
            .ToArray();
        var pages = new List<PdfDebugPage>(document.Pages.Count);
        for (int i = 0; i < document.Pages.Count; i++) {
            pages.Add(BuildPage(
                i + 1,
                document.Pages[i].ObjectNumber,
                objects,
                effectiveOptions,
                effectiveReadOptions.Limits.MaxContentNestingDepth));
        }

        return new PdfDebuggerReport(objectReports, document.Security.Revisions, pages.AsReadOnly(), document.RepairReport);
    }

    /// <summary>Dumps a PDF file into typed debugger records.</summary>
    public static PdfDebuggerReport Dump(string path, PdfDebuggerOptions? options = null, PdfReadOptions? readOptions = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Dump(File.ReadAllBytes(path), options, readOptions);
    }

    /// <summary>Dumps a readable PDF stream from its current position.</summary>
    public static PdfDebuggerReport Dump(Stream stream, PdfDebuggerOptions? options = null, PdfReadOptions? readOptions = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Dump(buffer.ToArray(), options, readOptions);
    }

    private static PdfDebugObject BuildObject(PdfIndirectObject indirect, Dictionary<int, PdfIndirectObject> objects, HashSet<int> reachable, PdfDebuggerOptions options) {
        PdfDictionary? dictionary = indirect.Value is PdfStream stream ? stream.Dictionary : indirect.Value as PdfDictionary;
        var references = new HashSet<int>();
        CollectReferences(indirect.Value, references);
        long? streamLength = null;
        long? decodedLength = null;
        string? preview = null;
        if (indirect.Value is PdfStream objectStream) {
            streamLength = objectStream.Data.LongLength;
            if (StreamDecoder.TryDecode(objectStream.Dictionary, objectStream.Data, options.MaxDecodedStreamPreviewBytes, out byte[] decoded, objects)) {
                decodedLength = decoded.LongLength;
                if (options.IncludeDecodedStreamPreviews) {
                    preview = CreatePreview(decoded);
                }
            }
        }

        return new PdfDebugObject(
            indirect.ObjectNumber,
            indirect.Generation,
            GetKind(indirect.Value),
            dictionary?.Items.Keys.OrderBy(static key => key, StringComparer.Ordinal).ToArray() ?? Array.Empty<string>(),
            references.OrderBy(static number => number).ToArray(),
            reachable.Contains(indirect.ObjectNumber),
            streamLength,
            decodedLength,
            preview);
    }

    private static PdfDebugPage BuildPage(
        int pageNumber,
        int objectNumber,
        Dictionary<int, PdfIndirectObject> objects,
        PdfDebuggerOptions options,
        int maxContentNestingDepth) {
        if (!objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect) || indirect.Value is not PdfDictionary page) {
            return new PdfDebugPage(pageNumber, objectNumber, Array.Empty<string>(), Array.Empty<int>(), Array.Empty<string>(), false);
        }

        PdfDictionary? resources = Resolve(page.Items.TryGetValue("Resources", out PdfObject? resourceValue) ? resourceValue : null, objects) as PdfDictionary;
        var contentObjectNumbers = new List<int>();
        var operators = new List<string>();
        bool truncated = false;
        if (page.Items.TryGetValue("Contents", out PdfObject? contents)) {
            foreach (PdfObject content in EnumerateContentObjects(contents)) {
                if (content is PdfReference reference) {
                    contentObjectNumbers.Add(reference.ObjectNumber);
                }

                if (Resolve(content, objects) is PdfStream contentStream && operators.Count < options.MaxContentOperatorsPerPage) {
                    byte[] decoded = StreamDecoder.Decode(contentStream.Dictionary, contentStream.Data, objects, options.MaxDecodedStreamPreviewBytes * 256);
                    PdfContentOperatorScanner.AppendOperators(
                        PdfEncoding.Latin1GetString(decoded),
                        operators,
                        options.MaxContentOperatorsPerPage,
                        ref truncated,
                        maxContentNestingDepth);
                }
            }
        }

        return new PdfDebugPage(
            pageNumber,
            objectNumber,
            resources?.Items.Keys.OrderBy(static key => key, StringComparer.Ordinal).ToArray() ?? Array.Empty<string>(),
            contentObjectNumbers.Distinct().OrderBy(static number => number).ToArray(),
            operators.AsReadOnly(),
            truncated);
    }

    private static IEnumerable<PdfObject> EnumerateContentObjects(PdfObject value) {
        if (value is PdfArray array) {
            for (int i = 0; i < array.Items.Count; i++) {
                yield return array.Items[i];
            }
        } else {
            yield return value;
        }
    }

    private static HashSet<int> GetReachableObjectNumbers(Dictionary<int, PdfIndirectObject> objects, PdfDocumentSecurityInfo security) {
        var reachable = new HashSet<int>();
        var pending = new Stack<int>();
        if (security.RootObjectNumber.HasValue) pending.Push(security.RootObjectNumber.Value);
        if (security.InfoObjectNumber.HasValue) pending.Push(security.InfoObjectNumber.Value);
        if (security.EncryptObjectNumber.HasValue) pending.Push(security.EncryptObjectNumber.Value);
        while (pending.Count > 0) {
            int objectNumber = pending.Pop();
            if (!reachable.Add(objectNumber) || !objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect)) {
                continue;
            }

            var references = new HashSet<int>();
            CollectReferences(indirect.Value, references);
            foreach (int reference in references) pending.Push(reference);
        }

        return reachable;
    }

    private static void CollectReferences(PdfObject value, HashSet<int> references) {
        switch (value) {
            case PdfReference reference:
                references.Add(reference.ObjectNumber);
                break;
            case PdfArray array:
                for (int i = 0; i < array.Items.Count; i++) CollectReferences(array.Items[i], references);
                break;
            case PdfDictionary dictionary:
                foreach (PdfObject item in dictionary.Items.Values) CollectReferences(item, references);
                break;
            case PdfStream stream:
                CollectReferences(stream.Dictionary, references);
                break;
        }
    }

    private static PdfObject? Resolve(PdfObject? value, Dictionary<int, PdfIndirectObject> objects) =>
        value is PdfReference reference && PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)
            ? indirect.Value
            : value;

    private static string GetKind(PdfObject value) {
        PdfDictionary? dictionary = value is PdfStream stream ? stream.Dictionary : value as PdfDictionary;
        string? type = dictionary?.Get<PdfName>("Type")?.Name;
        string? subtype = dictionary?.Get<PdfName>("Subtype")?.Name;
        string prefix = value is PdfStream ? "Stream" : value is PdfDictionary ? "Dictionary" : value.GetType().Name;
        return !string.IsNullOrEmpty(subtype) ? prefix + "." + subtype : !string.IsNullOrEmpty(type) ? prefix + "." + type : prefix;
    }

    private static string CreatePreview(byte[] decoded) {
        var builder = new StringBuilder(decoded.Length);
        for (int i = 0; i < decoded.Length; i++) {
            char value = (char)decoded[i];
            builder.Append(char.IsControl(value) && value != '\r' && value != '\n' && value != '\t' ? '\uFFFD' : value);
        }

        return builder.ToString();
    }
}
