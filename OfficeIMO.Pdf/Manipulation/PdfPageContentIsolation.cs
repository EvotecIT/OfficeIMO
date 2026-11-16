using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

/// <summary>Isolates preserved page streams before adding content in the page's default graphics state.</summary>
/// <remarks>
/// Page content does not have to balance its own <c>q</c>/<c>Q</c> operators. A stream that leaves several states saved, or
/// that restores more states than it saved, defeats a single surrounding pair: viewers then apply the original content's
/// transformation, clip or text state to everything appended after it, or drop a transformation placed in front of it.
/// Wrappers therefore save enough extra states to absorb every surplus restore and restore every state the original
/// content leaves open.
/// </remarks>
internal static class PdfPageContentIsolation {
    // Readers commonly cap q nesting near this depth; an extremely malformed stream is isolated as far as that allows.
    private const int MaximumExtraSaves = 27;

    /// <summary>
    /// Save-state nesting of existing page content: the extra saves a wrapper needs after its own leading <c>q</c>, and the
    /// restores that return the wrapped content to the state before that <c>q</c>.
    /// </summary>
    internal readonly struct Nesting {
        internal Nesting(int extraSaves, int restores) {
            ExtraSaves = extraSaves;
            Restores = restores;
        }

        internal int ExtraSaves { get; }
        internal int Restores { get; }

        /// <summary>Text placed after a wrapper's leading <c>q</c> and any wrapper transform.</summary>
        internal string ExtraSaveOperators => Repeat("q\n", ExtraSaves);

        /// <summary>Text of the stream that closes the wrapped content.</summary>
        internal string RestoreOperators => "\n" + Repeat("Q\n", Restores);
    }

    /// <summary>Wraps original stream entries without decoding, repairing, or changing their bytes.</summary>
    internal static PdfArray Isolate(Dictionary<int, PdfIndirectObject> objects, PdfObject contents,
        ref int nextObjectNumber, CancellationToken cancellationToken = default) =>
        Isolate(objects, contents, ref nextObjectNumber, out _, cancellationToken);

    /// <summary>Wraps original stream entries and reports the number of added wrapper bytes.</summary>
    internal static PdfArray Isolate(Dictionary<int, PdfIndirectObject> objects, PdfObject contents,
        ref int nextObjectNumber, out int wrapperBytes, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        List<PdfObject> entries = GetEntries(objects, contents, cancellationToken);
        Nesting nesting = Measure(objects, entries, cancellationToken);
        byte[] saveBytes = PdfEncoding.Latin1GetBytes("q\n" + nesting.ExtraSaveOperators);
        // This delimiter belongs to the added stream: the original may end at its last operator.
        byte[] restoreBytes = PdfEncoding.Latin1GetBytes(nesting.RestoreOperators);
        int save = checked(nextObjectNumber++);
        int restore = checked(nextObjectNumber++);
        objects[save] = new PdfIndirectObject(save, 0, new PdfStream(new PdfDictionary(), saveBytes));
        objects[restore] = new PdfIndirectObject(restore, 0, new PdfStream(new PdfDictionary(), restoreBytes));
        var isolated = new PdfArray();
        isolated.Items.Add(new PdfReference(save, 0));
        foreach (PdfObject entry in entries) isolated.Items.Add(entry);
        isolated.Items.Add(new PdfReference(restore, 0));
        wrapperBytes = saveBytes.Length + restoreBytes.Length;
        return isolated;
    }

    /// <summary>Appends a page-space stream after restoring the original page graphics state.</summary>
    /// <returns>The number of added wrapper bytes, or zero when there was no original content.</returns>
    internal static int Append(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page,
        int appendedObjectNumber, ref int nextObjectNumber, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var appended = new PdfReference(appendedObjectNumber, 0);
        if (!page.Items.TryGetValue("Contents", out PdfObject? contents)) {
            page.Items["Contents"] = appended;
            return 0;
        }
        PdfArray isolated = Isolate(objects, contents, ref nextObjectNumber, out int wrapperBytes, cancellationToken);
        isolated.Items.Add(appended);
        page.Items["Contents"] = isolated;
        return wrapperBytes;
    }

    /// <summary>Measures the save-state nesting of a page /Contents value, which may be absent.</summary>
    internal static Nesting Measure(Dictionary<int, PdfIndirectObject> objects, PdfObject? contents,
        CancellationToken cancellationToken = default) =>
        contents is null ? new Nesting(0, 1) : Measure(objects, GetEntries(objects, contents, cancellationToken), cancellationToken);

    private static Nesting Measure(Dictionary<int, PdfIndirectObject> objects, IReadOnlyList<PdfObject> entries,
        CancellationToken cancellationToken) {
        (int Lowest, int Final)? balance = MeasureBalance(objects, entries, cancellationToken);
        if (balance is null) return new Nesting(0, 1);
        int extraSaves = Math.Min(-balance.Value.Lowest, MaximumExtraSaves);
        return new Nesting(extraSaves, Math.Max(1, 1 + extraSaves + balance.Value.Final));
    }

    private static List<PdfObject> GetEntries(Dictionary<int, PdfIndirectObject> objects, PdfObject contents,
        CancellationToken cancellationToken) {
        var entries = new List<PdfObject>();
        if (PdfObjectLookup.Resolve(objects, contents) is PdfArray array) {
            foreach (PdfObject item in array.Items) {
                cancellationToken.ThrowIfCancellationRequested();
                entries.Add(item);
            }
        } else {
            entries.Add(contents);
        }
        return entries;
    }

    /// <summary>
    /// Returns the lowest and final save depth of the concatenated streams, or null when the content cannot be decoded or
    /// tokenized. Unreadable content keeps the single surrounding pair.
    /// </summary>
    private static (int Lowest, int Final)? MeasureBalance(Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyList<PdfObject> entries, CancellationToken cancellationToken) {
        // Content streams in a page array form one operator sequence, so depth carries across stream boundaries.
        using var concatenated = new MemoryStream();
        foreach (PdfObject entry in entries) {
            cancellationToken.ThrowIfCancellationRequested();
            if (PdfObjectLookup.Resolve(objects, entry) is not PdfStream stream ||
                !StreamDecoder.TryDecode(stream, PdfReadLimits.DefaultMaxDecodedStreamBytes, out byte[] decoded, objects)) {
                return null;
            }
            concatenated.Write(decoded, 0, decoded.Length);
            concatenated.WriteByte((byte)'\n');
        }

        int depth = 0, lowest = 0;
        try {
            PdfContentStreamInterpreter.InterpretUntil(
                PdfContentSource.FromBytes(concatenated.ToArray()),
                int.MaxValue,
                operation => {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (operation.Name == "q") {
                        depth++;
                    } else if (operation.Name == "Q" && --depth < lowest) {
                        lowest = depth;
                    }
                    return true;
                },
                maxGraphicsStateDepth: int.MaxValue);
        } catch (Exception exception) when (exception is InvalidDataException or PdfReadLimitException or FormatException) {
            return null;
        }
        return (lowest, depth);
    }

    private static string Repeat(string value, int count) {
        if (count <= 0) return string.Empty;
        var builder = new StringBuilder(value.Length * count);
        for (int i = 0; i < count; i++) builder.Append(value);
        return builder.ToString();
    }
}
