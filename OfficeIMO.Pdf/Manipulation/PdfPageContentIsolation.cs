using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

/// <summary>Isolates preserved page streams before adding content in the page's default graphics state.</summary>
/// <remarks>
/// Page content does not have to balance its own <c>q</c>/<c>Q</c> operators, and a single surrounding pair is not enough
/// when it does not. Content that leaves several states saved hands its transformation, clip or text state to everything
/// appended after the pair. A <c>Q</c> with nothing of its own to restore is ignored on the unwrapped page, but inside a
/// wrapper it restores the wrapper's state instead, so the original content escapes the wrapper and edits pick up what it
/// changes afterwards. Wrappers therefore close every state the original content leaves open, and streams that restore
/// more states than they saved are replaced by copies whose surplus <c>Q</c> operators are blanked, which keeps the
/// original page drawing exactly as before. Inline image data is never touched; once an inline image colour space cannot
/// be resolved, no later restore on the page is blanked.
/// </remarks>
internal static class PdfPageContentIsolation {
    /// <summary>The original content entries to wrap, and the restores that close them.</summary>
    internal sealed class Wrapping {
        internal Wrapping(IReadOnlyList<PdfObject> entries, int restores) {
            Entries = entries;
            Restores = restores;
        }

        /// <summary>Original stream entries, with replacements for streams whose surplus restores were blanked.</summary>
        internal IReadOnlyList<PdfObject> Entries { get; }

        /// <summary>Restores that return the wrapped content to the state before the wrapper's leading <c>q</c>.</summary>
        internal int Restores { get; }

        /// <summary>Text of the stream that closes the wrapped content.</summary>
        internal string RestoreOperators {
            get {
                var builder = new StringBuilder("\n", 1 + 2 * Restores);
                for (int i = 0; i < Restores; i++) builder.Append("Q\n");
                return builder.ToString();
            }
        }
    }

    /// <summary>Wraps original stream entries; only streams with surplus restores are replaced.</summary>
    internal static PdfArray Isolate(Dictionary<int, PdfIndirectObject> objects, PdfDictionary? page, PdfObject contents,
        ref int nextObjectNumber, CancellationToken cancellationToken = default) =>
        Isolate(objects, page, contents, ref nextObjectNumber, out _, cancellationToken);

    /// <summary>Wraps original stream entries and reports the number of added wrapper bytes.</summary>
    internal static PdfArray Isolate(Dictionary<int, PdfIndirectObject> objects, PdfDictionary? page, PdfObject contents,
        ref int nextObjectNumber, out int wrapperBytes, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        int next = nextObjectNumber;
        Wrapping wrapping = Prepare(objects, page, contents, stream => {
            int number = checked(next++);
            objects[number] = new PdfIndirectObject(number, 0, stream);
            return new PdfReference(number, 0);
        }, cancellationToken);
        nextObjectNumber = next;
        byte[] saveBytes = PdfEncoding.Latin1GetBytes("q\n");
        // This delimiter belongs to the added stream: the original may end at its last operator.
        byte[] restoreBytes = PdfEncoding.Latin1GetBytes(wrapping.RestoreOperators);
        int save = checked(nextObjectNumber++);
        int restore = checked(nextObjectNumber++);
        objects[save] = new PdfIndirectObject(save, 0, new PdfStream(new PdfDictionary(), saveBytes));
        objects[restore] = new PdfIndirectObject(restore, 0, new PdfStream(new PdfDictionary(), restoreBytes));
        var isolated = new PdfArray();
        isolated.Items.Add(new PdfReference(save, 0));
        foreach (PdfObject entry in wrapping.Entries) isolated.Items.Add(entry);
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
        PdfArray isolated = Isolate(objects, page, contents, ref nextObjectNumber, out int wrapperBytes, cancellationToken);
        isolated.Items.Add(appended);
        page.Items["Contents"] = isolated;
        return wrapperBytes;
    }

    /// <summary>
    /// Prepares a page /Contents value, which may be absent, for a wrapper that starts with one <c>q</c>.
    /// <paramref name="addStream"/> stores a replacement stream and returns the entry that refers to it.
    /// Content that cannot be decoded or tokenized is wrapped unchanged with a single restore.
    /// </summary>
    internal static Wrapping Prepare(Dictionary<int, PdfIndirectObject> objects, PdfDictionary? page, PdfObject? contents,
        Func<PdfStream, PdfObject> addStream, CancellationToken cancellationToken = default) {
        if (contents is null) return new Wrapping(Array.Empty<PdfObject>(), 1);
        List<PdfObject> entries = GetEntries(objects, contents, cancellationToken);
        var decodedEntries = new List<byte[]>(entries.Count);
        long totalBytes = 0;
        foreach (PdfObject entry in entries) {
            cancellationToken.ThrowIfCancellationRequested();
            if (PdfObjectLookup.Resolve(objects, entry) is not PdfStream stream ||
                !StreamDecoder.TryDecode(stream, PdfReadLimits.DefaultMaxDecodedStreamBytes, out byte[] decoded, objects) ||
                (totalBytes += decoded.LongLength) > PdfReadLimits.DefaultMaxPageContentBytes) {
                return new Wrapping(entries, 1);
            }
            decodedEntries.Add(decoded);
        }

        // Content streams in a page array form one operator sequence, so they are read as one source: a token that
        // runs across a stream boundary and the graphics-state depth both carry into the next stream. The added line
        // feed only separates streams for the scan; surplus offsets are mapped back to the stream that holds them.
        var starts = new int[entries.Count];
        byte[] combined = new byte[checked((int)totalBytes + Math.Max(0, entries.Count - 1))];
        int position = 0;
        for (int index = 0; index < entries.Count; index++) {
            if (index > 0) combined[position++] = (byte)'\n';
            starts[index] = position;
            Buffer.BlockCopy(decodedEntries[index], 0, combined, position, decodedEntries[index].Length);
            position += decodedEntries[index].Length;
        }

        // Inline image data ends where its colour space says, so aliases resolve through the page resources. Once an
        // inline image colour space cannot be resolved, its data could hide a Q, so no later restore is blanked: the
        // depth from that point is the same best estimate a viewer makes when it looks for the end of the data.
        PdfDictionary? resources = page is null ? null : ResolveInheritedResources(objects, page);
        bool uncertain = false;
        int ResolveComponents(PdfObject colorSpace) {
            int count = PdfPrintProductionColorInspector.ResolveInlineImageComponentCountForResources(colorSpace, resources, objects,
                MaximumObjectDepth, PdfReadLimits.DefaultMaxDecodedStreamBytes);
            if (count <= 0) uncertain = true;
            return count;
        }

        int depth = 0;
        var surplusRestores = new List<int>?[entries.Count];
        try {
            PdfContentStreamInterpreter.InterpretUntil(
                PdfContentSource.FromBytes(combined),
                int.MaxValue,
                operation => {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (operation.Name == "q") {
                        depth++;
                    } else if (operation.Name == "Q") {
                        if (depth > 0) {
                            depth--;
                        } else if (!uncertain) {
                            int offset = operation.OperatorOffset;
                            int index = Array.BinarySearch(starts, offset);
                            if (index < 0) index = ~index - 1;
                            if (index >= 0) (surplusRestores[index] ??= new List<int>()).Add(offset - starts[index]);
                        }
                    }
                    return true;
                },
                inlineImageComponentCount: name => ResolveComponents(new PdfName(name)),
                inlineImageArrayComponentCount: array => ResolveComponents(array),
                maxGraphicsStateDepth: int.MaxValue);
        } catch (Exception exception) when (exception is InvalidDataException or PdfReadLimitException or FormatException) {
            return new Wrapping(entries, 1);
        }

        var prepared = new List<PdfObject>(entries.Count);
        for (int index = 0; index < entries.Count; index++) {
            List<int>? surplus = surplusRestores[index];
            byte[] decoded = decodedEntries[index];
            if (surplus is null || !surplus.All(offset => offset >= 0 && offset < decoded.Length && decoded[offset] == (byte)'Q')) {
                prepared.Add(entries[index]);
                continue;
            }
            // A surplus Q is a no-op on the unwrapped page; blanking it keeps every other byte and offset in place.
            byte[] blanked = (byte[])decoded.Clone();
            foreach (int offset in surplus) blanked[offset] = (byte)' ';
            prepared.Add(addStream(new PdfStream(new PdfDictionary(), blanked)));
        }
        return new Wrapping(prepared, 1 + depth);
    }

    private const int MaximumObjectDepth = 128;

    private static PdfDictionary? ResolveInheritedResources(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page) {
        var visited = new HashSet<PdfDictionary>();
        for (PdfDictionary? node = page; node is not null && visited.Add(node);
            node = node.Items.TryGetValue("Parent", out PdfObject? parent) ? PdfObjectLookup.Resolve(objects, parent) as PdfDictionary : null) {
            if (node.Items.TryGetValue("Resources", out PdfObject? value)) return PdfObjectLookup.Resolve(objects, value) as PdfDictionary;
        }
        return null;
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
}
