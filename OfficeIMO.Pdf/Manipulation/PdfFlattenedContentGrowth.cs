using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Tracks generated flattening work and validates reused appearance streams.</summary>
internal sealed class PdfFlattenedContentGrowth {
    private readonly Dictionary<int, PdfIndirectObject> _objects;
    private readonly PdfReadLimits _limits;
    private readonly int? _firstGeneratedObject;
    private readonly CancellationToken _cancellationToken;
    private readonly Dictionary<PdfStream, byte[]> _decoded = new();
    private readonly PdfDecodedStreamBudget _sourceDecodeBudget;
    private long _totalPageBytes;
    private int _maximumPageBytes;
    private int _maximumRawGeneratedBytes;
    private int _maximumDecodedGeneratedBytes;
    private int _maximumOperations;
    private int _maximumOperands;
    private int _contexts;
    private long _operations;
    private long _sourceOperations;
    private int _sourceContexts;

    internal PdfFlattenedContentGrowth(Dictionary<int, PdfIndirectObject> objects, PdfLoadOptions? readOptions,
        int? firstGeneratedObject, CancellationToken cancellationToken, PdfDecodedStreamBudget? sourceDecodeBudget = null) {
        _objects = objects;
        _limits = PdfLoadOptions.Resolve(readOptions).Limits;
        _firstGeneratedObject = firstGeneratedObject;
        _sourceDecodeBudget = sourceDecodeBudget ?? new PdfDecodedStreamBudget(_limits);
        _cancellationToken = cancellationToken;
    }

    /// <summary>Admits retained normal appearances before generated output can increase shared read limits.</summary>
    internal static void ValidateSourceAppearances(PdfReadDocument source, CancellationToken cancellationToken) {
        var inspection = new PdfFlattenedContentGrowth(source.Objects, source.ReadOptions, firstGeneratedObject: null, cancellationToken, source.DecodedStreamBudget);
        foreach (PdfReadPage page in source.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            page.VisitNormalAppearanceStreams((stream, resources) => {
                byte[] decoded = source.DecodedStreamBudget.Decode(stream, source.Objects,
                    source.ReadOptions.Limits.MaxDecodedStreamBytes, cancellationToken);
                if (++inspection._sourceContexts > inspection._limits.MaxPrintProductionContexts)
                    throw PdfReadLimitException.Create(PdfReadLimitKind.PrintProductionContexts,
                        inspection._limits.MaxPrintProductionContexts, inspection._sourceContexts);
                long bytes = 0;
                int operations = 0, operands = 0;
                inspection.VisitContent(PdfEncoding.Latin1GetString(decoded), resources, generated: false, depth: 0,
                    new HashSet<PdfStream> { stream }, ref bytes, ref operations, ref operands);
            }, cancellationToken);
        }
    }

    internal void AddPage(PdfDictionary page, PdfStream addedStream, int isolationBytes) {
        _cancellationToken.ThrowIfCancellationRequested();
        PdfDictionary? resources = page.Items.TryGetValue("Resources", out PdfObject? value)
            ? PdfObjectLookup.ResolveChainCancellable(_objects, value, _cancellationToken) as PdfDictionary : null;
        long pageBytes = addedStream.Data.LongLength + isolationBytes;
        int operations = isolationBytes == 0 ? 0 : 2;
        int operands = 0;
        _contexts += isolationBytes == 0 ? 1 : 3;
        _maximumRawGeneratedBytes = Math.Max(_maximumRawGeneratedBytes, addedStream.Data.Length);
        _maximumDecodedGeneratedBytes = Math.Max(_maximumDecodedGeneratedBytes, addedStream.Data.Length);
        VisitContent(PdfEncoding.Latin1GetString(addedStream.Data), resources, generated: true, depth: 0,
            new HashSet<PdfStream>(), ref pageBytes, ref operations, ref operands);
        _maximumPageBytes = Math.Max(_maximumPageBytes, (int)Math.Min(int.MaxValue, pageBytes));
        _totalPageBytes = checked(_totalPageBytes + pageBytes);
        _maximumOperations = Math.Max(_maximumOperations, operations);
        _maximumOperands = Math.Max(_maximumOperands, operands);
        _operations = checked(_operations + operations);
    }

    private void VisitContent(string content, PdfDictionary? resources, bool generated, int depth,
        HashSet<PdfStream> active, ref long pageBytes, ref int operations, ref int operands) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (depth > _limits.MaxContentNestingDepth)
            throw PdfReadLimitException.Create(PdfReadLimitKind.ContentNestingDepth, _limits.MaxContentNestingDepth, depth);
        var invoked = new List<(PdfStream Stream, bool Generated)>();
        int localOperations = 0, localOperands = 0;
        PdfContentStreamInterpreter.Interpret(content,
            generated ? Math.Max(1, content.Length) : _limits.MaxContentOperations, operation => {
                _cancellationToken.ThrowIfCancellationRequested();
                if (!generated && ++_sourceOperations > _limits.MaxPrintProductionOperations)
                    throw PdfReadLimitException.Create(PdfReadLimitKind.PrintProductionOperations, _limits.MaxPrintProductionOperations, _sourceOperations);
                localOperations++;
                localOperands = checked(localOperands + operation.Operands.Count);
                if (operation.Name != "Do" || operation.Operands.Count != 1 || operation.Operands[0] is not string name ||
                    resources is null || !resources.Items.TryGetValue("XObject", out PdfObject? xObjectValue) ||
                    PdfObjectLookup.ResolveChainCancellable(_objects, xObjectValue, _cancellationToken) is not PdfDictionary xObjects ||
                    !xObjects.Items.TryGetValue(name, out PdfObject? xObject) ||
                    PdfObjectLookup.ResolveChainCancellable(_objects, xObject, _cancellationToken) is not PdfStream stream ||
                    stream.Dictionary.Get<PdfName>("Subtype")?.Name != "Form") return;
                // A nested form clipped to an empty box paints nothing and needs no stream decoding.
                if (PdfAppearancePlacement.IsKnownEmpty(stream.Dictionary,
                    value => PdfObjectLookup.ResolveChainCancellable(_objects, value, _cancellationToken), 1D, 1D)) return;
                invoked.Add((stream, _firstGeneratedObject.HasValue && xObject is PdfReference reference &&
                    reference.ObjectNumber >= _firstGeneratedObject.Value));
            }, maxNestingDepth: _limits.MaxContentNestingDepth,
            maxOperands: generated ? Math.Max(1, content.Length) : _limits.MaxContentOperands);
        if (generated) {
            operations = checked(operations + localOperations);
            operands = checked(operands + localOperands);
        }
        foreach (var invocation in invoked) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (!active.Add(invocation.Stream)) continue;
            try {
                if (!_decoded.TryGetValue(invocation.Stream, out byte[]? decoded)) {
                    decoded = invocation.Generated
                        ? Filters.StreamDecoder.Decode(invocation.Stream, _objects,
                            Math.Max(_limits.MaxDecodedStreamBytes, invocation.Stream.Data.Length), _cancellationToken)
                        : _sourceDecodeBudget.Decode(invocation.Stream, _objects, _limits.MaxDecodedStreamBytes, _cancellationToken);
                    _decoded.Add(invocation.Stream, decoded);
                    if (invocation.Generated) _maximumDecodedGeneratedBytes = Math.Max(_maximumDecodedGeneratedBytes, decoded.Length);
                }
                if (!invocation.Generated && ++_sourceContexts > _limits.MaxPrintProductionContexts)
                    throw PdfReadLimitException.Create(PdfReadLimitKind.PrintProductionContexts, _limits.MaxPrintProductionContexts, _sourceContexts);
                // Existing appearances were already part of an admitted source read. Validate them here,
                // but reserve additional output work only for content created by this mutation.
                if (invocation.Generated) {
                    pageBytes = checked(pageBytes + decoded.LongLength);
                    _contexts++;
                }
                PdfDictionary? childResources = invocation.Stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? childValue)
                    ? (PdfObjectLookup.ResolveChainCancellable(_objects, childValue, _cancellationToken) as PdfDictionary ?? resources) : resources;
                VisitContent(PdfEncoding.Latin1GetString(decoded), childResources, invocation.Generated, depth + 1,
                    active, ref pageBytes, ref operations, ref operands);
            } finally { active.Remove(invocation.Stream); }
        }
    }

    internal PdfGeneratedOutputGrowth ToGrowth() {
        long generatedBytes = 0L;
        foreach (var entry in _objects) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (entry.Key < _firstGeneratedObject || entry.Value.Value is not PdfStream stream) continue;
            generatedBytes = checked(generatedBytes + stream.Data.LongLength);
            _maximumRawGeneratedBytes = Math.Max(_maximumRawGeneratedBytes, stream.Data.Length);
            _maximumDecodedGeneratedBytes = Math.Max(_maximumDecodedGeneratedBytes, stream.Data.Length);
        }
        return new PdfGeneratedOutputGrowth(
            minimumRawStreamBytes: _maximumRawGeneratedBytes,
            minimumDecodedStreamBytes: _maximumDecodedGeneratedBytes,
            additionalTotalDecodedStreamBytes: generatedBytes,
            additionalPageContentBytes: _maximumPageBytes,
            additionalRetainedContentBytes: _totalPageBytes,
            additionalContentOperations: _maximumOperations,
            additionalContentOperands: _maximumOperands,
            additionalContentNestingDepth: _contexts == 0 ? 0 : 1,
            additionalPrintProductionContexts: _contexts,
            additionalPrintProductionOperations: (int)Math.Min(int.MaxValue, _operations * 2L));
    }
}
