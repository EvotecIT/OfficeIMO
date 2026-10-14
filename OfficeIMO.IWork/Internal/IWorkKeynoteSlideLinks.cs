using System.Globalization;

namespace OfficeIMO.IWork.Internal;

/// <summary>Resolves the qualified native slide UUID and next-slide targets into projected slide positions.</summary>
internal sealed class IWorkKeynoteSlideLinks {
    private readonly SortedDictionary<Guid, int> _nativeTargets;
    private readonly SortedDictionary<int, int> _projectedPositions;

    private IWorkKeynoteSlideLinks(SortedDictionary<Guid, int> nativeTargets,
        SortedDictionary<int, int> projectedPositions) {
        _nativeTargets = nativeTargets;
        _projectedPositions = projectedPositions;
    }

    internal string? Resolve(string? target, int sourceSlidePosition) {
        int position;
        if (TryNativeIdentifier(target, out Guid identifier)
            && _nativeTargets.TryGetValue(identifier, out position) && position > 0
            || target == "?slide=next" && sourceSlidePosition < int.MaxValue
            && _projectedPositions.TryGetValue(sourceSlidePosition + 1, out position)) {
            return "#slide-" + position.ToString(CultureInfo.InvariantCulture);
        }
        // Keep an unresolved native target visible to destination preflight; never silently clear it.
        return target;
    }

    internal static IWorkKeynoteSlideLinks Create(IWorkSourceDocument source,
        IReadOnlyList<IWorkArchiveRecord> nodes, IReadOnlyList<int> nodePositions,
        IReadOnlyList<IWorkKeynoteSlide> slides, IWorkProjectionBudget budget,
        IWorkSourceReferenceIssueCollector references) {
        var positions = new SortedDictionary<int, int>();
        for (int index = 0; index < slides.Count; index++) positions.Add(slides[index].Index, index + 1);
        var targets = new SortedDictionary<Guid, int>();
        var wanted = new SortedSet<Guid>();
        foreach (string? target in Hyperlinks(slides)) {
            source.CancellationToken.ThrowIfCancellationRequested();
            if (TryNativeIdentifier(target, out Guid identifier)) wanted.Add(identifier);
        }
        var result = new IWorkKeynoteSlideLinks(targets, positions);
        if (wanted.Count == 0) return result;

        var selectedNodes = new SortedDictionary<ulong, int>();
        for (int index = 0; index < nodes.Count; index++) {
            source.CancellationToken.ThrowIfCancellationRequested();
            if (selectedNodes.ContainsKey(nodes[index].Identifier)) selectedNodes[nodes[index].Identifier] = 0;
            else selectedNodes.Add(nodes[index].Identifier,
                positions.TryGetValue(nodePositions[index], out int position) ? position : 0);
        }
        IWorkArchiveRecord? metadata = source.Index.UniqueOfType(11006, out _);
        if (metadata == null) return result;
        try {
            IWorkWireMessage message = source.Index.Message(metadata);
            if (message.HasUnexpectedWireKind(3, IWorkWireKind.Bytes)) return Invalid("3");
            budget.AddDrawableReferences(message.FieldCount(3));
            var observed = new SortedSet<Guid>();
            foreach (IWorkWireValue componentValue in message.EnumerateValues(3)) {
                source.CancellationToken.ThrowIfCancellationRequested();
                IWorkWireMessage component = message.ParseNestedMessage(componentValue.Bytes!);
                if (component.HasUnexpectedWireKind(11, IWorkWireKind.Bytes)) return Invalid("3/11");
                budget.AddDrawableReferences(component.FieldCount(11));
                foreach (IWorkWireValue mapValue in component.EnumerateValues(11)) {
                    source.CancellationToken.ThrowIfCancellationRequested();
                    IWorkWireMessage map = component.ParseNestedMessage(mapValue.Bytes!);
                    if (map.TotalFieldCount != 2 || map.FieldCount(1) != 1 || map.FieldCount(2) != 1
                        || map.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                        || map.HasUnexpectedWireKind(2, IWorkWireKind.Bytes)
                        || map.GetUnsigned(1) is not ulong objectIdentifier) return Invalid("3/11");
                    IWorkWireMessage uuid = map.ParseNestedMessage(map.GetBytes(2)!);
                    if (uuid.TotalFieldCount != 2 || uuid.FieldCount(1) != 1 || uuid.FieldCount(2) != 1
                        || uuid.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
                        || uuid.HasUnexpectedWireKind(2, IWorkWireKind.Varint)
                        || uuid.GetUnsigned(1) is not ulong lower || uuid.GetUnsigned(2) is not ulong upper)
                        return Invalid("3/11/2");
                    Guid identifier = Identifier(lower, upper);
                    if (!wanted.Contains(identifier)) continue;
                    int position = selectedNodes.TryGetValue(objectIdentifier, out int selected) ? selected : 0;
                    targets[identifier] = observed.Add(identifier) ? position : 0;
                }
            }
            return result;
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            return Invalid("3/11");
        }

        IWorkKeynoteSlideLinks Invalid(string path) {
            targets.Clear();
            references.Declarations.Record(metadata, path, null, IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            return result;
        }
    }

    private static bool TryNativeIdentifier(string? value, out Guid identifier) {
        identifier = default;
        const string prefix = "?slideid=";
        return value?.Length == prefix.Length + 36 && value.StartsWith(prefix, StringComparison.Ordinal)
            && Guid.TryParseExact(value.Substring(prefix.Length), "D", out identifier);
    }

    private static Guid Identifier(ulong lower, ulong upper) {
        // TSP UUID words hold consecutive little-endian bytes, unlike the CFUUID formula representation.
        char[] text = new char[32];
        WriteWord(lower, 0);
        WriteWord(upper, 16);
        return Guid.ParseExact(new string(text), "N");

        void WriteWord(ulong word, int offset) {
            const string digits = "0123456789abcdef";
            for (int index = 0; index < 8; index++) {
                byte value = (byte)(word >> (index * 8));
                text[offset + index * 2] = digits[value >> 4];
                text[offset + index * 2 + 1] = digits[value & 15];
            }
        }
    }

    private static IEnumerable<string?> Hyperlinks(IReadOnlyList<IWorkKeynoteSlide> slides) {
        foreach (IWorkKeynoteSlide slide in slides) {
            foreach (IWorkTextBox box in slide.TextBoxes.Concat(slide.TitleBox == null
                         ? Array.Empty<IWorkTextBox>() : new[] { slide.TitleBox })) {
                yield return box.Hyperlink;
                foreach (IWorkTextRun run in box.Content.Paragraphs.SelectMany(paragraph => paragraph.Runs))
                    yield return run.Hyperlink;
            }
            foreach (IWorkImageAsset image in slide.Images) yield return image.Hyperlink;
            foreach (IWorkTextRun run in slide.Tables.SelectMany(table => table.Cells)
                         .Where(cell => cell.RichText != null).SelectMany(cell => cell.RichText!.Paragraphs)
                         .Concat(slide.PresenterNoteContent.Paragraphs).SelectMany(paragraph => paragraph.Runs))
                yield return run.Hyperlink;
        }
    }
}
