using System.Text.Json;
using OfficeIMO.Core.Internal;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Opt-in wire feasibility through the canonical source codecs, without promoting a public edit profile.</summary>
internal static class SourceWrapOrder {
    internal static int Prepare(string input, string output, string specification) {
        if (new FileInfo(input).Length > 64 * 1024 * 1024 || new FileInfo(specification).Length > 1024 * 1024)
            throw new InvalidDataException("Wrap-order probe input exceeds its byte limit.");
        var spec = JsonSerializer.Deserialize<Specification>(File.ReadAllBytes(specification), Program.JsonOptions)
            ?? throw new InvalidDataException("Missing wrap-order specification.");
        if (spec.ObjectIds == null || spec.FrameWraps == null || spec.ObjectIds.Length == 0)
            throw new InvalidDataException("The probe requires an existing page permutation and explicit frame relationships.");
        byte[] bytes = File.ReadAllBytes(input);
        if (!OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error))
            throw new InvalidDataException(error);
        var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
        var source = new PublisherNativeSource(bytes, compound!);
        if (!source.CanRewrite(context.Options.Limits, context.Token)) throw new NotSupportedException("Protected or invalid native source.");
        PublisherContents contents = new PublisherContentsReader(new PublisherBinaryData(source.ContentsBytes, "Contents"), context).Read();
        if (!contents.ObjectPages.TryGetValue(spec.ObjectIds[0], out uint page)
            || !contents.ObjectPages.Where(pair => pair.Value == page).Select(pair => pair.Key).OrderBy(id => id)
                .SequenceEqual(spec.ObjectIds.OrderBy(id => id)) || spec.ObjectIds.Distinct().Count() != spec.ObjectIds.Length)
            throw new InvalidDataException("The probe requires every existing object on one page exactly once.");
        var wraps = spec.FrameWraps.ToDictionary(pair => pair.Key, pair => (IReadOnlyList<uint>)pair.Value);
        if (wraps.Keys.Any(id => !contents.ObjectPages.TryGetValue(id, out uint owner) || owner != page))
            throw new InvalidDataException("The probe frame is outside the selected page.");
        byte[] rewritten = PublisherContentsSourceWriter.Rewrite(source.ContentsBytes, contents,
            new Dictionary<uint, uint>(), new Dictionary<uint, string>(), new Dictionary<uint, uint?>(), context,
            context.Options.Limits.MaxInputBytes, frameWraps: wraps) ?? throw new NotSupportedException("Unqualified Contents rewrite.");
        byte[] artwork = PublisherEscherOrderWriter.Rewrite(source.EscherBytes,
            new Dictionary<uint, IReadOnlyList<uint>> { [page] = spec.ObjectIds }, context, context.Options.Limits.MaxInputBytes)
            ?? throw new NotSupportedException("Unqualified complete sibling-record ordering.");
        var streams = new Dictionary<string, byte[]> { ["Contents"] = rewritten, ["Escher/EscherStm"] = artwork };
        source.InvalidatePreview(streams, context.Token);
        byte[] result = source.Rewrite(streams, context.Options.Limits.MaxInputBytes, context.Token);
        using (var file = new FileStream(output, FileMode.CreateNew, FileAccess.Write, FileShare.Read)) file.Write(result);
        Console.WriteLine(JsonSerializer.Serialize(new { schema = "officeimo.publisher.wrap-order-probe.v1",
            sourceSha256 = NativeInventory.Hash(bytes), outputSha256 = NativeInventory.Hash(result), page,
            objectIds = spec.ObjectIds, frameWraps = spec.FrameWraps, nativeAcceptance = "not-run" }, Program.JsonOptions));
        return 0;
    }

    private sealed class Specification {
        public uint[]? ObjectIds { get; set; }
        public Dictionary<uint, uint[]>? FrameWraps { get; set; }
    }
}
