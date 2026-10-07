using OfficeIMO.Epub;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;

if (args.Length != 3) throw new ArgumentException("Supply the pinned Moby-Dick media-overlay EPUB, its expected SHA-256, and a new output directory.");
var input = new FileInfo(args[0]);
if (input.Length > 128L * 1024 * 1024) throw new InvalidDataException("Input exceeds the 128 MiB qualification limit.");
byte[] original = File.ReadAllBytes(input.FullName);
string sourceHash = Convert.ToHexString(SHA256.HashData(original)).ToLowerInvariant();
if (!string.Equals(sourceHash, args[1], StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Input SHA-256 does not match the supplied provenance.");
string output = Path.GetFullPath(args[2]);
if (Directory.Exists(output) || File.Exists(output)) throw new IOException("Output exists; use a new task-owned directory.");
var source = EpubPublication.Load(new MemoryStream(original));
var book = EpubPublication.Load(new MemoryStream(original));
var sourcePayloads = ProducerAssertions.Payloads(original);
var relocation = new Dictionary<string, string>(StringComparer.Ordinal) {
    ["OPS/chapter_001.xhtml"] = "OPS/revised/text/chapter_001.xhtml",
    ["OPS/chapter_001_overlay.smil"] = "OPS/revised/overlays/chapter_001.smil",
    ["OPS/audio/mobydick_001_002_melville.mp4"] = "OPS/revised/audio/narration.mp4"
};
XNamespace html = "http://www.w3.org/1999/xhtml";
string introduction = source.GetContentXml("xintroduction_001").Root!.Element(html + "body")!.Value;
if (source.Manifest.Count(item => item.MediaType == "application/smil+xml") != 2)
    throw new InvalidDataException("Expected the pinned two-overlay sample profile.");
Directory.CreateDirectory(output);
var stages = new List<object>();
var writeOptions = new EpubWriteOptions { ModifiedAt = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero) };
WriteStage("unchanged", original, new Dictionary<string, string>());
book.RenameResource("xchapter_001", relocation["OPS/chapter_001.xhtml"]);
book.RenameResource("chapter_001_overlay", relocation["OPS/chapter_001_overlay.smil"]);
book.RenameResource("chapter_001_audio", relocation["OPS/audio/mobydick_001_002_melville.mp4"]);
WriteStage("renamed-narration", book.Write(writeOptions).Bytes, relocation);
var content = book.GetContentXml("xintroduction_001");
var section = content.Descendants(html + "section").Single();
section.SetAttributeValue("id", "editorial-introduction");
section.Elements(html + "div").First().SetAttributeValue("id", "editorial-boundary");
book.SetContentXml("xintroduction_001", content);
book.SplitChapter("xintroduction_001", "editorial-boundary", "editorial-second", "OPS/revised/introduction-part2.xhtml", "Etymology continued");
if (book.GetContentXml("xintroduction_001").Root!.Element(html + "body")!.Value +
    book.GetContentXml("editorial-second").Root!.Element(html + "body")!.Value != introduction)
    throw new InvalidDataException("Split changed introduction text or its order.");
WriteStage("split-introduction", book.Write(writeOptions).Bytes, relocation);
book.MergeChapters("xintroduction_001", "editorial-second", "editorial-second-start");
if (book.GetContentXml("xintroduction_001").Root!.Element(html + "body")!.Value != introduction)
    throw new InvalidDataException("Merge changed introduction text or its order.");
WriteStage("merged-introduction", book.Write(writeOptions).Bytes, relocation);
Console.WriteLine(Path.Combine(output, "evidence.json"));

void WriteStage(string name, byte[] bytes, IReadOnlyDictionary<string, string> moves) {
    var reopened = EpubPublication.Load(new MemoryStream(bytes));
    object preservation = ProducerAssertions.Verify(source, reopened, sourcePayloads, ProducerAssertions.Payloads(bytes), moves);
    var preflight = reopened.Preflight();
    if (preflight.Checks.Single(check => check.Code == "media-overlays").Status != EpubPreflightStatus.Passed)
        throw new InvalidDataException("Edited sample failed media-overlay preflight.");
    File.WriteAllBytes(Path.Combine(output, name + ".epub"), bytes);
    stages.Add(new { name, sha256 = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(), preservation,
        nativePreflight = new { preflight.HasErrors, preflight.HasUncheckedItems,
            checks = preflight.Checks.Select(check => new { check.Code, status = check.Status.ToString(), check.Diagnostics }) } });
    File.WriteAllText(Path.Combine(output, "evidence.json"), JsonSerializer.Serialize(new {
        sourceHash, completed = stages.Count == 4, upstreamRevisionAuthenticatedByRunner = false, expectedSourceProfile = "IDPF/epub3-samples 7651e2002b631e6577fadf7e9e0692fa6efb8746 / 30/moby-dick-mo",
        assertions = "Audio and other non-XML assets stay byte-identical; SMIL structure/timing and existing TOC targets survive relocation; chapter body text survives editing, including introduction split/merge.",
        stages, independentValidation = "not-performed-by-this-runner", nativePlayback = "not-performed"
    }, new JsonSerializerOptions { WriteIndented = true }));
}
