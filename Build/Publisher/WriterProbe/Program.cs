using System.Text.Json;
using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

internal static class Program {
    internal static readonly JsonSerializerOptions JsonOptions = new() {
        WriteIndented = true,
        Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter() }
    };

    public static int Main(string[] args) {
        try {
            if (args.Length == 3 && args[0] == "table-utility-source") return SourceTableUtilityControl.Prepare(args[1], args[2]);
            if (args.Length == 4 && args[0] == "source-newsletter") return SourceNewsletterLifecycle.Run(args[1], args[2], args[3]);
            if (args.Length == 4 && args[0] == "wrap-order") return SourceWrapOrder.Prepare(args[1], args[2], args[3]);
            if ((args.Length == 3 && args[0] == "preserve") || (args.Length is 8 or 9 && args[0] == "edit")
                || (args.Length == 5 && args[0] is "edit-text" or "edit-story" or "edit-paragraphs"))
                return PrepareLoaded(args);
            if (args.Length is >= 2 and <= 3 && args[0] == "inventory") {
                string json = JsonSerializer.Serialize(NativeInventory.Read(args[1]), JsonOptions);
                if (args.Length == 3) WriteNew(args[2], System.Text.Encoding.UTF8.GetBytes(json));
                else Console.WriteLine(json);
                return 0;
            }
            if (args.Length is >= 3 and <= 5 && args[0] == "create") {
                CandidatePublicationSpec specification = CandidateCases.Create(args[1], args.Length >= 4 ? args[3] : null,
                    args.Length == 5 ? args[4] : null);
                byte[] bytes = CandidateWriter.Create(specification);
                WriteNew(args[2], bytes);
                Console.WriteLine(JsonSerializer.Serialize(new {
                    schema = "officeimo.publisher.writer-probe.candidate.v1",
                    profile = specification.PhysicalPageCount == 1 ? "template-free-one-page-probe" : "template-free-" + args[1] + "-probe",
                    hypothesis = args[0],
                    kind = args[1],
                    rootClassId = CandidateWriter.PublicationClassId,
                    pageGraph = new {
                        orderedReferences = NativePublicationGraph.GetOrderedPageCount(specification),
                        dpgMaster = NativePublicationGraph.GetMasterPageCount(specification),
                        dpgSpecial = NativePublicationGraph.SpecialPageBoundary,
                        firstPhysicalIndex = NativePublicationGraph.GetMasterPageCount(specification),
                        lastPhysicalIndex = NativePublicationGraph.GetOrderedPageCount(specification) - NativePublicationGraph.SpecialPageBoundary
                    },
                    output = Path.GetFullPath(args[2]),
                    bytes = bytes.Length,
                    sha256 = NativeInventory.Hash(bytes),
                    sourcePublication = (string?)null,
                    nativeAcceptance = "not-run",
                    note = "Template-free synthesis hypothesis; acceptance by OfficeIMO does not establish acceptance by Microsoft Publisher."
                }, JsonOptions));
                return 0;
            }
            if (args.Length == 3 && args[0] == "render") {
                Console.WriteLine(JsonSerializer.Serialize(ManagedComparison.Render(args[1], args[2]), JsonOptions));
                return 0;
            }
            if (args.Length is 3 or 4 && args[0] == "diff") {
                string json = JsonSerializer.Serialize(NativeDiff.Compare(args[1], args[2]), JsonOptions);
                if (args.Length == 4) WriteNew(args[3], System.Text.Encoding.UTF8.GetBytes(json));
                else Console.WriteLine(json);
                return 0;
            }
            Console.Error.WriteLine("inventory <input.pub> [report.json]\ndiff <left.pub> <right.pub> [report.json]\ncreate <blank|text|styled-text|shape|image|combined> <candidate.pub> [text-or-image-path]\ncreate <blank-pages|text-pages> <candidate.pub> <page-count>\npreserve <input.pub> <candidate.pub>\nedit <input.pub> <candidate.pub> <object-id> <x> <y> <width> <height> [RRGGBB]\nedit-text <input.pub> <candidate.pub> <story-id> <text>\nedit-story <input.pub> <candidate.pub> <story-id> <runs.json>\nrender <input.pub> <unused-output-directory>");
            Console.Error.WriteLine("edit-paragraphs <input.pub> <candidate.pub> <story-id> <paragraph-run-arrays.json>");
            Console.Error.WriteLine("source-newsletter <generated-newsletter.pub> <replacement-jpeg> <empty-output-directory>");
            Console.Error.WriteLine("wrap-order <input.pub> <candidate.pub> <object-order-and-frame-wraps.json>");
            Console.Error.WriteLine("create <object-pages|mixed-object-pages|master-content|newsletter> <candidate.pub> <image-path>");
            Console.Error.WriteLine("create <linked-story|linked-story-overflow> <candidate.pub>");
            Console.Error.WriteLine("create <table|table-edited|table-merged|table-decorated|table-decorated-edited> <candidate.pub>");
            Console.Error.WriteLine("table-utility-source <decorated-bare-pair-control.pub> <derived-input.pub>");
            Console.Error.WriteLine("create <table-grid|table-grid-empty|table-grid-edited> <candidate.pub> <rows>x<columns>");
            Console.Error.WriteLine("create <frame-layout|frame-layout-edited> <candidate.pub>");
            Console.Error.WriteLine("create <frame-layout-language|linked-story-language|linked-story-language-overflow> <candidate.pub>");
            Console.Error.WriteLine("create <image-pages|named-image-pages|cropped-image-pages|shared-image-pages> <candidate.pub> <first-image-path> <last-image-path>");
            return 2;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception.GetType().Name + ": " + exception.Message);
            return 1;
        }
    }

    private static int PrepareLoaded(string[] args) {
        PublisherDocument document = PublisherDocument.Load(args[1]);
        if (args[0] is "edit-story" or "edit-paragraphs") {
            if (new FileInfo(args[4]).Length > 64 * 1024) throw new ArgumentException("The styled-run input exceeds 64 KiB.");
            uint storyId = uint.Parse(args[3], NumberStyles.None, CultureInfo.InvariantCulture);
            PublisherTextStory story = document.TextStories.Single(item => item.Id == storyId);
            string json = File.ReadAllText(args[4]);
            ProbeRun[][] values = args[0] == "edit-story"
                ? new[] { JsonSerializer.Deserialize<ProbeRun[]>(json) ?? throw new ArgumentException("The styled-run input must be an array.") }
                : JsonSerializer.Deserialize<ProbeRun[][]>(json) ?? throw new ArgumentException("The paragraph input must be an array of run arrays.");
            if (values.Length != story.Paragraphs.Count) throw new ArgumentException("Retain the original paragraph count.");
            var paragraphs = values.Select((items, paragraphIndex) => {
                OfficeRichTextParagraph paragraph = story.Paragraphs[paragraphIndex];
                var runs = items.Select((value, index) => {
                    OfficeRichTextRun original = paragraph.Runs[Math.Min(index, paragraph.Runs.Count - 1)];
                    return new OfficeRichTextRun(value.Text, value.FontSize, original.Color, value.Bold, original.Italic, original.Underline,
                        original.FontFamily, original.Strikethrough, original.BackgroundColor, original.UnderlineStyle,
                        original.StrikethroughStyle, original.Baseline, original.ParagraphIndent);
                }).ToArray();
                return new OfficeRichTextParagraph(runs, paragraph.Alignment, paragraph.LineHeight,
                    paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor).WithTabStops(paragraph.TabStops);
            }).ToArray();
            document.ReplaceStoryParagraphs(storyId, paragraphs);
        }
        if (args[0] == "edit-text") {
            uint storyId = uint.Parse(args[3], NumberStyles.None, CultureInfo.InvariantCulture);
            PublisherTextStory story = document.TextStories.Single(item => item.Id == storyId);
            OfficeRichTextParagraph paragraph = story.Paragraphs.Single();
            OfficeRichTextRun run = paragraph.Runs.Single();
            document.ReplaceStoryParagraphs(storyId, new[] { new OfficeRichTextParagraph(new[] {
                new OfficeRichTextRun(args[4], run.FontSize, run.Color, run.Bold, run.Italic, run.Underline,
                    run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle, run.StrikethroughStyle,
                    run.Baseline, run.ParagraphIndent)
            }, paragraph.Alignment, paragraph.LineHeight, paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor).WithTabStops(paragraph.TabStops) });
        }
        if (args[0] == "edit") {
            uint id = uint.Parse(args[3], NumberStyles.None, CultureInfo.InvariantCulture);
            double Number(int index) => double.Parse(args[index], NumberStyles.Float, CultureInfo.InvariantCulture);
            double x = Number(4), y = Number(5), width = Number(6), height = Number(7);
            PublisherPage[] pages = document.Pages.Concat(document.MasterPages).ToArray();
            if (pages.SelectMany(page => page.TextFrames).Any(frame => frame.Id == id)) document.SetTextFrameBounds(id, x, y, width, height);
            else if (pages.SelectMany(page => page.Pictures).Any(picture => picture.Id == id)) document.SetPictureBounds(id, x, y, width, height);
            else if (pages.SelectMany(page => page.Shapes).Any(shape => shape.Id == id)) document.SetShapeBounds(id, x, y, width, height);
            else if (pages.SelectMany(page => page.Tables).Any(table => table.Id == id)) document.SetTableBounds(id, x, y, width, height);
            else throw new ArgumentException("The source has no placed object with that publication identifier.");
            if (args.Length == 9) {
                if (args[8].Length != 6) throw new ArgumentException("A fill requires exactly six RGB hexadecimal digits.");
                uint rgb = uint.Parse(args[8], NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                PublisherShape shape = document.Pages.Concat(document.MasterPages).SelectMany(page => page.Shapes).Single(item => item.Id == id);
                OfficeShape prior = shape.CreateShape();
                document.SetShapePaint(id, OfficeColor.FromRgb((byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb),
                    prior.StrokeColor, prior.StrokeWidth, prior.FillOpacity ?? 1, prior.StrokeOpacity ?? 1);
            }
        }
        PublisherNativePreparation prepared = document.PrepareNativeBytes(64L * 1024 * 1024);
        byte[]? bytes = prepared.IsPrepared ? prepared.CopyBytes() : null;
        if (bytes != null) WriteNew(args[2], bytes);
        Console.WriteLine(JsonSerializer.Serialize(new {
            schema = "officeimo.publisher.writer-probe.source-edit.v1",
            operation = args[0], prepared = prepared.IsPrepared,
            output = bytes == null ? null : Path.GetFullPath(args[2]),
            bytes = prepared.ByteLength, sha256 = bytes == null ? null : NativeInventory.Hash(bytes),
            diagnostics = prepared.FidelityDiagnostics,
            styleDependencies = prepared.StyleDependencies,
            nativeAcceptance = "unassessed",
            repairFreeOpening = "unassessed"
        }, JsonOptions));
        return prepared.IsPrepared ? 0 : 1;
    }

    private static void WriteNew(string path, byte[] bytes) {
        using var stream = new FileStream(Path.GetFullPath(path), FileMode.CreateNew, FileAccess.Write);
        stream.Write(bytes);
    }
    private sealed record ProbeRun(string Text, double FontSize, bool Bold);
}
