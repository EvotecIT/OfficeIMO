using System.Text;
using OfficeIMO.Pdf;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig.DocumentLayoutAnalysis.ReadingOrderDetector;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static partial class PdfBenchmarkValidation {
    // This validates the fixed report's existing parsed tag objects, not arbitrary PDF syntax.
    private static void ValidateHtmlLogicalSourceOrder(byte[] bytes, PdfBenchmarkScenario scenario, string engine) {
        PdfReadDocument source = PdfReadDocument.Open(bytes);
        PdfTaggedContentInfo tagged = source.TaggedContent
            ?? throw new InvalidDataException($"{engine} has no report structure tree.");
        int objectLimit = checked(1024 + scenario.PageCount * (scenario.RowsPerPage + 1) * 64);
        if (objectLimit > 100_000) throw Reject("the fixture exceeds the 100000-object admission limit");
        PdfRawDocumentView raw = source.RawStructure(new PdfRawStructureOptions {
            MaxObjects = objectLimit, MaxCollectionItems = objectLimit, MaxDepth = 32
        });
        if (raw.IsTruncated) throw Reject("the bounded raw structure projection is truncated");
        var objects = raw.Objects.ToDictionary(item => item.ObjectNumber);
        if (objects.Values.Any(item => Name(item.Value, "Subtype") == "Form"))
            throw Reject("Form-XObject MCID scope is unsupported by this admission gate");
        int rootId = tagged.StructTreeRootObjectNumber
            ?? throw Reject("an indirect structure-tree root is required");
        PdfRawValue root = Object(rootId).Value;
        if (Name(root, "Type") != "StructTreeRoot") throw Reject("the structure root is invalid");
        var seen = new HashSet<int>();
        var roots = new HtmlReportTag("StructTreeRoot", rootId, null, null);
        ReadKids(Entry(root, "K"), roots, 0);
        HtmlReportTag[] rootChildren = roots.Children.ToArray();
        if (rootChildren.Length != 1 || rootChildren[0].Role != "Document"
            || !seen.SetEquals(tagged.StructureElements.Select(item => item.ObjectNumber)))
            throw Reject("the document structure is not one complete rooted tree");

        using UglyToad.PdfPig.PdfDocument independent = UglyToad.PdfPig.PdfDocument.Open(bytes);
        if (independent.NumberOfPages != scenario.PageCount || source.Pages.Count != scenario.PageCount)
            throw Reject("the page inventory differs from the authored report");
        var marked = new Dictionary<(int Page, int Mcid), MarkedContentElement>();
        var pageLetters = new List<IReadOnlyList<Letter>>(scenario.PageCount);
        for (int pageNumber = 1; pageNumber <= independent.NumberOfPages; pageNumber++) {
            var page = independent.GetPage(pageNumber);
            int pageObject = source.Pages[page.Number - 1].ObjectNumber;
            pageLetters.Add(page.Letters);
            var pending = new Stack<MarkedContentElement>(page.GetMarkedContents());
            var visited = new HashSet<MarkedContentElement>(ReferenceEqualityComparer.Instance);
            while (pending.TryPop(out var item)) {
                if (!visited.Add(item)) continue;
                int? mcid = item.MarkedContentIdentifier;
                if (mcid is >= 0 && !marked.TryAdd((pageObject, mcid.Value), item))
                    throw Reject("a page MCID has ambiguous marked-content ownership");
                foreach (var child in item.Children) pending.Push(child);
            }
        }
        var claimedContent = new HashSet<(int Page, int Mcid)>();
        var claimedLetters = new HashSet<Letter>(ReferenceEqualityComparer.Instance);
        var wordExtractor = new NearestNeighbourWordExtractor(
            new NearestNeighbourWordExtractor.NearestNeighbourWordExtractorOptions { MaxDegreeOfParallelism = 3 });
        var pageSegmenter = new DefaultPageSegmenter(
            new DefaultPageSegmenter.DefaultPageSegmenterOptions { MaxDegreeOfParallelism = 3 });
        using IEnumerator<HtmlReportTag> owners = Owners(rootChildren[0]).GetEnumerator();
        for (int page = 1; page <= scenario.PageCount; page++) {
            int pageObject = source.Pages[page - 1].ObjectNumber;
            BindText(Next("H1"), scenario.PageTitle(page), pageObject);
            HtmlReportTag figure = Next("Figure");
            if (figure.AlternateText != $"Benchmark report illustration for page {page.ToString(System.Globalization.CultureInfo.InvariantCulture)}")
                throw Reject("a Figure does not identify its authored illustration");
            if (LetterChunks(figure, pageObject).SelectMany(letters => letters).Any(letter => !string.IsNullOrWhiteSpace(letter.Value)))
                throw Reject("the report illustration unexpectedly owns text");
            for (int paragraph = 0; paragraph < scenario.ParagraphsPerPage; paragraph++)
                BindText(Next("P"), scenario.Narrative(page, paragraph), pageObject);
            HtmlReportTag table = Next("Table");
            HtmlReportTag[] rows = TableParts(table, "TR").ToArray();
            IReadOnlyList<string[]> expected = scenario.TableRows(page);
            if (rows.Length != expected.Count) throw Reject("the tagged table row inventory differs");
            for (int row = 0; row < rows.Length; row++) {
                HtmlReportTag[] cells = TableParts(rows[row], row == 0 ? "TH" : "TD").ToArray();
                if (cells.Length != expected[row].Length) throw Reject("the tagged table cell inventory differs");
                for (int column = 0; column < cells.Length; column++) BindText(cells[column], expected[row][column], pageObject);
            }
            BindTablePaint(table, pageObject);
        }
        if (owners.MoveNext()) throw Reject("the structure contains extra authored content owners");
        if (Descendants(roots).Sum(node => node.OrderedKids.Count(kid => kid is HtmlReportContent)) != claimedContent.Count)
            throw Reject("the tagged source contains unowned or repeated marked-content references");
        if (pageLetters.SelectMany(value => value).Any(letter => !string.IsNullOrWhiteSpace(letter.Value)
            && !claimedLetters.Contains(letter)))
            throw Reject("visible decoded report text is not owned exactly once by the tagged source");

        HtmlReportTag Next(string role) {
            if (!owners.MoveNext() || owners.Current.Role != role) throw Reject("the authored semantic-owner order differs");
            return owners.Current;
        }
        void BindText(HtmlReportTag owner, string expected, int pageObject) {
            var actual = new StringBuilder();
            // /K determines the sequence of content chunks. Geometry orders only
            // the letters inside one MCID; it must not undo a sibling Span swap.
            foreach (IReadOnlyList<Letter> letters in LetterChunks(owner, pageObject)) {
                var lines = ReadingOrderHelper.OrderByReadingOrder(pageSegmenter.GetBlocks(wordExtractor.GetWords(letters))
                    .SelectMany(block => block.TextLines));
                actual.Append(string.Concat(lines.SelectMany(line => ReadingOrderHelper.OrderByReadingOrder(line.Words))
                    .Select(word => word.Text)));
                foreach (Letter letter in letters.Where(value => !string.IsNullOrWhiteSpace(value.Value)))
                    if (!claimedLetters.Add(letter)) throw Reject("a decoded report glyph is owned by multiple source elements");
            }
            if (!string.Equals(WithoutWhitespace(actual.ToString()), WithoutWhitespace(expected), StringComparison.Ordinal))
                throw Reject($"the tagged {owner.Role} does not retain its exact authored page/text identity");
        }
        IEnumerable<IReadOnlyList<Letter>> LetterChunks(HtmlReportTag owner, int pageObject) {
            int references = 0;
            foreach (HtmlReportContent reference in ContentInOrder(owner)) {
                references++;
                yield return ClaimContent(reference, pageObject);
            }
            if (references == 0) throw Reject("an authored source owner has no marked-content binding");
        }
        IReadOnlyList<Letter> ClaimContent(HtmlReportContent reference, int pageObject) {
            var key = (reference.Page, reference.Mcid);
            if (key.Page != pageObject || !claimedContent.Add(key) || !marked.TryGetValue(key, out var content))
                throw Reject("a source element has missing, repeated, or wrong-page MCID content");
            var letters = new HashSet<Letter>(ReferenceEqualityComparer.Instance);
            var pending = new Stack<MarkedContentElement>();
            var visited = new HashSet<MarkedContentElement>(ReferenceEqualityComparer.Instance);
            pending.Push(content);
            while (pending.TryPop(out var item)) {
                if (!visited.Add(item)) continue;
                foreach (Letter letter in item.Letters) letters.Add(letter);
                foreach (var child in item.Children) pending.Push(child);
            }
            return letters.ToArray();
        }
        void BindTablePaint(HtmlReportTag node, int pageObject) {
            // Table and row backgrounds/borders may have their own non-text
            // MCIDs. Cell text still needs its exact authored owner and order.
            foreach (HtmlReportKid kid in node.OrderedKids) {
                if (kid is HtmlReportContent content) {
                    if (ClaimContent(content, pageObject).Any(letter => !string.IsNullOrWhiteSpace(letter.Value)))
                        throw Reject("a table container owns decoded text outside an authored cell");
                } else if (kid is HtmlReportChild child && child.Value.Role is not ("TH" or "TD")) {
                    BindTablePaint(child.Value, pageObject);
                }
            }
        }
        IEnumerable<HtmlReportContent> ContentInOrder(HtmlReportTag owner) {
            foreach (HtmlReportKid kid in owner.OrderedKids) {
                if (kid is HtmlReportContent content) yield return content;
                else if (kid is HtmlReportChild child)
                    foreach (var descendant in ContentInOrder(child.Value)) yield return descendant;
            }
        }
        IEnumerable<HtmlReportTag> Owners(HtmlReportTag node) {
            if (node.Role is "H1" or "Figure" or "Table" || node.Role == "P" && !Descendants(node).Any(item => item.Role == "Figure")) {
                yield return node; yield break;
            }
            if (node.Role is not ("Document" or "Sect" or "Div" or "Span" or "P" or "Part" or "Art" or "NonStruct"))
                throw Reject("an unsupported report wrapper role was encountered: " + node.Role);
            foreach (var child in node.Children) foreach (var owner in Owners(child)) yield return owner;
        }
        IEnumerable<HtmlReportTag> TableParts(HtmlReportTag node, string target) {
            foreach (var child in node.Children) {
                if (child.Role == target) { yield return child; continue; }
                if (child.Role is not ("THead" or "TBody" or "TFoot" or "Div" or "Span"))
                    throw Reject("the tagged table contains an unexpected role: " + child.Role);
                foreach (var part in TableParts(child, target)) yield return part;
            }
        }
        IEnumerable<HtmlReportTag> Descendants(HtmlReportTag node) {
            yield return node;
            foreach (var child in node.Children) foreach (var descendant in Descendants(child)) yield return descendant;
        }
        void ReadKids(PdfRawValue? value, HtmlReportTag parent, int depth) {
            if (value == null || value.Kind == PdfRawValueKind.Null) return;
            if (depth > 32 || value.IsTruncated) throw Reject("the report structure nesting/projection limit was reached");
            int? id = null;
            if (value.Kind == PdfRawValueKind.Reference) {
                id = value.ReferenceObjectNumber;
                var indirect = Object(id ?? throw Reject("a structure reference lacks an object number"));
                if (indirect.Generation != value.ReferenceGeneration) throw Reject("a structure reference generation differs");
                value = indirect.Value;
            }
            if (value.IsTruncated) throw Reject("a report structure value is truncated");
            if (value.Kind == PdfRawValueKind.Array) {
                foreach (var item in value.Items) ReadKids(item, parent, depth + 1);
                return;
            }
            if (Name(value, "Type") == "StructElem") {
                if (id.HasValue && !seen.Add(id.Value)) throw Reject("the report structure repeats an element or contains a cycle");
                if (Reference(value, "P") != parent.ObjectNumber) throw Reject("a structure child's declared parent differs");
                string role = Name(value, "S") ?? throw Reject("a structure element has no role");
                if (tagged.RoleMap.TryGetValue(role, out string? mapped)) role = mapped;
                var node = new HtmlReportTag(role, id, Reference(value, "Pg") ?? parent.PageObject, Text(value, "Alt"));
                parent.OrderedKids.Add(new HtmlReportChild(node));
                ReadKids(Entry(value, "K"), node, depth + 1);
                return;
            }
            int? page = parent.PageObject;
            int? mcid = null;
            if (value.Kind == PdfRawValueKind.Number) mcid = Integer(value);
            else if (value.Kind == PdfRawValueKind.Dictionary && (Name(value, "Type") == "MCR" || Entry(value, "MCID") != null)) {
                if (Entry(value, "Stm") is { Kind: not PdfRawValueKind.Null })
                    throw Reject("explicit content-stream MCID scope is unsupported by this admission gate");
                page = Reference(value, "Pg") ?? page;
                mcid = Integer(Entry(value, "MCID"));
            } else throw Reject("an unsupported or missing report /K reference was encountered");
            if (!page.HasValue || !mcid.HasValue) throw Reject("marked content lacks its page/MCID identity");
            parent.OrderedKids.Add(new HtmlReportContent(page.Value, mcid.Value));
        }
        PdfRawObjectView Object(int id) => objects.TryGetValue(id, out var value) ? value : throw Reject("a structure reference cannot be resolved");
        static PdfRawValue? Entry(PdfRawValue value, string key) => value.Entries.TryGetValue(key, out var entry) ? entry : null;
        static string? Name(PdfRawValue value, string key) => Entry(value, key) is { Kind: PdfRawValueKind.Name } entry ? entry.Text : null;
        static string? Text(PdfRawValue value, string key) => Entry(value, key) is { Kind: PdfRawValueKind.TextString } entry ? entry.Text : null;
        static int? Reference(PdfRawValue value, string key) => Entry(value, key) is { Kind: PdfRawValueKind.Reference } entry ? entry.ReferenceObjectNumber : null;
        static int? Integer(PdfRawValue? value) => value is { Kind: PdfRawValueKind.Number, Number: double number }
            && number >= 0D && number <= int.MaxValue && number == Math.Truncate(number) ? (int)number : null;
        InvalidDataException Reject(string detail) => new($"{engine} tagged report admission failed: {detail}.");
    }

    private abstract record HtmlReportKid;
    private sealed record HtmlReportChild(HtmlReportTag Value) : HtmlReportKid;
    private sealed record HtmlReportContent(int Page, int Mcid) : HtmlReportKid;

    private sealed class HtmlReportTag(string role, int? objectNumber, int? pageObject, string? alternateText) {
        internal string Role { get; } = role;
        internal int? ObjectNumber { get; } = objectNumber;
        internal int? PageObject { get; } = pageObject;
        internal string? AlternateText { get; } = alternateText;
        internal List<HtmlReportKid> OrderedKids { get; } = new();
        internal IEnumerable<HtmlReportTag> Children => OrderedKids.OfType<HtmlReportChild>().Select(kid => kid.Value);
    }
}
