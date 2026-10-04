using System.Text.Json;

namespace OfficeIMO.AI;

public sealed partial class OfficeAiEngine {
    private sealed record Batch(OfficeAiExecutionRequest Request, IReadOnlyDictionary<string, OfficeAiEvidence> Evidence,
        IReadOnlyDictionary<string, OfficeAiImage> Images, IReadOnlyList<string> Ids,
        IReadOnlyDictionary<string, EvidenceSlice> Slices);
    private sealed record EvidenceSlice(string OriginalId, int Start, int Length);
    private sealed record Plan(IReadOnlyList<Batch> Batches, IReadOnlyList<string> Omitted, IReadOnlyList<int> EmptyPages);
    private sealed record PlanningItem(OfficeAiEvidence? Text, OfficeAiImage? Image) {
        public string Id => Text?.Id ?? Image!.Id;
    }

    private Plan Prepare(OfficeAiDocument document, OfficeAiRequest request, OfficeAiExecutionProfile profile, string requestId, CancellationToken token) {
        var pageScope = request.Pages.ToHashSet();
        var idScope = request.EvidenceIds.ToHashSet(StringComparer.Ordinal);
        if (pageScope.Any(page => !document.Pages.Contains(page))) throw new ArgumentException("A selected page is absent from the captured document.");
        bool InPage(int? page) => pageScope.Count == 0 || (page.HasValue && pageScope.Contains(page.Value));
        var text = document.Evidence.Where(item => InPage(item.Page)).ToArray();
        var images = request.IncludeImages ? document.Images.Where(item => InPage(item.Page)).ToArray() : Array.Empty<OfficeAiImage>();
        var allowedIds = text.Select(item => item.Id).Concat(images.Select(item => item.Id)).ToHashSet(StringComparer.Ordinal);
        if (idScope.Any(id => !allowedIds.Contains(id))) throw new ArgumentException("An evidence selection is outside the permitted page/image scope.");
        text = text.Where(item => idScope.Count == 0 || idScope.Contains(item.Id)).ToArray();
        images = images.Where(item => idScope.Count == 0 || idScope.Contains(item.Id)).ToArray();
        if (images.Length > 0 && !profile.SupportsImages)
            throw new NotSupportedException("The selected evidence requires a vision profile or locally recognized text.");
        var pagesWithEvidence = text.Where(item => item.Page.HasValue).Select(item => item.Page!.Value)
            .Concat(images.Select(item => item.Page)).ToHashSet();
        int[] emptyPages = (pageScope.Count > 0 ? pageScope : document.Pages.ToHashSet())
            .Where(page => !pagesWithEvidence.Contains(page))
            .OrderBy(page => page).ToArray();
        // Explicit block selections do not imply that other pages were requested.
        if (idScope.Count > 0) emptyPages = Array.Empty<int>();
        int maxCharacters = Math.Min(profile.MaxRequestCharacters, request.Limits.MaxRequestCharacters);
        var batches = new List<Batch>();
        var omitted = new List<string>();
        string outputSchema = CreateOutputSchema(request);
        OfficeAiExecutionRequest? CreateRequest(IReadOnlyList<PlanningItem> items) {
            OfficeAiEvidence[] observations = items.Where(item => item.Text is not null).Select(item => item.Text!).ToArray();
            OfficeAiImage[] attachments = items.Where(item => item.Image is not null).Select(item => item.Image!).ToArray();
            if (attachments.Sum(image => (long)image.ByteLength) > Math.Min(profile.MaxImageBytes, request.Limits.MaxImageBytes)
                || attachments.Sum(image => (long)image.Width * image.Height) > request.Limits.MaxImagePixels) return null;
            string input = JsonSerializer.Serialize(new {
                schema = "officeimo.ai.request.v1", sourceHash = document.SourceHash, snapshotHash = document.SnapshotHash, pageProvenance = document.PageProvenance,
                operation = request.Operation.ToString(), instruction = request.Instruction,
                conversationContext = request.ConversationContext,
                resultLimits = new { maxResultItems = request.Limits.MaxResultItems, maxTableCells = request.Limits.MaxTableCells,
                    maxTableColumns = request.Limits.MaxTableColumns },
                fields = request.Fields.Select((field, index) => new { key = FieldKey(index), name = field.Name, type = field.Type.ToString(), dateFormat = field.DateFormat }),
                evidence = observations.Select(item => new { id = item.Id, kind = item.Kind, text = item.Text, page = item.Page,
                    sourceBlockId = item.SourceBlockId, sourceAnchor = item.SourceAnchor, sourceLocation = item.SourceLocation }),
                images = attachments.Select(item => new { id = item.Id, page = item.Page, width = item.Width, height = item.Height })
            });
            return new(requestId + "-" + (batches.Count + 1), Instructions, input, outputSchema,
                Array.AsReadOnly(attachments), request.Limits.MaxResponseCharacters);
        }
        void AddBatch(IReadOnlyList<PlanningItem> items, OfficeAiExecutionRequest execution, EvidenceSlice? slice = null) {
            batches.Add(new(execution,
                items.Where(item => item.Text is not null).ToDictionary(item => item.Id, item => item.Text!, StringComparer.Ordinal),
                items.Where(item => item.Image is not null).ToDictionary(item => item.Id, item => item.Image!, StringComparer.Ordinal),
                Array.AsReadOnly(items.Select(item => slice?.OriginalId ?? item.Id).ToArray()),
                slice is null ? new Dictionary<string, EvidenceSlice>(StringComparer.Ordinal)
                    : new Dictionary<string, EvidenceSlice>(StringComparer.Ordinal) { [items[0].Id] = slice }));
        }
        if (FindFittingPrefix(1, maxCharacters, _ => CreateRequest(Array.Empty<PlanningItem>()), token).Count == 0)
            throw new ArgumentException("Instructions and schema exceed the execution profile's request limit. For table parsing, reduce MaxTableColumns or increase the request budget.");
        var textByPage = text.ToLookup(item => item.Page ?? 0);
        var imagesByPage = images.ToLookup(item => item.Page);
        // Keep page images next to the page's native observations, in the established source order.
        PlanningItem[] ordered = text.Select(item => item.Page ?? 0).Concat(images.Select(item => item.Page)).Distinct()
            .SelectMany(page => textByPage[page].Select(item => new PlanningItem(item, null))
                .Concat(imagesByPage[page].Select(item => new PlanningItem(null, item)))).ToArray();
        int position = 0;
        while (position < ordered.Length) {
            token.ThrowIfCancellationRequested();
            if (batches.Count >= request.Limits.MaxRequests) { omitted.AddRange(ordered.Skip(position).Select(item => item.Id)); break; }
            PackedPrefix packed = FindFittingPrefix(ordered.Length - position, maxCharacters,
                length => CreateRequest(new ArraySegment<PlanningItem>(ordered, position, length)), token);
            if (packed.Count > 0) {
                AddBatch(new ArraySegment<PlanningItem>(ordered, position, packed.Count), packed.Request!);
                position += packed.Count;
                continue;
            }
            PlanningItem oversized = ordered[position++];
            if (oversized.Text is OfficeAiEvidence item) {
                // Split using actual serialized/transport size and keep original offsets outside model control.
                int offset = 0;
                while (offset < item.Text.Length) {
                    token.ThrowIfCancellationRequested();
                    if (batches.Count >= request.Limits.MaxRequests) { omitted.Add(item.Id); break; }
                    string sliceId = item.Id + "@" + offset;
                    int low = 0, high = Math.Min(item.Text.Length - offset, maxCharacters);
                    while (low < high) {
                        int length = low + (high - low + 1) / 2;
                        if (FindFittingPrefix(1, maxCharacters, _ => CreateRequest(new[] {
                            new PlanningItem(item with { Id = sliceId, Text = item.Text.Substring(offset, length) }, null)
                        }), token).Count > 0) low = length;
                        else high = length - 1;
                    }
                    int take = NaturalBoundary(item.Text, offset, low);
                    if (take == 0) { omitted.Add(item.Id); break; }
                    var fragment = new[] { new PlanningItem(item with { Id = sliceId, Text = item.Text.Substring(offset, take) }, null) };
                    PackedPrefix measured = FindFittingPrefix(1, maxCharacters, _ => CreateRequest(fragment), token);
                    if (measured.Count == 0) { omitted.Add(item.Id); break; }
                    AddBatch(fragment, measured.Request!, new(item.Id, offset, take));
                    offset += take;
                }
            } else omitted.Add(oversized.Id);
        }
        return new Plan(batches.AsReadOnly(), Array.AsReadOnly(omitted.Distinct(StringComparer.Ordinal).ToArray()), Array.AsReadOnly(emptyPages));
    }
    private static int NaturalBoundary(string text, int start, int maximum) {
        int end = start + maximum;
        if (end < text.Length && end > start && char.IsHighSurrogate(text[end - 1]) && char.IsLowSurrogate(text[end])) end--;
        if (end == text.Length) return end - start;
        // Prefer a paragraph/sentence/word boundary without throwing away most of the window.
        int minimum = start + (end - start) * 3 / 4;
        for (int index = end - 1; index >= minimum; index--)
            if (char.IsWhiteSpace(text[index])) return index + 1 - start;
        return end - start;
    }
}
