using System.Text.Json;

namespace OfficeIMO.AI;

public sealed partial class OfficeAiEngine {
    private sealed record Batch(OfficeAiExecutionRequest Request, IReadOnlyDictionary<string, OfficeAiEvidence> Evidence,
        IReadOnlyDictionary<string, OfficeAiImage> Images, IReadOnlyList<string> Ids);
    private sealed record Plan(IReadOnlyList<Batch> Batches, IReadOnlyList<string> Omitted, IReadOnlyList<int> EmptyPages);

    private Plan Prepare(OfficeAiDocument document, OfficeAiRequest request, OfficeAiExecutionProfile profile, string requestId) {
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
        int[] emptyPages = (pageScope.Count > 0 ? pageScope : document.Pages.ToHashSet())
            .Where(page => !text.Any(item => item.Page == page) && !images.Any(item => item.Page == page))
            .OrderBy(page => page).ToArray();
        // Explicit block selections do not imply that other pages were requested.
        if (idScope.Count > 0) emptyPages = Array.Empty<int>();
        int maxCharacters = Math.Min(profile.MaxRequestCharacters, request.Limits.MaxRequestCharacters);
        var batches = new List<Batch>();
        var omitted = new List<string>();
        var currentText = new List<OfficeAiEvidence>();
        var currentImages = new List<OfficeAiImage>();
        string Serialize() => JsonSerializer.Serialize(new {
            schema = "officeimo.ai.request.v1", sourceHash = document.SourceHash, pageProvenance = document.PageProvenance,
            operation = request.Operation.ToString(), instruction = request.Instruction,
            fields = request.Fields.Select(field => new { name = field.Name, type = field.Type.ToString(), dateFormat = field.DateFormat }),
            evidence = currentText.Select(item => new { id = item.Id, kind = item.Kind, text = item.Text, page = item.Page }),
            images = currentImages.Select(item => new { id = item.Id, page = item.Page, width = item.Width, height = item.Height })
        });
        OfficeAiExecutionRequest CreateRequest() => new(requestId + "-" + (batches.Count + 1), Instructions, Serialize(), Schema,
            Array.AsReadOnly(currentImages.ToArray()), request.Limits.MaxResponseCharacters);
        bool Fits() => currentImages.Sum(image => (long)image.ByteLength) <= Math.Min(profile.MaxImageBytes, request.Limits.MaxImageBytes)
            && currentImages.Sum(image => (long)image.Width * image.Height) <= request.Limits.MaxImagePixels
            && _executor.MeasureRequestCharacters(CreateRequest()) <= maxCharacters;
        void Flush() {
            if (currentText.Count + currentImages.Count == 0) return;
            string[] ids = currentText.Select(item => item.Id).Concat(currentImages.Select(item => item.Id)).ToArray();
            if (batches.Count >= request.Limits.MaxRequests) omitted.AddRange(ids);
            else batches.Add(new Batch(CreateRequest(),
                currentText.ToDictionary(item => item.Id, StringComparer.Ordinal), currentImages.ToDictionary(item => item.Id, StringComparer.Ordinal), Array.AsReadOnly(ids)));
            currentText.Clear(); currentImages.Clear();
        }
        if (!Fits()) throw new ArgumentException("Instructions and schema exceed the execution profile's request limit.");
        var textByPage = text.ToLookup(item => item.Page ?? 0);
        var imagesByPage = images.ToLookup(item => item.Page);
        foreach (int page in text.Select(item => item.Page ?? 0).Concat(images.Select(item => item.Page)).Distinct()) {
            foreach (OfficeAiEvidence item in textByPage[page]) {
                currentText.Add(item);
                if (Fits()) continue;
                currentText.RemoveAt(currentText.Count - 1); Flush(); currentText.Add(item);
                if (!Fits()) { omitted.Add(item.Id); currentText.Clear(); }
            }
            // Keep a page's images adjacent to its native observations wherever budgets allow.
            foreach (OfficeAiImage item in imagesByPage[page]) {
                currentImages.Add(item);
                if (Fits()) continue;
                currentImages.RemoveAt(currentImages.Count - 1); Flush(); currentImages.Add(item);
                if (!Fits()) { omitted.Add(item.Id); currentImages.Clear(); }
            }
        }
        Flush();
        return new Plan(batches.AsReadOnly(), omitted.AsReadOnly(), Array.AsReadOnly(emptyPages));
    }
}
