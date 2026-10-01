using DocumentFormat.OpenXml.Packaging;
using System.Threading;

namespace OfficeIMO.Word;

/// <summary>Stages media writes and MIME changes while keeping originals reachable until commit.</summary>
internal sealed class WordImagePartReplacement {
    internal WordImagePartReplacement(ImagePart part, byte[] original, byte[] candidate, string contentType) {
        Part = part; Original = original; Candidate = candidate; ContentType = contentType;
        Parents = part.GetParentParts().SelectMany(parent => parent.Parts
            .Where(pair => pair.OpenXmlPart.Uri == part.Uri)
            .Select(pair => (Parent: parent, Id: pair.RelationshipId))).ToArray();
        if (contentType != part.ContentType && Parents.GroupBy(item => item.Parent.Uri).Any(group => group.Count() > 1))
            throw new NotSupportedException("MIME replacement cannot preserve duplicate relationships to one image from the same parent.");
    }
    private ImagePart Part { get; }
    private byte[] Original { get; }
    private byte[] Candidate { get; }
    private string ContentType { get; }
    private (OpenXmlPart Parent, string Id)[] Parents { get; }
    private ImagePart? Replacement;

    internal static void Apply(MainDocumentPart main, IReadOnlyList<WordImagePartReplacement> staged, CancellationToken token) {
        if (staged.Count == 0) return;
        HeaderPart? holder = null;
        var writes = new List<WordImagePartReplacement>();
        var links = new List<(OpenXmlPart Parent, string Id, ImagePart Original)>();
        try {
            if (staged.Any(item => item.Part.ContentType != item.ContentType)) {
                // An unreferenced temporary header supplies SDK-supported image relationships.
                // No section references it, and it is removed before this operation returns.
                holder = main.AddNewPart<HeaderPart>();
                foreach (var item in staged.Where(item => item.Part.ContentType != item.ContentType)) {
                    token.ThrowIfCancellationRequested();
                    holder.AddPart(item.Part);
                    item.Replacement = holder.AddImagePart(item.ContentType);
                    using var candidate = new MemoryStream(item.Candidate, writable: false);
                    item.Replacement.FeedData(candidate);
                }
            }
            foreach (var item in staged) {
                token.ThrowIfCancellationRequested();
                if (item.Replacement == null) {
                    writes.Add(item);
                    using var candidate = new MemoryStream(item.Candidate, writable: false);
                    item.Part.FeedData(candidate);
                } else {
                    foreach (var parent in item.Parents) {
                        token.ThrowIfCancellationRequested();
                        links.Add((parent.Parent, parent.Id, item.Part));
                        parent.Parent.DeletePart(parent.Id);
                        parent.Parent.AddPart(item.Replacement, parent.Id);
                    }
                }
            }
            token.ThrowIfCancellationRequested();
        } catch {
            for (int i = links.Count - 1; i >= 0; i--) {
                var link = links[i];
                link.Parent.DeletePart(link.Id);
                link.Parent.AddPart(link.Original, link.Id);
            }
            for (int i = writes.Count - 1; i >= 0; i--) {
                using var original = new MemoryStream(writes[i].Original, writable: false);
                writes[i].Part.FeedData(original);
            }
            throw;
        } finally {
            if (holder != null) main.DeletePart(holder);
        }
    }
}
