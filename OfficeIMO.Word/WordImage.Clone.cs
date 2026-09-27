using Blip = DocumentFormat.OpenXml.Drawing.Blip;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word;

public partial class WordImage {
    private WordImage CloneToParagraph(WordParagraph paragraph) {
        var sourceOwner = GetContainingPart();
        var destinationOwner = WordPartOwnership.Resolve(paragraph._document, paragraph._paragraph);
        var relationships = new Dictionary<string, string>(StringComparer.Ordinal);
        string Remap(string id) {
            if (ReferenceEquals(sourceOwner, destinationOwner)) return id;
            if (relationships.TryGetValue(id, out var mapped)) return mapped;
            if (sourceOwner.TryGetPartById(id, out var part) && part is ImagePart imagePart) {
                var destinationPart = destinationOwner.AddPart(imagePart);
                mapped = destinationOwner.GetIdOfPart(destinationPart);
            } else {
                var external = sourceOwner.ExternalRelationships.FirstOrDefault(item => item.Id == id)
                    ?? throw new InvalidOperationException("The source image relationship cannot be resolved.");
                mapped = destinationOwner.AddExternalRelationship(external.RelationshipType, external.Uri).Id;
            }
            relationships.Add(id, mapped);
            return mapped;
        }
        if (_vmlShape != null) {
            var shape = (V.Shape)_vmlShape.CloneNode(true);
            shape.Id = "image-" + Guid.NewGuid().ToString("N");
            foreach (var image in shape.Descendants<V.ImageData>())
                if (image.RelationshipId?.Value is string id) image.RelationshipId = Remap(id);
            var run = new Run(new Picture(shape));
            paragraph._paragraph.Append(run);
            return new WordImage(paragraph._document, paragraph._paragraph, run, shape);
        }
        var drawing = (DocumentFormat.OpenXml.Wordprocessing.Drawing)_Image.CloneNode(true);
        foreach (var blip in drawing.Descendants<Blip>()) {
            if (blip.Embed?.Value is string id) blip.Embed = Remap(id);
            if (blip.Link?.Value is string link) blip.Link = Remap(link);
        }
        paragraph._paragraph.Append(new Run(drawing));
        return new WordImage(paragraph._document, drawing);
    }
}
