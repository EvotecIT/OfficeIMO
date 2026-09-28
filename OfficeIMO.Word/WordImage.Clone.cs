using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word;

public partial class WordImage {
    private WordImage CloneToParagraph(WordParagraph paragraph) {
        var sourceOwner = GetContainingPart();
        var destinationOwner = WordPartOwnership.Resolve(paragraph._document, paragraph._paragraph);
        const string relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var relationships = new Dictionary<string, string>(StringComparer.Ordinal);
        string Remap(string id) {
            if (ReferenceEquals(sourceOwner, destinationOwner)) return id;
            if (relationships.TryGetValue(id, out var mapped)) return mapped;
            if (sourceOwner.TryGetPartById(id, out var part) && part is ImagePart imagePart) {
                var destinationPart = destinationOwner.AddPart(imagePart);
                mapped = destinationOwner.GetIdOfPart(destinationPart);
            } else if (sourceOwner.HyperlinkRelationships.FirstOrDefault(item => item.Id == id) is HyperlinkRelationship hyperlink) {
                mapped = destinationOwner.AddHyperlinkRelationship(hyperlink.Uri, hyperlink.IsExternal).Id;
            } else {
                var external = sourceOwner.ExternalRelationships.FirstOrDefault(item => item.Id == id)
                    ?? throw new InvalidOperationException("The source drawing relationship cannot be resolved.");
                mapped = destinationOwner.AddExternalRelationship(external.RelationshipType, external.Uri).Id;
            }
            relationships.Add(id, mapped);
            return mapped;
        }
        if (_vmlShape != null) {
            var shape = (V.Shape)_vmlShape.CloneNode(true);
            shape.Id = "image-" + Guid.NewGuid().ToString("N");
            var picture = new Picture();
            if (shape.Type?.Value is string reference && reference.StartsWith("#", StringComparison.Ordinal)) {
                var definition = sourceOwner.RootElement?.Descendants<V.Shapetype>()
                    .FirstOrDefault(item => item.Id?.Value == reference.Substring(1));
                if (definition != null) {
                    var copiedDefinition = (V.Shapetype)definition.CloneNode(true);
                    copiedDefinition.Id = "image-type-" + Guid.NewGuid().ToString("N");
                    shape.Type = "#" + copiedDefinition.Id.Value;
                    picture.Append(copiedDefinition);
                }
            }
            foreach (OpenXmlElement element in picture.Descendants().Concat(new OpenXmlElement[] { shape }).Concat(shape.Descendants())) {
                foreach (OpenXmlAttribute attribute in element.GetAttributes().Where(attribute => attribute.NamespaceUri == relationshipNamespace).ToArray())
                    if (!string.IsNullOrEmpty(attribute.Value))
                        element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, Remap(attribute.Value!)));
            }
            picture.Append(shape);
            var run = new Run(picture);
            paragraph._paragraph.Append(run);
            return new WordImage(paragraph._document, paragraph._paragraph, run, shape);
        }
        var drawing = (DocumentFormat.OpenXml.Wordprocessing.Drawing)_Image.CloneNode(true);
        foreach (OpenXmlElement element in new OpenXmlElement[] { drawing }.Concat(drawing.Descendants()))
            foreach (OpenXmlAttribute attribute in element.GetAttributes().Where(attribute => attribute.NamespaceUri == relationshipNamespace).ToArray())
                if (!string.IsNullOrEmpty(attribute.Value))
                    element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, Remap(attribute.Value!)));
        WordDrawingIdAllocator.Reassign(paragraph._document, drawing);
        paragraph._paragraph.Append(new Run(drawing));
        return new WordImage(paragraph._document, drawing);
    }
}
