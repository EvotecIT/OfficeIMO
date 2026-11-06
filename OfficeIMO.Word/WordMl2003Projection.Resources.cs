using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
    private static readonly XNamespace Vml = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private readonly Dictionary<string, XElement> _binarySources = new(StringComparer.Ordinal);
    private readonly Dictionary<string, (byte[] Bytes, string ContentType)> _decodedImages = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _imageRelationships = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _hyperlinkRelationships = new(StringComparer.Ordinal);
    private int _relationshipCount;
    private int _packageEntryCount = 1; // [Content_Types].xml
    private readonly HashSet<OpenXmlPartContainer> _relationshipOwners = new();
    private bool _corePropertiesReserved;
    private static readonly HashSet<string> VmlElements = Names("shape shapetype imagedata stroke formulas f path fill");
    private static readonly Dictionary<string, string> VmlAttributes = new(StringComparer.Ordinal) {
        ["shape"] = "id type style alt title coordsize coordorigin wrapcoords filled fillcolor stroked strokecolor strokeweight insetpen spid spt preferrelative allowincell allowoverlap behindDoc",
        ["shapetype"] = "id coordsize coordorigin path filled stroked spt preferrelative",
        ["imagedata"] = "src title cropleft cropright croptop cropbottom blacklevel gain gamma grayscale bilevel chromakey embosscolor recolortarget detectmouseclick",
        ["stroke"] = "joinstyle endcap dashstyle color weight opacity on",
        ["path"] = "v extrusionok gradientshapeok connecttype connectlocs connectangles textboxrect fillok strokeok shadowok arrowok",
        ["f"] = "eqn", ["formulas"] = "", ["fill"] = "on color color2 opacity opacity2 type angle focus focusposition focussize method"
    };
    private long _remainingReconstructionBytes;
    private OfficePackageSecurityOptions _resourcePolicy = null!;

    private void IndexImages(XElement root) {
        foreach (XElement data in root.Descendants(Word + "binData")) {
            CheckAttributes(data, "name");
            string? name = (string?)data.Attribute(Word + "name");
            if (name == null || name.Length == 0 || data.HasElements || _binarySources.ContainsKey(name))
                throw new InvalidDataException("WordML 2003 contains a missing or duplicate binary resource name.");
            _binarySources.Add(name, data);
        }
    }

    private void AppendHyperlink(OpenXmlCompositeElement target, XElement source) {
        CheckAttributes(source, "dest bookmark screenTip target");
        string? destination = (string?)source.Attribute(Word + "dest");
        string? anchor = (string?)source.Attribute(Word + "bookmark");
        var hyperlink = new Hyperlink {
            Anchor = anchor, Tooltip = (string?)source.Attribute(Word + "screenTip"),
            TargetFrame = (string?)source.Attribute(Word + "target")
        };
        if (!string.IsNullOrEmpty(destination)) {
            if (!Uri.TryCreate(destination, UriKind.RelativeOrAbsolute, out Uri? uri)) {
                Omit(source, "invalid hyperlink destination");
            } else {
                string key = _story.Uri.OriginalString + "|" + destination;
                if (!_hyperlinkRelationships.TryGetValue(key, out string? id)) {
                    ReserveRelationship(_story);
                    id = _story.AddHyperlinkRelationship(uri, true).Id;
                    _hyperlinkRelationships.Add(key, id);
                }
                hyperlink.Id = id;
            }
        }
        foreach (XElement child in source.Elements()) AppendInline(hyperlink, child);
        target.AppendChild(hyperlink);
    }

    private Picture? ProjectPicture(XElement source) {
        CheckAttributes(source, "");
        var picture = new XElement(ModernWord + "pict");
        picture.SetAttributeValue(XNamespace.Xmlns + "w", ModernWord.NamespaceName);
        picture.SetAttributeValue(XNamespace.Xmlns + "r", Relationships.NamespaceName);
        foreach (XElement child in source.Elements()) {
            if (child.Name == Word + "binData") continue;
            if (child.Name != Vml + "shape" && child.Name != Vml + "shapetype") {
                Omit(child, "picture content");
                continue;
            }
            XElement? mapped = ProjectVml(child);
            if (mapped != null) picture.Add(mapped);
        }
        if (!picture.Elements(Vml + "shape").Any()) return null;
        return new Picture(picture.ToString(SaveOptions.DisableFormatting));
    }

    private XElement? ProjectVml(XElement source) {
        _token.ThrowIfCancellationRequested();
        if (source.Name.Namespace != Vml || !VmlElements.Contains(source.Name.LocalName)) {
            if (source.Name == Office + "lock") {
                const string allowed = "ext aspectratio position selection grouping ungrouping rotation cropping verticies adjusthandles text shapetype";
                CheckAttributes(source, allowed, XNamespace.None);
                var locked = new XElement(source.Name);
                foreach (XAttribute attribute in source.Attributes()) {
                    if (IsMetadataAttribute(attribute) || attribute.Name.Namespace == XNamespace.None && ContainsName(allowed, attribute.Name.LocalName))
                        locked.Add(new XAttribute(attribute));
                }
                foreach (XElement child in source.Elements()) Omit(child, "nested VML lock content");
                if (source.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
                    Omit(source, "VML lock text content");
                return locked;
            }
            Omit(source, "VML picture structure");
            return null;
        }
        var result = new XElement(source.Name);
        foreach (XAttribute attribute in source.Attributes()) {
            if (attribute.IsNamespaceDeclaration) { result.Add(new XAttribute(attribute)); continue; }
            if (source.Name == Vml + "imagedata" && attribute.Name.LocalName == "src" && attribute.Name.Namespace == XNamespace.None) continue;
            if (!(attribute.Name.Namespace == XNamespace.None || attribute.Name.Namespace == Office) ||
                !ContainsName(VmlAttributes[source.Name.LocalName], attribute.Name.LocalName)) {
                Omit(attribute.Name, "VML picture attribute " + attribute.Name.LocalName);
                continue;
            }
            result.Add(new XAttribute(attribute));
        }
        if (source.Name == Vml + "imagedata") {
            string? name = (string?)source.Attribute("src");
            if (name == null || !_binarySources.TryGetValue(name, out XElement? data)) {
                Omit(source, "picture without an embedded binary resource");
                return null;
            }
            string key = _story.Uri.OriginalString + "|" + name;
            if (!_imageRelationships.TryGetValue(key, out string? id)) {
                if (!_decodedImages.TryGetValue(name, out var image)) {
                    string encoded = data.Value;
                    long symbols = 0;
                    foreach (char character in encoded)
                        if (character is not (' ' or '\t' or '\r' or '\n' or '=')) symbols++;
                    long decodedLength = symbols * 6 / 8;
                    if (decodedLength > _resourcePolicy.MaxPartUncompressedBytes || decodedLength > _remainingReconstructionBytes)
                        throw new InvalidDataException("WordML 2003 exceeds the decoded image resource limit.");
                    byte[] bytes;
                    try { bytes = Convert.FromBase64String(encoded); }
                    catch (FormatException error) { throw new InvalidDataException("WordML 2003 contains an invalid image payload.", error); }
                    if (bytes.LongLength > _resourcePolicy.MaxPartUncompressedBytes || bytes.LongLength > _remainingReconstructionBytes)
                        throw new InvalidDataException("WordML 2003 exceeds the decoded image resource limit.");
                    OfficeImageInfo info = OfficeImageReader.Identify(bytes);
                    image = (bytes, info.MimeType);
                    _decodedImages.Add(name, image);
                }
                if (image.Bytes.LongLength > _remainingReconstructionBytes)
                    throw new InvalidDataException("WordML 2003 exceeds the written image resource limit.");
                ReserveRelationship(_story, createsPart: true);
                _remainingReconstructionBytes -= image.Bytes.LongLength;
                ImagePart part = _story switch {
                    MainDocumentPart main => main.AddImagePart(image.ContentType),
                    HeaderPart header => header.AddImagePart(image.ContentType),
                    FooterPart footer => footer.AddImagePart(image.ContentType),
                    _ => throw new InvalidDataException("WordML image story has no supported package owner.")
                };
                using var input = new MemoryStream(image.Bytes, writable: false);
                part.FeedData(input);
                id = _story.GetIdOfPart(part);
                _imageRelationships.Add(key, id);
            }
            result.SetAttributeValue(Relationships + "id", id);
        }
        foreach (XElement child in source.Elements()) {
            if (!IsSupportedVmlChild(source.Name, child.Name)) {
                Omit(child, "VML picture structure inside " + source.Name.LocalName);
                continue;
            }
            XElement? mapped = ProjectVml(child);
            if (mapped != null) result.Add(mapped);
        }
        if (source.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
            Omit(source, "VML picture text content");
        return result;
    }

    private static bool IsSupportedVmlChild(XName parent, XName child) {
        if (parent == Vml + "shape" || parent == Vml + "shapetype") {
            return child == Office + "lock" || child.Namespace == Vml &&
                ContainsName("path formulas fill stroke imagedata", child.LocalName);
        }
        return parent == Vml + "formulas" && child == Vml + "f";
    }

    private void ReserveRelationship(OpenXmlPartContainer owner, bool createsPart = false) {
        bool firstRelationship = !_relationshipOwners.Contains(owner);
        int entries = (createsPart ? 1 : 0) + (firstRelationship ? 1 : 0);
        if (_relationshipCount >= _resourcePolicy.MaxRelationships ||
            _packageEntryCount > _resourcePolicy.MaxPartCount - entries)
            throw new InvalidDataException("WordML 2003 exceeds the package relationship or part limit.");
        _relationshipCount++;
        _packageEntryCount += entries;
        if (firstRelationship) _relationshipOwners.Add(owner);
    }
}
