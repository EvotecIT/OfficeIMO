using OfficeIMO.Html;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static HtmlResourceKind DirectResourceKind(XElement element, string attribute) {
        string name = element.Name.LocalName;
        if (name == "a" || name == "area" || element.Name == Ncx + "content") return HtmlResourceKind.Hyperlink;
        if (element.Name == Html + "link") return HtmlResourcePipeline.GetLinkResourceKind((string?)element.Attribute("rel"), (string?)element.Attribute("as"));
        if (attribute == "poster" || name == "img" || name == "image" || name == "feImage" || name == "use" ||
            (name == "source" && element.Parent?.Name == Html + "picture")) return HtmlResourceKind.Image;
        if (name == "audio" || name == "video" || name == "source" || name == "track") return HtmlResourceKind.Media;
        return HtmlResourceKind.Other;
    }

    private bool ValidateAuthoredResources(XDocument content, string owner, HtmlResourceManifest resources, ManifestLookup manifest, CancellationToken token,
        Func<string, bool>? validateStylesheet = null) {
        HtmlUrlPolicy policy = HtmlResourceUrlPolicy.Create(null);
        bool hasRemoteResources = false;
        string? baseHref = content.Root?.Name == Html + "html" ? content.Root.Element(Html + "head")?.Elements(Html + "base")
            .Select(element => (string?)element.Attribute("href")).FirstOrDefault(value => value != null) : null;
        // Rendering selects alternatives; publication validation must also inspect inactive direct carriers.
        foreach (var resource in DirectContentResources(content, owner, token)) Validate(resource.Reference, resource.Kind);
        // The shared owner additionally discovers inline CSS and its font/image dependencies.
        foreach (HtmlResourceReference resource in resources.Resources) {
            token.ThrowIfCancellationRequested();
            Validate(EpubReference.Resolve(owner, baseHref, resource.ResolvedSource.Length == 0 ? resource.Source : resource.ResolvedSource), resource.Kind);
        }
        // Inspect retained inline CSS independently of the renderer's current media selection.
        foreach (string css in content.Descendants().Where(element => element.Name.LocalName == "style" &&
            (element.Name.Namespace == Html || element.Name.NamespaceName == "http://www.w3.org/2000/svg")).Select(element => element.Value)
            .Concat(content.Descendants().Attributes("style").Select(attribute => attribute.Value))) {
            token.ThrowIfCancellationRequested();
            HtmlExternalStylesheetAnalysis analysis = HtmlResourcePipeline.AnalyzeExternalStylesheet(css,
                new Uri("epub://package/" + EncodePath(owner)), new HtmlResourcePipelineOptions(), includeInactiveResources: true);
            foreach (HtmlResourceReference resource in analysis.Imports.Select(import => import.Reference).Concat(analysis.FontResources).Concat(analysis.ImageResources))
                Validate(EpubReference.Resolve(owner, baseHref, resource.Source), resource.Kind);
        }
        return hasRemoteResources;

        void Validate(EpubReference reference, HtmlResourceKind kind) {
            if (reference.Kind == EpubReferenceKind.Data) {
                if (kind == HtmlResourceKind.Hyperlink || !HtmlDataUri.TryParse(reference.ResolvedValue!, out HtmlDataUri data) ||
                    !IsInertDataMediaType(kind, data.MediaType))
                    throw new NotSupportedException("Authored data URLs are limited to inert image, audio/video and font resources; embedded documents and SVG data are unsupported.");
                return;
            }
            if (kind == HtmlResourceKind.Hyperlink) return;
            if (reference.Kind == EpubReferenceKind.Container) {
                if (reference.ContainerPath == null || !manifest.IsLocal(reference.ContainerPath))
                    throw new InvalidDataException("Authored publication resource is not declared in the manifest: " + reference.Original);
                if (manifest.IsScripted(reference.ContainerPath))
                    throw new NotSupportedException("Authored content cannot embed a retained scripted resource.");
                if (kind == HtmlResourceKind.Stylesheet) {
                    var declaration = manifest.AtPath(reference.ContainerPath)[0];
                    if (!HasMediaType(declaration.MediaType, "text/css"))
                        throw new InvalidDataException("A stylesheet reference must select a declared CSS resource.");
                    if (validateStylesheet != null) hasRemoteResources |= validateStylesheet(reference.ContainerPath);
                }
            } else if (reference.Kind == EpubReferenceKind.External) {
                if (!HtmlUrlPolicyEvaluator.IsAllowed(reference.ResolvedValue, policy))
                    throw new NotSupportedException("Authored content contains a URL blocked by the shared HTML policy.");
                if (PackageVersion != "3.0" || (kind != HtmlResourceKind.Media && kind != HtmlResourceKind.Font))
                    throw new NotSupportedException("Remote publication resources are limited to EPUB 3 audio, video and fonts.");
                if (manifest.Remote(reference.Target!) is not EpubManifestItem declaration)
                    throw new InvalidDataException("Remote publication resource is not declared in the manifest: " + reference.Original);
                if (!IsInertDataMediaType(kind, declaration.MediaType))
                    throw new NotSupportedException("Remote resource declaration does not match an allowed audio, video or font media type.");
                hasRemoteResources = true;
            } else if (reference.Kind == EpubReferenceKind.Invalid) throw new InvalidDataException("Invalid authored resource URL: " + reference.Original);
        }
    }

    private static bool IsInertDataMediaType(HtmlResourceKind kind, string mediaType) {
        string type = mediaType.Split(';')[0].Trim();
        if (kind == HtmlResourceKind.Image) return new[] { "image/png", "image/jpeg", "image/gif", "image/webp", "image/avif" }.Contains(type, StringComparer.OrdinalIgnoreCase);
        if (kind == HtmlResourceKind.Media) return type.StartsWith("audio/", StringComparison.OrdinalIgnoreCase) || type.StartsWith("video/", StringComparison.OrdinalIgnoreCase);
        if (kind == HtmlResourceKind.Font) return type.StartsWith("font/", StringComparison.OrdinalIgnoreCase) ||
            new[] { "application/vnd.ms-opentype", "application/font-sfnt", "application/font-woff" }.Contains(type, StringComparer.OrdinalIgnoreCase);
        return false;
    }
}
