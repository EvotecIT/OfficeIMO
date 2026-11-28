using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>Creates the playback references and the poster of a media frame.</summary>
        internal (string File, string Playback, string Poster) AddPackageMediaParts(OpenXml.PowerPointPartRef slide,
            string fileReferenceType, Uri? source, byte[]? media, string? contentType, string? suffix,
            byte[] poster, string posterContentType, string posterExtension) {
            OfficeOpcEditablePackage package = GetPackage(edit: true) ?? throw new InvalidOperationException("The presentation can't be edited.");
            string slidePart = slide.Name;
            OfficeXmlTree relationships = package.GetRelationships(slidePart);
            long nextRelationship = 1;
            HashSet<string> used = ReadPackageRelationshipIds(relationships);
            string fileId = ReserveSlideRelationshipId(used, ref nextRelationship);
            string playbackId = source != null ? fileId : ReserveSlideRelationshipId(used, ref nextRelationship);
            string? mediaPart = source == null ? NextPackageMediaPart(package, contentType!, suffix!) : null;
            string target = mediaPart ?? PowerPointHyperlinkResolver.TargetText(source!);
            if (mediaPart != null) package.CreateBinaryPart(mediaPart, contentType, media!);
            try {
                package.AddRelationship(slidePart, fileReferenceType, target, external: source != null, id: fileId);
                if (mediaPart != null) package.AddRelationship(slidePart, PowerPointMedia.PlaybackMediaRelationshipType, target, id: playbackId);
                if (!TryAddPackageImagePart(slidePart, poster, posterContentType, posterExtension, out string posterId, out _, out _, reuse: false))
                    throw new InvalidOperationException("The presentation can't store the media poster.");
                InvalidatePackageRelationships();
                MarkPackageEdited();
                return (fileId, playbackId, posterId);
            } catch {
                int file = FindRelationship(relationships, fileId);
                if (file != OfficeXmlTree.NoNode) relationships.Remove(file);
                if (mediaPart != null) {
                    int playback = FindRelationship(relationships, playbackId);
                    if (playback != OfficeXmlTree.NoNode) relationships.Remove(playback);
                    package.DeletePart(mediaPart);
                }
                InvalidatePackageRelationships();
                throw;
            }
        }
    }
}
