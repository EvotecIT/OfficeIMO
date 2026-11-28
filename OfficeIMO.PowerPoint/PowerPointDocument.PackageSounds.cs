using System.Globalization;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private Dictionary<string, int>? _packageMediaSequence;

        internal bool TryAddPackageTransitionSound(string slide, Stream audio, string contentType,
            string extension, out string relationshipId) {
            relationshipId = string.Empty;
            if (GetPackage(edit: true) is not { } package) return false;
            string suffix = PowerPointEmbeddedSound.ValidateSource(audio, contentType, extension);
            // The SDK accepts arbitrary suffixes. Standard WAV and AIFF suffixes use the owned part writer.
            if (suffix.ToLowerInvariant() is not ("wav" or "aif" or "aiff") || contentType != contentType.Trim()) return false;

            if (audio.CanSeek) {
                if (audio.Length > int.MaxValue) return false;
                audio.Position = 0;
            }
            byte[] data = OfficeStreamReader.ReadRemainingBytes(audio);
            OfficeXmlTree relationships = package.GetRelationships(slide);
            long nextRelationship = 1;
            string id = ReserveSlideRelationshipId(ReadPackageRelationshipIds(relationships), ref nextRelationship);
            string part = NextPackageMediaPart(package, contentType, suffix);

            package.CreateBinaryPart(part, contentType, data);
            try {
                package.AddRelationship(slide, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "audio", part, id: id);
            } catch {
                package.DeletePart(part);
                throw;
            }

            InvalidatePackageRelationships();
            MarkPackageEdited();
            relationshipId = id;
            return true;
        }

        private string NextPackageMediaPart(OfficeOpcEditablePackage package, string contentType, string suffix) {
            _packageMediaSequence ??= new Dictionary<string, int>(StringComparer.Ordinal);
            _packageMediaSequence.TryGetValue(contentType, out int last);
            string part;
            do {
                last++;
                // The SDK uses an absolute media part name and writes the supplied suffix without a dot.
                part = "/media/mediadata" + (last == 1 ? string.Empty : last.ToString(CultureInfo.InvariantCulture)) + suffix;
            } while (package.PartExists(part));
            _packageMediaSequence[contentType] = last;
            return part;
        }
    }
}
