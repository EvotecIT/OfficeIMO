using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Internal;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents an embedded OLE object placed on a slide.
    /// </summary>
    public sealed class PowerPointOleObject : PowerPointShape {
        internal const string DefaultContentType =
            "application/vnd.openxmlformats-officedocument.oleObject";
        internal const int MaximumStorageBytes = 64 * 1024 * 1024;

        private readonly SlidePart _slidePart;

        internal PowerPointOleObject(P.GraphicFrame frame,
            SlidePart slidePart) : base(frame) {
            _slidePart = slidePart
                ?? throw new ArgumentNullException(nameof(slidePart));
        }

        private P.OleObject OleObject =>
            ((P.GraphicFrame)Element).Graphic?.GraphicData?
                .GetFirstChild<P.OleObject>()
            ?? throw new InvalidOperationException(
                "The graphic frame has no OLE object definition.");

        internal EmbeddedObjectPart EmbeddedPart {
            get {
                string relationshipId = OleObject.Id?.Value
                    ?? throw new InvalidOperationException(
                        "The OLE object has no embedded-part relationship.");
                return _slidePart.GetPartById(relationshipId) as
                    EmbeddedObjectPart
                    ?? throw new InvalidOperationException(
                        "The OLE relationship does not target an embedded-object part.");
            }
        }

        /// <summary>Gets or sets the OLE programmatic class identifier.</summary>
        public string? ProgId {
            get => OleObject.ProgId?.Value;
            set => OleObject.ProgId = string.IsNullOrWhiteSpace(value)
                ? null
                : new DocumentFormat.OpenXml.StringValue(value);
        }

        /// <summary>Gets or sets whether the object is displayed as an icon.</summary>
        public bool ShowAsIcon {
            get => OleObject.ShowAsIcon?.Value == true;
            set => OleObject.ShowAsIcon = value ? true : null;
        }

        /// <summary>Gets or sets how the object follows the presentation color scheme.</summary>
        public P.OleObjectFollowColorSchemeValues FollowColorScheme {
            get => OleObject.GetFirstChild<P.OleObjectEmbed>()?
                .FollowColorScheme?.Value
                ?? P.OleObjectFollowColorSchemeValues.None;
            set {
                P.OleObjectEmbed embed = OleObject
                    .GetFirstChild<P.OleObjectEmbed>()
                    ?? throw new InvalidOperationException(
                        "The OLE object is linked rather than embedded.");
                embed.FollowColorScheme = value;
            }
        }

        /// <summary>Gets the MIME content type of the embedded object part.</summary>
        public string ContentType => EmbeddedPart.ContentType;

        /// <summary>Copies the complete embedded OLE compound storage to a stream.</summary>
        public void CopyDataTo(Stream destination) {
            if (destination == null) {
                throw new ArgumentNullException(nameof(destination));
            }
            if (!destination.CanWrite) {
                throw new ArgumentException(
                    "Destination stream must be writable.", nameof(destination));
            }
            using Stream source = EmbeddedPart.GetStream(
                FileMode.Open, FileAccess.Read);
            source.CopyTo(destination);
        }

        /// <summary>Returns the complete embedded OLE compound storage.</summary>
        public byte[] GetData() {
            using var output = new MemoryStream();
            CopyDataTo(output);
            return output.ToArray();
        }

        /// <summary>
        ///     Replaces the embedded storage while retaining its relationship and content type.
        /// </summary>
        public void UpdateData(Stream storage) {
            if (storage == null) throw new ArgumentNullException(nameof(storage));
            if (!storage.CanRead) {
                throw new ArgumentException(
                    "OLE storage stream must be readable.", nameof(storage));
            }
            byte[] storageBytes = ReadStorage(storage);
            if (!TryValidateStorage(storageBytes, out string? reason)) {
                throw new InvalidDataException(reason
                    ?? "The embedded object is not an OLE compound storage.");
            }
            using var source = new MemoryStream(storageBytes,
                writable: false);
            EmbeddedObjectPart embeddedPart = EmbeddedPart;
            if (IsSharedOutsideThisFrame(embeddedPart)) {
                ReplaceSharedEmbeddedPart(embeddedPart, source);
                return;
            }
            embeddedPart.FeedData(source);
        }

        private bool IsSharedOutsideThisFrame(
            EmbeddedObjectPart embeddedPart) {
            int localConsumers = _slidePart.RootElement?
                .Descendants<P.OleObject>()
                .Count(candidate => {
                    string? relationshipId = candidate.Id?.Value;
                    return relationshipId != null
                        && relationshipId.Length != 0
                        && _slidePart.TryGetPartById(relationshipId,
                            out OpenXmlPart? target)
                        && ReferenceEquals(target, embeddedPart);
                }) ?? 0;
            if (localConsumers > 1) return true;
            return embeddedPart.GetParentParts().Any(parent =>
                !ReferenceEquals(parent, _slidePart));
        }

        private void ReplaceSharedEmbeddedPart(EmbeddedObjectPart original,
            Stream replacement) {
            EmbeddedObjectPart detached = _slidePart
                .AddEmbeddedObjectPart(original.ContentType);
            try {
                detached.FeedData(replacement);
                OleObject.Id = _slidePart.GetIdOfPart(detached);
            } catch {
                if (_slidePart.Parts.Any(pair =>
                        ReferenceEquals(pair.OpenXmlPart, detached))) {
                    _slidePart.DeletePart(detached);
                }
                throw;
            }
        }

        internal static bool TryValidateStorage(byte[] storageBytes,
            out string? reason) {
            var limits = new OfficeCompoundReadOptions(
                maxStreamBytes: MaximumStorageBytes,
                maxTotalStreamBytes: MaximumStorageBytes);
            return OfficeCompoundFileReader.TryRead(storageBytes, limits,
                out OfficeCompoundFile? compound, out reason)
                && compound != null;
        }

        private static byte[] ReadStorage(Stream storage) {
            if (storage.CanSeek) storage.Position = 0;
            byte[] bytes = OfficeStreamReader.ReadAllBytes(storage,
                MaximumStorageBytes);
            if (bytes.Length == 0) {
                throw new InvalidDataException(
                    "The embedded OLE storage is empty.");
            }
            return bytes;
        }
    }
}
