using OfficeIMO.Core.Internal;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    ///     Retains one native embedded-object mapping and its projected editable state.
    /// </summary>
    internal sealed class LegacyPptOleObjectProjection {
        private const int MaximumStorageBytes = 64 * 1024 * 1024;
        private readonly byte[] _storageBytes;
        private readonly string? _progId;
        private readonly bool _showAsIcon;
        private readonly PowerPointOleFollowColorScheme _colorFollow;

        private LegacyPptOleObjectProjection(
            LegacyPptEmbeddedOleObject source, LegacyPptPartView slidePart, OfficeXmlNode frame) {
            Source = source ?? throw new ArgumentNullException(nameof(source));
            OfficeXmlNode ole = OleObject(frame)
                ?? throw new InvalidOperationException(
                    "The graphic frame has no OLE object definition.");
            LegacyPptPartView embeddedPart = EmbeddedPart(slidePart, ole);
            EmbeddedPartUri = embeddedPart.Name;
            EmbeddedContentType = embeddedPart.ContentType ?? string.Empty;
            _storageBytes = source.GetBytes();
            _progId = ole.Attribute("progId");
            _showAsIcon = ole.Bool("showAsIcon") == true;
            _colorFollow = FollowColorScheme(ole);
        }

        internal LegacyPptEmbeddedOleObject Source { get; }

        internal string EmbeddedPartUri { get; }

        internal string EmbeddedContentType { get; }

        /// <summary>Maps an embedded object of a binary shape to the projected graphic frame read from a view.</summary>
        internal static LegacyPptOleObjectProjection Create(
            LegacyPptEmbeddedOleObject source, LegacyPptPartView slidePart, OfficeXmlNode frame) =>
            new(source, slidePart, frame);

        /// <summary>The <c>p:oleObj</c> of a graphic frame, or null when it has none.</summary>
        internal static OfficeXmlNode? OleObject(OfficeXmlNode frame) {
            OfficeXmlNode ole = frame.Child(LegacyPptXml.A, "graphic")
                .Child(LegacyPptXml.A, "graphicData").Child(LegacyPptXml.P, "oleObj");
            return ole.IsNone ? null : ole;
        }

        private static LegacyPptPartView EmbeddedPart(LegacyPptPartView slidePart, OfficeXmlNode ole) {
            string relationshipId = ole.Attribute(LegacyPptXml.R, "id")
                ?? throw new InvalidOperationException(
                    "The OLE object has no embedded-part relationship.");
            if (!slidePart.TryGetRelationship(relationshipId, out LegacyPptRelationshipView relationship)
                || relationship.IsExternal || relationship.TargetPart.IsNone
                || !relationship.IsType(LegacyPptPackageView.OleObjectRelationshipType)) {
                throw new InvalidOperationException(
                    "The OLE relationship does not target an embedded-object part.");
            }
            return relationship.TargetPart;
        }

        private static PowerPointOleFollowColorScheme FollowColorScheme(OfficeXmlNode ole) =>
            ole.Child(LegacyPptXml.P, "embed").Attribute("followColorScheme") is { } token
            && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token,
                out PowerPointOleFollowColorScheme value)
                ? value
                : PowerPointOleFollowColorScheme.None;

        internal bool TryGetChange(LegacyPptPartView slidePart, OfficeXmlNode frame,
            out LegacyPptOleObjectEdit? edit) {
            edit = null;
            if (slidePart.IsNone || OleObject(frame) is not { } ole) return false;
            LegacyPptPartView embeddedPart;
            try {
                embeddedPart = EmbeddedPart(slidePart, ole);
            } catch (InvalidOperationException) {
                return false;
            }
            if (!string.Equals(embeddedPart.Name,
                    EmbeddedPartUri, StringComparison.Ordinal)
                || !string.Equals(embeddedPart.ContentType,
                    EmbeddedContentType, StringComparison.Ordinal)
                || embeddedPart.Relationships.Count > 0) {
                return false;
            }

            byte[] storageBytes;
            try {
                storageBytes = embeddedPart.GetBytes();
            } catch (Exception exception) when (exception is IOException
                                                or InvalidDataException
                                                or UnauthorizedAccessException) {
                return false;
            }
            if (storageBytes.Length > MaximumStorageBytes
                || !OfficeCompoundFileReader.TryRead(storageBytes,
                    out OfficeCompoundFile? compound, out _)
                || compound == null) {
                return false;
            }

            string? progId = ole.Attribute("progId");
            bool showAsIcon = ole.Bool("showAsIcon") == true;
            PowerPointOleFollowColorScheme colorFollow = FollowColorScheme(ole);
            bool storageChanged = !BytesEqual(_storageBytes, storageBytes);
            bool progIdChanged = !string.Equals(_progId, progId,
                StringComparison.Ordinal);
            bool drawAspectChanged = _showAsIcon != showAsIcon;
            bool colorFollowChanged = _colorFollow != colorFollow;
            if (storageChanged || progIdChanged || drawAspectChanged
                || colorFollowChanged) {
                edit = new LegacyPptOleObjectEdit(this, storageBytes,
                    storageChanged, progIdChanged, drawAspectChanged,
                    colorFollowChanged, progId, showAsIcon, colorFollow);
            }
            return true;
        }

        private static bool BytesEqual(byte[] left, byte[] right) {
            if (ReferenceEquals(left, right)) return true;
            if (left.Length != right.Length) return false;
            for (int index = 0; index < left.Length; index++) {
                if (left[index] != right[index]) return false;
            }
            return true;
        }
    }

    internal sealed class LegacyPptOleObjectEdit {
        internal LegacyPptOleObjectEdit(
            LegacyPptOleObjectProjection projection, byte[] storageBytes,
            bool storageChanged, bool progIdChanged,
            bool drawAspectChanged, bool colorFollowChanged,
            string? progId, bool showAsIcon,
            PowerPointOleFollowColorScheme colorFollow) {
            Projection = projection;
            StorageBytes = storageBytes;
            StorageChanged = storageChanged;
            ProgIdChanged = progIdChanged;
            DrawAspectChanged = drawAspectChanged;
            ColorFollowChanged = colorFollowChanged;
            ProgId = progId;
            ShowAsIcon = showAsIcon;
            ColorFollow = colorFollow;
        }

        internal LegacyPptOleObjectProjection Projection { get; }
        internal byte[] StorageBytes { get; }
        internal bool StorageChanged { get; }
        internal bool ProgIdChanged { get; }
        internal bool DrawAspectChanged { get; }
        internal bool ColorFollowChanged { get; }
        internal string? ProgId { get; }
        internal bool ShowAsIcon { get; }
        internal PowerPointOleFollowColorScheme ColorFollow { get; }
        internal bool MetadataChanged => ProgIdChanged || DrawAspectChanged
            || ColorFollowChanged;
    }
}
