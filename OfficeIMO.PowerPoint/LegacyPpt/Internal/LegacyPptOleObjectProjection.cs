using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    ///     Retains one native embedded-object mapping and its projected editable state.
    /// </summary>
    internal sealed class LegacyPptOleObjectProjection {
        private readonly byte[] _storageBytes;
        private readonly string? _progId;
        private readonly bool _showAsIcon;
        private readonly P.OleObjectFollowColorSchemeValues _colorFollow;

        private LegacyPptOleObjectProjection(
            LegacyPptEmbeddedOleObject source, PowerPointOleObject projected) {
            Source = source ?? throw new ArgumentNullException(nameof(source));
            if (projected == null) {
                throw new ArgumentNullException(nameof(projected));
            }
            EmbeddedPartUri = projected.EmbeddedPart.Uri.ToString();
            EmbeddedContentType = projected.EmbeddedPart.ContentType;
            _storageBytes = source.GetBytes();
            _progId = projected.ProgId;
            _showAsIcon = projected.ShowAsIcon;
            _colorFollow = projected.FollowColorScheme;
        }

        internal LegacyPptEmbeddedOleObject Source { get; }

        internal string EmbeddedPartUri { get; }

        internal string EmbeddedContentType { get; }

        internal static LegacyPptOleObjectProjection Create(
            LegacyPptEmbeddedOleObject source, PowerPointOleObject projected) =>
            new(source, projected);

        internal bool TryGetChange(PowerPointOleObject current,
            out LegacyPptOleObjectEdit? edit) {
            edit = null;
            if (current == null
                || !string.Equals(current.EmbeddedPart.Uri.ToString(),
                    EmbeddedPartUri, StringComparison.Ordinal)
                || !string.Equals(current.EmbeddedPart.ContentType,
                    EmbeddedContentType, StringComparison.Ordinal)
                || current.EmbeddedPart.Parts.Any()
                || current.EmbeddedPart.ExternalRelationships.Any()
                || current.EmbeddedPart.HyperlinkRelationships.Any()) {
                return false;
            }

            byte[] storageBytes;
            try {
                using Stream stream = current.EmbeddedPart.GetStream(
                    FileMode.Open, FileAccess.Read);
                storageBytes = OfficeStreamReader.ReadAllBytes(stream,
                    64 * 1024 * 1024);
            } catch (Exception exception) when (exception is IOException
                                                or InvalidDataException
                                                or UnauthorizedAccessException) {
                return false;
            }
            if (!OfficeCompoundFileReader.TryRead(storageBytes,
                    out OfficeCompoundFile? compound, out _)
                || compound == null) {
                return false;
            }

            bool storageChanged = !BytesEqual(_storageBytes, storageBytes);
            bool progIdChanged = !string.Equals(_progId, current.ProgId,
                StringComparison.Ordinal);
            bool drawAspectChanged = _showAsIcon != current.ShowAsIcon;
            bool colorFollowChanged = _colorFollow !=
                current.FollowColorScheme;
            if (storageChanged || progIdChanged || drawAspectChanged
                || colorFollowChanged) {
                edit = new LegacyPptOleObjectEdit(this, storageBytes,
                    storageChanged, progIdChanged, drawAspectChanged,
                    colorFollowChanged, current.ProgId,
                    current.ShowAsIcon, current.FollowColorScheme);
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
            P.OleObjectFollowColorSchemeValues colorFollow) {
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
        internal P.OleObjectFollowColorSchemeValues ColorFollow { get; }
        internal bool MetadataChanged => ProgIdChanged || DrawAspectChanged
            || ColorFollowChanged;
    }
}
