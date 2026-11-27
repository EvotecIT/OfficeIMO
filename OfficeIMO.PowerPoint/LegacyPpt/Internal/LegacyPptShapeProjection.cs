using System.Collections.ObjectModel;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Maps one projected Open XML shape to its OfficeArt shape container.</summary>
    internal sealed class LegacyPptShapeProjection {
        internal LegacyPptShapeProjection(uint openXmlShapeId, uint officeArtShapeId, long recordOffset,
            LegacyPptShapeKind kind, LegacyPptBounds bounds, string text,
            LegacyPptPlaceholder? placeholder,
            string? textFormattingFingerprint,
            IReadOnlyList<LegacyPptInteraction> shapeInteractions,
            IReadOnlyList<LegacyPptTextInteraction> textInteractions,
            LegacyPptAnimation? animation,
            ISet<uint> projectableSoundIds,
            bool canEditTextFormatting,
            string? textFrameFingerprint,
            bool canEditTextFrame,
            string? shapeTransformFingerprint,
            string? shapeGeometryFingerprint,
            string? groupCoordinateFingerprint,
            string? shapeVisualStyleFingerprint,
            string? pictureFormattingFingerprint,
            string? shapeVisibilityFingerprint,
            string shapeMetadataFingerprint,
            bool canEditShapeMetadata,
            LegacyPptOleObjectProjection? oleObject = null) {
            OpenXmlShapeId = openXmlShapeId;
            OfficeArtShapeId = officeArtShapeId;
            RecordOffset = recordOffset;
            Kind = kind;
            Bounds = bounds;
            Text = text ?? string.Empty;
            Placeholder = placeholder;
            TextFormattingFingerprint = textFormattingFingerprint;
            ShapeInteractions = new ReadOnlyCollection<LegacyPptInteraction>(
                shapeInteractions.ToArray());
            TextInteractions = new ReadOnlyCollection<LegacyPptTextInteraction>(
                textInteractions.ToArray());
            Animation = animation;
            CanEditInteractions = ShapeInteractions.All(interaction =>
                    IsEditableInteraction(interaction, projectableSoundIds))
                && TextInteractions.All(item => IsEditableInteraction(
                    item.Interaction, projectableSoundIds))
                && ShapeInteractions.GroupBy(item => item.Trigger)
                    .All(group => group.Count() == 1)
                && !HasOverlappingTextTriggers(TextInteractions);
            CanEditAnimation = animation == null || IsEditableAnimation(
                animation, projectableSoundIds);
            CanEditTextFormatting = canEditTextFormatting;
            TextFrameFingerprint = textFrameFingerprint;
            CanEditTextFrame = canEditTextFrame;
            ShapeTransformFingerprint = shapeTransformFingerprint;
            ShapeGeometryFingerprint = shapeGeometryFingerprint;
            GroupCoordinateFingerprint = groupCoordinateFingerprint;
            ShapeVisualStyleFingerprint = shapeVisualStyleFingerprint;
            PictureFormattingFingerprint = pictureFormattingFingerprint;
            ShapeVisibilityFingerprint = shapeVisibilityFingerprint;
            ShapeMetadataFingerprint = shapeMetadataFingerprint;
            CanEditShapeMetadata = canEditShapeMetadata;
            OleObject = oleObject;
        }

        internal uint OpenXmlShapeId { get; }

        internal uint OfficeArtShapeId { get; }

        internal long RecordOffset { get; }

        internal LegacyPptShapeKind Kind { get; }

        internal LegacyPptBounds Bounds { get; }

        internal string Text { get; }

        internal LegacyPptPlaceholder? Placeholder { get; }

        internal string? TextFormattingFingerprint { get; }

        internal IReadOnlyList<LegacyPptInteraction> ShapeInteractions { get; }

        internal IReadOnlyList<LegacyPptTextInteraction> TextInteractions { get; }

        internal LegacyPptAnimation? Animation { get; }

        internal bool CanEditInteractions { get; }

        internal bool CanEditAnimation { get; }

        internal bool CanEditTextFormatting { get; }

        internal string? TextFrameFingerprint { get; }

        internal bool CanEditTextFrame { get; }

        internal string? ShapeTransformFingerprint { get; }

        internal bool CanEditShapeTransform =>
            ShapeTransformFingerprint != null;

        internal string? ShapeGeometryFingerprint { get; }

        internal bool CanEditShapeGeometry =>
            ShapeGeometryFingerprint != null;

        internal string? GroupCoordinateFingerprint { get; }

        internal bool CanEditGroupCoordinate =>
            GroupCoordinateFingerprint != null;

        internal string? ShapeVisualStyleFingerprint { get; }

        internal bool CanEditShapeVisualStyle =>
            ShapeVisualStyleFingerprint != null;

        internal string? PictureFormattingFingerprint { get; }

        internal bool CanEditPictureFormatting =>
            PictureFormattingFingerprint != null;

        internal string? ShapeVisibilityFingerprint { get; }

        internal bool CanEditShapeVisibility =>
            ShapeVisibilityFingerprint != null;

        internal string ShapeMetadataFingerprint { get; }

        internal bool CanEditShapeMetadata { get; }

        internal LegacyPptOleObjectProjection? OleObject { get; }

        internal bool ShapeTransformMatches(LegacyPptProjectedShape shape) =>
            string.Equals(ShapeTransformFingerprint,
                CreateShapeTransformFingerprint(shape),
                StringComparison.Ordinal);

        internal static string? CreateShapeTransformFingerprint(
            LegacyPptProjectedShape projected) {
            PowerPointShape shape = projected.Shape;
            if (!LegacyPptWriter.TryReadShapeTransform(shape,
                    out _, out _)) {
                return null;
            }
            return string.Join("\n",
                shape.Rotation?.ToString("R",
                    System.Globalization.CultureInfo.InvariantCulture)
                    ?? string.Empty,
                shape.HorizontalFlip == true ? "1" : "0",
                shape.VerticalFlip == true ? "1" : "0");
        }

        internal bool TextFrameMatches(LegacyPptProjectedShape shape) =>
            string.Equals(TextFrameFingerprint,
                CreateTextFrameFingerprint(shape),
                StringComparison.Ordinal);

        internal static string? CreateTextFrameFingerprint(
            LegacyPptProjectedShape projected) {
            if (projected.Shape is not PowerPointTextBox textBox
                || !LegacyPptWriter.TryReadTextFrameForWrite(textBox,
                    out _, out _)
                || !projected.Element.Is(LegacyPptXml.P, "sp")) {
                return null;
            }
            return LegacyPptMasterProjection.Markup(projected.Element
                .Child(LegacyPptXml.P, "txBody").Child(LegacyPptXml.A, "bodyPr"));
        }

        internal bool ShapeGeometryMatches(LegacyPptProjectedShape shape) =>
            string.Equals(ShapeGeometryFingerprint,
                CreateShapeGeometryFingerprint(shape),
                StringComparison.Ordinal);

        internal static string? CreateShapeGeometryFingerprint(
            LegacyPptProjectedShape projected) {
            PowerPointShape shape = projected.Shape;
            if (!LegacyPptWriter.TryReadOfficeArtShapeType(shape,
                    requireConnector: false, out ushort shapeType, out _)
                || shapeType is not 2 and not 23
                || !LegacyPptWriter.TryReadShapeGeometry(shape, shapeType,
                    out _, out _)) {
                return null;
            }
            OfficeXmlNode geometry = projected.Element
                .SchemaDescendants(LegacyPptXml.A, "prstGeom").FirstOrDefault();
            OfficeXmlNode values = geometry.Child(LegacyPptXml.A, "avLst");
            return string.Concat(values.Children(LegacyPptXml.A, "gd").ToList()
                .Select(LegacyPptMasterProjection.Markup));
        }

        internal bool GroupCoordinateMatches(LegacyPptProjectedShape shape) =>
            string.Equals(GroupCoordinateFingerprint,
                CreateGroupCoordinateFingerprint(shape),
                StringComparison.Ordinal);

        internal static string? CreateGroupCoordinateFingerprint(
            LegacyPptProjectedShape projected) {
            if (projected.Shape is not PowerPointGroupShape group
                || !LegacyPptWriter.TryReadGroupForWrite(group,
                    out _, out _)) {
                return null;
            }
            OfficeXmlNode transform = projected.Element
                .Child(LegacyPptXml.P, "grpSpPr").Child(LegacyPptXml.A, "xfrm");
            OfficeXmlNode offset = transform.Child(LegacyPptXml.A, "chOff");
            OfficeXmlNode extents = transform.Child(LegacyPptXml.A, "chExt");
            return string.Join("\n",
                FormatLong(offset.Long("x")), FormatLong(offset.Long("y")),
                FormatLong(extents.Long("cx")), FormatLong(extents.Long("cy")));
        }

        private static string FormatLong(long? value) =>
            value?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;

        internal bool ShapeVisualStyleMatches(LegacyPptProjectedShape shape) =>
            string.Equals(ShapeVisualStyleFingerprint,
                CreateShapeVisualStyleFingerprint(shape),
                StringComparison.Ordinal);

        internal static string? CreateShapeVisualStyleFingerprint(
            LegacyPptProjectedShape projected) {
            if (!LegacyPptWriter.TryReadShapeVisualStyle(projected.Shape,
                    out _, out _)) {
                return null;
            }
            OfficeXmlNode element = projected.Element;
            OfficeXmlNode properties = element.Is(LegacyPptXml.P, "sp")
                || element.Is(LegacyPptXml.P, "cxnSp")
                || element.Is(LegacyPptXml.P, "pic")
                    ? element.Child(LegacyPptXml.P, "spPr")
                    : default;
            return string.Concat(properties.Children()
                .ToList()
                .Where(child => child.Is(LegacyPptXml.A, "noFill")
                    || child.Is(LegacyPptXml.A, "solidFill")
                    || child.Is(LegacyPptXml.A, "ln")
                    || child.Is(LegacyPptXml.A, "effectLst"))
                .Select(LegacyPptMasterProjection.Markup));
        }

        internal bool PictureFormattingMatches(LegacyPptProjectedShape picture) =>
            string.Equals(PictureFormattingFingerprint,
                CreatePictureFormattingFingerprint(picture),
                StringComparison.Ordinal);

        internal static string? CreatePictureFormattingFingerprint(
            LegacyPptProjectedShape projected) {
            if (projected.Shape is not PowerPointImage
                || projected.Shape is PowerPointMedia
                || !projected.Element.Is(LegacyPptXml.P, "pic")) {
                return null;
            }
            OfficeXmlNode picture = projected.Element;
            string protection = LegacyPptMasterProjection.Markup(picture
                .Path(LegacyPptXml.P, "nvPicPr", "cNvPicPr"));
            OfficeXmlNode fill = picture.Child(LegacyPptXml.P, "blipFill");
            string crop = LegacyPptMasterProjection.Markup(fill.Child(LegacyPptXml.A, "srcRect"));
            string effects = string.Concat(fill.Child(LegacyPptXml.A, "blip")
                .Children().ToList().Select(LegacyPptMasterProjection.Markup));
            return protection + "\n" + crop + "\n" + effects;
        }

        internal bool ShapeVisibilityMatches(PowerPointShape shape) =>
            string.Equals(ShapeVisibilityFingerprint,
                CreateShapeVisibilityFingerprint(shape),
                StringComparison.Ordinal);

        internal static string CreateShapeVisibilityFingerprint(
            PowerPointShape shape) => shape.Hidden ? "1" : "0";

        internal bool ShapeMetadataMatches(PowerPointShape shape) =>
            string.Equals(ShapeMetadataFingerprint,
                CreateShapeMetadataFingerprint(shape),
                StringComparison.Ordinal);

        internal static string CreateShapeMetadataFingerprint(
            PowerPointShape shape) {
            string name = shape.Name ?? string.Empty;
            string description = shape.AltText ?? string.Empty;
            return string.Concat(name.Length.ToString(
                    System.Globalization.CultureInfo.InvariantCulture),
                ":", name, "\n", description.Length.ToString(
                    System.Globalization.CultureInfo.InvariantCulture),
                ":", description);
        }

        internal bool TextFormattingMatches(LegacyPptProjectedShape shape) =>
            string.Equals(TextFormattingFingerprint,
                CreateTextFormattingFingerprint(shape),
                StringComparison.Ordinal);

        /// <summary>
        /// The fingerprint of a text shape's formatting: its text body with the text, body properties and run
        /// hyperlinks taken out, plus the bytes of the images its picture bullets use.
        /// </summary>
        internal static string CreateTextFormattingFingerprint(LegacyPptProjectedShape projected) {
            OfficeXmlNode textBody = projected.Element.Is(LegacyPptXml.P, "sp")
                ? projected.Element.Child(LegacyPptXml.P, "txBody")
                : default;
            if (textBody.IsNone) return string.Empty;
            OfficeXmlNode clone = textBody.Clone();
            LegacyPptProjectionFingerprint.Clear(clone.Child(LegacyPptXml.A, "bodyPr"));
            foreach (OfficeXmlNode text in clone.SchemaDescendants(LegacyPptXml.A, "t")) text.SetInnerText(string.Empty);
            foreach (OfficeXmlNode properties in clone.SchemaDescendants(LegacyPptXml.A, "rPr")) {
                properties.RemoveChildren(LegacyPptXml.A, "hlinkClick");
                properties.RemoveChildren(LegacyPptXml.A, "hlinkMouseOver");
                if (!HasAttributes(properties) && properties.FirstChild.IsNone) properties.Remove();
            }
            return LegacyPptXml.Digest(clone) + CreatePictureBulletImageFingerprint(
                textBody, projected.Part);
        }

        private static bool HasAttributes(OfficeXmlNode element) {
            if (element.Tree is not { } tree) return false;
            int count = tree.AttributeCountOf(element.Node);
            for (int index = 0; index < count; index++) {
                tree.GetAttributeAt(element.Node, index, out int namespaceId, out string localName, out _);
                if (namespaceId == OfficeXmlNamespaces.Xmlns
                    || namespaceId == OfficeXmlNamespaces.None && localName == "xmlns") continue;
                return true;
            }
            return false;
        }

        /// <summary>
        /// The relationship ids, content types and image bytes of the picture bullets below <paramref name="root"/>,
        /// so an image replaced in place is not mistaken for unchanged formatting.
        /// </summary>
        internal static string CreatePictureBulletImageFingerprint(
            OfficeXmlNode root, LegacyPptPartView ownerPart) {
            if (root.IsNone || ownerPart.IsNone) return string.Empty;
            var result = new System.Text.StringBuilder();
            foreach (OfficeXmlNode bullet in root.SchemaDescendants(LegacyPptXml.A, "buBlip")) {
                string? relationshipId = bullet.Child(LegacyPptXml.A, "blip")
                    .Attribute(LegacyPptXml.R, "embed");
                result.Append("|buBlip:").Append(relationshipId);
                LegacyPptRelationshipView? image = null;
                foreach (LegacyPptRelationshipView relationship in ownerPart.Relationships) {
                    if (!LegacyPptPackageFingerprint.IsPartRelationship(relationship)
                        || !string.Equals(relationship.Id, relationshipId, StringComparison.Ordinal)) continue;
                    image = relationship;
                    break;
                }
                if (image is not { } found
                    || !found.IsType(LegacyPptPackageView.ImageRelationshipType)) {
                    result.Append(":missing");
                    continue;
                }
                result.Append(':').Append(found.TargetPart.ContentType).Append(':');
                result.Append(LegacyPptBackgroundProjectionFingerprint.HashBytes(
                    found.TargetPart.GetBytes()));
            }
            return result.ToString();
        }

        internal bool PlaceholderMatches(
            LegacyPptWriter.LegacyPptWriterPlaceholder? current) =>
            current == null ? Placeholder == null
                : current.IsEquivalentTo(Placeholder);

        private static bool IsEditableAnimation(LegacyPptAnimation animation,
            ISet<uint> projectableSoundIds) {
            const uint editableFlags = 0x00004055U;
            if ((animation.RawFlags & ~editableFlags) != 0
                || animation.OleVerb != 0
                || animation.RawUnused != 0
                || animation.HasSoundOverride
                || animation.SlideCount != ushort.MaxValue
                || animation.Automatic && animation.DelayMilliseconds < 0
                || !animation.Automatic && animation.DelayMilliseconds != 0
                || animation.PlaysOnShapeClick
                || animation.Synchronous
                || animation.HiddenWhileNotPlaying) return false;
            return !animation.PlaysSound
                || projectableSoundIds.Contains(animation.SoundIdReference);
        }

        private static bool IsEditableInteraction(LegacyPptInteraction interaction,
            ISet<uint> projectableSoundIds) {
            byte allowedFlags = interaction.Action ==
                LegacyPptInteractionAction.CustomShow ? (byte)0x07 : (byte)0x03;
            if (interaction.OleVerb != 0
                || (interaction.Flags & ~allowedFlags) != 0) return false;
            if (interaction.SoundIdReference != 0
                && !projectableSoundIds.Contains(interaction.SoundIdReference)) {
                return false;
            }
            if (interaction.Action == LegacyPptInteractionAction.Macro) {
                return !string.IsNullOrEmpty(interaction.Name)
                    && interaction.Jump == LegacyPptInteractionJump.None
                    && interaction.HyperlinkType == LegacyPptHyperlinkType.Nil
                    && interaction.HyperlinkIdReference == 0;
            }
            if (interaction.Action == LegacyPptInteractionAction.RunProgram) {
                return !string.IsNullOrEmpty(interaction.Name)
                    && interaction.Jump == LegacyPptInteractionJump.None
                    && interaction.HyperlinkType == LegacyPptHyperlinkType.Nil
                    && interaction.HyperlinkIdReference == 0
                    && Uri.TryCreate(interaction.Name, UriKind.RelativeOrAbsolute,
                        out _);
            }
            if (interaction.Action == LegacyPptInteractionAction.CustomShow) {
                return interaction.CustomShow?.IsEditable == true
                    && interaction.Jump == LegacyPptInteractionJump.None
                    && interaction.HyperlinkType == LegacyPptHyperlinkType.Nil
                    && interaction.HyperlinkIdReference == 0;
            }
            if (interaction.Action == LegacyPptInteractionAction.Jump) {
                return interaction.Jump != LegacyPptInteractionJump.None
                    && interaction.HyperlinkType == LegacyPptHyperlinkType.Nil
                    && interaction.HyperlinkIdReference == 0
                    && string.IsNullOrEmpty(interaction.Name);
            }
            if (interaction.Action != LegacyPptInteractionAction.Hyperlink) return false;
            if (interaction.Jump != LegacyPptInteractionJump.None
                || !string.IsNullOrEmpty(interaction.Name)
                || interaction.HyperlinkType == LegacyPptHyperlinkType.CustomShow
                || interaction.Hyperlink != null
                && interaction.Hyperlink.ExtensionFlags != 0) return false;
            return (interaction.HyperlinkType != LegacyPptHyperlinkType.SlideNumber
                    && interaction.Hyperlink?.Uri != null)
                || (interaction.HyperlinkType == LegacyPptHyperlinkType.SlideNumber
                    && interaction.Hyperlink?.IsInternalSlideTarget == true)
                || interaction.HyperlinkType == LegacyPptHyperlinkType.NextSlide
                || interaction.HyperlinkType == LegacyPptHyperlinkType.PreviousSlide
                || interaction.HyperlinkType == LegacyPptHyperlinkType.FirstSlide
                || interaction.HyperlinkType == LegacyPptHyperlinkType.LastSlide;
        }

        private static bool HasOverlappingTextTriggers(
            IReadOnlyList<LegacyPptTextInteraction> interactions) {
            foreach (IGrouping<LegacyPptInteractionTrigger, LegacyPptTextInteraction> group
                     in interactions.GroupBy(item => item.Interaction.Trigger)) {
                int previousEnd = -1;
                foreach (LegacyPptTextInteraction item in group.OrderBy(item => item.Start)) {
                    if (item.Start < previousEnd) return true;
                    previousEnd = Math.Max(previousEnd,
                        checked(item.Start + item.Length));
                }
            }
            return false;
        }
    }

    /// <summary>
    /// A projected shape as the preservation fingerprints read it: the engine shape, for the writer's checks and
    /// non-XML values, and its element and owning part in a <see cref="LegacyPptPackageView"/>, for its markup.
    /// </summary>
    internal readonly struct LegacyPptProjectedShape {
        internal LegacyPptProjectedShape(PowerPointShape shape, LegacyPptPartView part, OfficeXmlNode element) {
            Shape = shape ?? throw new ArgumentNullException(nameof(shape));
            Part = part;
            Element = element;
        }

        internal PowerPointShape Shape { get; }

        internal LegacyPptPartView Part { get; }

        internal OfficeXmlNode Element { get; }

        /// <summary>The shape with its element resolved in <paramref name="view"/>.</summary>
        internal static LegacyPptProjectedShape Resolve(LegacyPptPackageView view, PowerPointShape shape) {
            ArgumentNullException.ThrowIfNull(view);
            ArgumentNullException.ThrowIfNull(shape);
            string partName = shape.OwnerPartRef.Name;
            return new LegacyPptProjectedShape(shape, view.Part(partName),
                view.Resolve(partName, shape.XmlElement));
        }
    }
}
