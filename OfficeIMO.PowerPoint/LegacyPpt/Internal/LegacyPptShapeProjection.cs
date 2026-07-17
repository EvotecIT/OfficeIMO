using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

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

        internal bool ShapeTransformMatches(PowerPointShape shape) =>
            string.Equals(ShapeTransformFingerprint,
                CreateShapeTransformFingerprint(shape),
                StringComparison.Ordinal);

        internal static string? CreateShapeTransformFingerprint(
            PowerPointShape shape) {
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

        internal bool TextFrameMatches(PowerPointShape shape) =>
            string.Equals(TextFrameFingerprint,
                CreateTextFrameFingerprint(shape),
                StringComparison.Ordinal);

        internal static string? CreateTextFrameFingerprint(
            PowerPointShape shape) {
            if (shape is not PowerPointTextBox textBox
                || !LegacyPptWriter.TryReadTextFrameForWrite(textBox,
                    out _, out _)
                || textBox.Element is not P.Shape source) {
                return null;
            }
            return LegacyPptTextProjection.CreateTextFrameFingerprint(
                source.TextBody);
        }

        internal bool ShapeGeometryMatches(PowerPointShape shape) =>
            string.Equals(ShapeGeometryFingerprint,
                CreateShapeGeometryFingerprint(shape),
                StringComparison.Ordinal);

        internal static string? CreateShapeGeometryFingerprint(
            PowerPointShape shape) {
            if (!LegacyPptWriter.TryReadOfficeArtShapeType(shape,
                    requireConnector: false, out ushort shapeType, out _)
                || shapeType is not 2 and not 23
                || !LegacyPptWriter.TryReadShapeGeometry(shape, shapeType,
                    out _, out _)) {
                return null;
            }
            A.AdjustValueList? values = shape.Element
                .Descendants<A.PresetGeometry>().FirstOrDefault()?
                .AdjustValueList;
            return string.Concat(values?.Elements<A.ShapeGuide>()
                .Select(guide => guide.OuterXml)
                ?? Enumerable.Empty<string>());
        }

        internal bool GroupCoordinateMatches(PowerPointShape shape) =>
            string.Equals(GroupCoordinateFingerprint,
                CreateGroupCoordinateFingerprint(shape),
                StringComparison.Ordinal);

        internal static string? CreateGroupCoordinateFingerprint(
            PowerPointShape shape) {
            if (shape is not PowerPointGroupShape group
                || !LegacyPptWriter.TryReadGroupForWrite(group,
                    out _, out _)) {
                return null;
            }
            A.TransformGroup transform = group.GroupShape
                .GroupShapeProperties!.TransformGroup!;
            return string.Join("\n",
                transform.ChildOffset!.X!.Value.ToString(
                    System.Globalization.CultureInfo.InvariantCulture),
                transform.ChildOffset.Y!.Value.ToString(
                    System.Globalization.CultureInfo.InvariantCulture),
                transform.ChildExtents!.Cx!.Value.ToString(
                    System.Globalization.CultureInfo.InvariantCulture),
                transform.ChildExtents.Cy!.Value.ToString(
                    System.Globalization.CultureInfo.InvariantCulture));
        }

        internal bool ShapeVisualStyleMatches(PowerPointShape shape) =>
            string.Equals(ShapeVisualStyleFingerprint,
                CreateShapeVisualStyleFingerprint(shape),
                StringComparison.Ordinal);

        internal static string? CreateShapeVisualStyleFingerprint(
            PowerPointShape shape) {
            if (!LegacyPptWriter.TryReadShapeVisualStyle(shape,
                    out _, out _)) {
                return null;
            }
            P.ShapeProperties? properties = shape.Element switch {
                P.Shape value => value.ShapeProperties,
                P.ConnectionShape value => value.ShapeProperties,
                P.Picture value => value.ShapeProperties,
                _ => null
            };
            string visual = string.Concat(properties?.ChildElements
                .Where(child => child is A.NoFill or A.SolidFill
                    or A.Outline or A.EffectList)
                .Select(child => child.OuterXml)
                ?? Enumerable.Empty<string>());
            return visual;
        }

        internal bool PictureFormattingMatches(PowerPointPicture picture) =>
            string.Equals(PictureFormattingFingerprint,
                CreatePictureFormattingFingerprint(picture),
                StringComparison.Ordinal);

        internal static string? CreatePictureFormattingFingerprint(
            PowerPointShape shape) {
            if (shape is not PowerPointPicture
                || shape is PowerPointMedia
                || shape.Element is not P.Picture picture) {
                return null;
            }
            string protection = picture.NonVisualPictureProperties?
                .NonVisualPictureDrawingProperties?.OuterXml
                ?? string.Empty;
            string crop = picture.BlipFill?.SourceRectangle?.OuterXml
                ?? string.Empty;
            string effects = string.Concat(picture.BlipFill?.Blip?
                .ChildElements.Select(child => child.OuterXml)
                ?? Enumerable.Empty<string>());
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
            string description = shape.Description ?? string.Empty;
            return string.Concat(name.Length.ToString(
                    System.Globalization.CultureInfo.InvariantCulture),
                ":", name, "\n", description.Length.ToString(
                    System.Globalization.CultureInfo.InvariantCulture),
                ":", description);
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
}
