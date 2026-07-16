using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private static bool TryRewriteShapeContainer(
            LegacyPptRecord shapeContainer, ProjectedShapeEdit edit,
            out byte[] bytes) {
            LegacyPptWriter.LegacyPptWriterTextBoxContent? textContent =
                null;
            if (edit.TextFormatting != null
                && !LegacyPptWriter.TryBuildTextBoxContent(
                    edit.TextFormatting,
                    edit.TextFonts
                        ?? throw new InvalidOperationException(
                            "A text-formatting edit has no font catalog."),
                    edit.PictureBullets
                        ?? LegacyPptWriter
                            .LegacyPptWriterPictureBulletCatalog.Empty,
                    out textContent, out _)) {
                bytes = shapeContainer.CopyRecordBytes();
                return false;
            }
            if (textContent != null && shapeContainer.Children.Count(
                    child => child.Type == OfficeArtClientData) > 1) {
                bytes = shapeContainer.CopyRecordBytes();
                return false;
            }
            var children = new List<byte[]>(shapeContainer.Children.Count + 1);
            bool patchedAnchor = !edit.Bounds.HasValue;
            bool rewritePrimaryFopt = edit.ShapeTransform != null
                || edit.ShapeGeometry != null
                || edit.ShapeVisualStyle != null
                || edit.ShapeVisibility != null
                || edit.ShapeMetadata != null
                || edit.PictureFormatting != null
                || edit.TextFrame != null;
            bool patchedPrimaryFopt = !rewritePrimaryFopt;
            bool patchedFsp = edit.ShapeTransform == null;
            bool patchedGroupCoordinate = edit.GroupCoordinate == null;
            bool patchedPictureRecolor = edit.PictureFormatting == null;
            bool hasPrimaryFopt = shapeContainer.Children.Any(
                child => child.Type == OfficeArtFopt);
            bool hasTertiaryFopt = shapeContainer.Children.Any(
                child => child.Type == OfficeArtTertiaryFopt);
            bool patchedText = edit.Text == null
                && edit.TextFormatting == null
                && edit.Interactions?.RewriteTextInteractions != true;
            bool patchedShapeInteractions = edit.Interactions?
                .RewriteShapeInteractions != true;
            bool patchedAnimation = !edit.RewriteAnimation;
            bool patchedPlaceholder = !edit.RewritePlaceholder;
            bool patchedStyle9 = textContent == null;
            bool appendedShapeInteractions = false;
            bool appendedAnimation = false;
            bool sawClientData = false;
            foreach (LegacyPptRecord child in shapeContainer.Children) {
                if (!patchedGroupCoordinate
                    && child.Type == OfficeArtFspgr) {
                    children.Add(LegacyPptWriter
                        .BuildPreservedGroupCoordinateRecord(child,
                            edit.GroupCoordinate!));
                    patchedGroupCoordinate = true;
                } else if (child.Type == OfficeArtFsp
                    && (!patchedFsp
                        || !hasPrimaryFopt && !patchedPrimaryFopt)) {
                    children.Add(!patchedFsp
                        ? LegacyPptWriter.BuildPreservedFspRecord(
                            child, edit.ShapeTransform!)
                        : child.CopyRecordBytes());
                    patchedFsp = true;
                    if (!hasPrimaryFopt && !patchedPrimaryFopt) {
                        byte[]? primary = LegacyPptWriter
                            .BuildPreservedShapeFoptRecord(null,
                                edit.ShapeTransform
                                    ?? edit.ShapeGeometry
                                    ?? edit.ShapeVisualStyle
                                    ?? edit.ShapeVisibility
                                    ?? edit.ShapeMetadata
                                    ?? (PowerPointShape?)edit.PictureFormatting
                                    ?? edit.TextFrame!,
                                edit.ShapeTransform != null,
                                edit.ShapeGeometry != null,
                                edit.ShapeVisualStyle != null,
                                edit.PictureFormatting != null,
                                edit.TextFrame != null,
                                edit.ShapeMetadata != null,
                                edit.ShapeVisibility != null);
                        if (primary != null) children.Add(primary);
                        patchedPrimaryFopt = true;
                    }
                } else if (!patchedAnchor && edit.Bounds.HasValue
                    && (child.Type == OfficeArtClientAnchor
                        || child.Type == OfficeArtChildAnchor)) {
                    children.Add(BuildAnchor(child.Type, child.Instance,
                        edit.Bounds.Value));
                    patchedAnchor = true;
                } else if (!patchedPrimaryFopt
                           && child.Type == OfficeArtFopt) {
                    byte[]? primary = LegacyPptWriter
                        .BuildPreservedShapeFoptRecord(child,
                            edit.ShapeTransform ?? edit.ShapeVisualStyle
                                ?? edit.ShapeVisibility
                                ?? edit.ShapeMetadata
                                ?? edit.ShapeGeometry
                                ?? (PowerPointShape?)edit.PictureFormatting
                                ?? edit.TextFrame!,
                            edit.ShapeTransform != null,
                            edit.ShapeGeometry != null,
                            edit.ShapeVisualStyle != null,
                            edit.PictureFormatting != null,
                            edit.TextFrame != null,
                            edit.ShapeMetadata != null,
                            edit.ShapeVisibility != null);
                    if (primary != null) children.Add(primary);
                    patchedPrimaryFopt = true;
                    if (edit.PictureFormatting != null
                        && !hasTertiaryFopt) {
                        byte[]? tertiary = LegacyPptWriter
                            .BuildPreservedPictureTertiaryFoptRecord(null,
                                edit.PictureFormatting!);
                        if (tertiary != null) children.Add(tertiary);
                        patchedPictureRecolor = true;
                    }
                } else if (child.Type == OfficeArtTertiaryFopt
                           && (!patchedPictureRecolor
                               || edit.ShapeVisibility != null)) {
                    byte[]? tertiary = LegacyPptWriter
                        .BuildPreservedTertiaryFoptRecord(child,
                            patchedPictureRecolor
                                ? null
                                : edit.PictureFormatting,
                            edit.ShapeVisibility);
                    if (tertiary != null) children.Add(tertiary);
                    if (edit.PictureFormatting != null) {
                        patchedPictureRecolor = true;
                    }
                } else if (!patchedText
                           && child.Type == OfficeArtClientTextbox) {
                    bool rewritten = edit.TextFormatting != null
                        ? TryRewriteTextBoxFormatting(child, textContent!,
                            edit.Interactions?.RewriteTextInteractions
                                == true,
                            edit.Interactions?.Interactions.TextInteractions
                                ?? Array.Empty<LegacyPptWriter
                                    .LegacyPptWriterTextInteraction>(),
                            out byte[] textbox)
                        : edit.Interactions?
                        .RewriteTextInteractions == true
                        ? TryRewriteTextInteractions(child,
                            edit.OriginalText, edit.Text,
                            edit.Interactions.Interactions.TextInteractions,
                            out textbox)
                        : TryRewriteTextBox(child, edit.OriginalText,
                            edit.Text!, out textbox);
                    if (!rewritten) {
                        bytes = shapeContainer.CopyRecordBytes();
                        return false;
                    }
                    children.Add(textbox);
                    patchedText = true;
                } else if (child.Type == OfficeArtClientData
                           && (edit.Interactions?
                                   .RewriteShapeInteractions == true
                               || edit.RewriteAnimation
                               || edit.RewritePlaceholder
                               || textContent != null)) {
                    sawClientData = true;
                    byte[] clientData = child.CopyRecordBytes();
                    LegacyPptRecord rewrittenClientData = child;
                    if (textContent != null) {
                        if (!TryRewriteClientDataStyle9(
                                rewrittenClientData,
                                textContent.Style9Record,
                                out clientData)) {
                            bytes = shapeContainer.CopyRecordBytes();
                            return false;
                        }
                        rewrittenClientData = LegacyPptRecordReader
                            .ReadSingle(clientData, 0,
                                new LegacyPptImportOptions());
                        patchedStyle9 = true;
                    }
                    if (edit.Interactions?.RewriteShapeInteractions == true
                        && !TryRewriteClientDataInteractions(
                            rewrittenClientData,
                            edit.Interactions.Interactions.ShapeInteractions,
                            append: !appendedShapeInteractions,
                            out clientData)) {
                        bytes = shapeContainer.CopyRecordBytes();
                        return false;
                    }
                    rewrittenClientData = LegacyPptRecordReader.ReadSingle(
                            clientData, 0,
                            new LegacyPptImportOptions());
                    if (edit.RewriteAnimation
                        && !TryRewriteClientDataAnimation(
                            rewrittenClientData,
                            append: !appendedAnimation
                                ? edit.Animation
                                : null,
                            out clientData)) {
                        bytes = shapeContainer.CopyRecordBytes();
                        return false;
                    }
                    rewrittenClientData = LegacyPptRecordReader.ReadSingle(
                        clientData, 0, new LegacyPptImportOptions());
                    if (edit.RewritePlaceholder
                        && !TryRewriteClientDataPlaceholder(
                            rewrittenClientData, edit.Placeholder,
                            out clientData)) {
                        bytes = shapeContainer.CopyRecordBytes();
                        return false;
                    }
                    children.Add(clientData);
                    if (edit.Interactions?
                            .RewriteShapeInteractions == true) {
                        appendedShapeInteractions = true;
                        patchedShapeInteractions = true;
                    }
                    if (edit.RewriteAnimation) {
                        appendedAnimation |= edit.Animation != null;
                        patchedAnimation = true;
                    }
                    if (edit.RewritePlaceholder) {
                        patchedPlaceholder = true;
                    }
                } else {
                    children.Add(child.CopyRecordBytes());
                }
            }
            if (!sawClientData
                && (edit.Interactions?.RewriteShapeInteractions == true
                    || edit.RewriteAnimation || edit.RewritePlaceholder
                    || textContent != null)) {
                var clientChildren = new List<byte[]>();
                if (edit.Placeholder != null) {
                    clientChildren.Add(LegacyPptWriter
                        .BuildPlaceholderAtom(edit.Placeholder.Position,
                            edit.Placeholder.Type,
                            edit.Placeholder.Size));
                }
                if (edit.Animation != null) {
                    clientChildren.Add(LegacyPptWriter
                        .BuildAnimationInfoRecord(edit.Animation));
                }
                if (edit.Interactions?
                        .RewriteShapeInteractions == true) {
                    clientChildren.AddRange(edit.Interactions.Interactions
                        .ShapeInteractions.Select(LegacyPptWriter
                            .BuildInteractiveInfoRecord));
                    patchedShapeInteractions = true;
                }
                if (textContent?.Style9Record != null) {
                    clientChildren.Add(LegacyPptWriter
                        .BuildShapePpt9ProgrammableTagsRecord(
                            textContent.Style9Record));
                }
                patchedStyle9 = true;
                patchedAnimation = true;
                patchedPlaceholder = true;
                if (clientChildren.Count > 0) {
                    children.Add(BuildRecord(version: 0x0F, instance: 0,
                        OfficeArtClientData, Concat(clientChildren)));
                }
            }
            if (!patchedAnchor || !patchedPrimaryFopt || !patchedFsp
                || !patchedGroupCoordinate
                || !patchedPictureRecolor || !patchedText
                || !patchedShapeInteractions || !patchedAnimation
                || !patchedPlaceholder || !patchedStyle9) {
                bytes = shapeContainer.CopyRecordBytes();
                return false;
            }
            bytes = BuildRecord(shapeContainer.Version,
                shapeContainer.Instance, shapeContainer.Type,
                Concat(children));
            return true;
        }
    }
}
