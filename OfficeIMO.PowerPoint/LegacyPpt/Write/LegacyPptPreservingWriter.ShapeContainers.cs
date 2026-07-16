using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private static bool TryRewriteShapeContainer(
            LegacyPptRecord shapeContainer, ProjectedShapeEdit edit,
            out byte[] bytes) {
            var children = new List<byte[]>(shapeContainer.Children.Count + 1);
            bool patchedAnchor = !edit.Bounds.HasValue;
            bool patchedPictureFormatting = edit.PictureFormatting == null;
            bool patchedPictureRecolor = edit.PictureFormatting == null;
            bool hasPictureTertiaryFopt = shapeContainer.Children.Any(
                child => child.Type == OfficeArtTertiaryFopt);
            bool patchedText = edit.Text == null
                && edit.Interactions?.RewriteTextInteractions != true;
            bool patchedShapeInteractions = edit.Interactions?
                .RewriteShapeInteractions != true;
            bool patchedAnimation = !edit.RewriteAnimation;
            bool patchedPlaceholder = !edit.RewritePlaceholder;
            bool appendedShapeInteractions = false;
            bool appendedAnimation = false;
            bool sawClientData = false;
            foreach (LegacyPptRecord child in shapeContainer.Children) {
                if (!patchedAnchor && edit.Bounds.HasValue
                    && (child.Type == OfficeArtClientAnchor
                        || child.Type == OfficeArtChildAnchor)) {
                    children.Add(BuildAnchor(child.Type, child.Instance,
                        edit.Bounds.Value));
                    patchedAnchor = true;
                } else if (!patchedPictureFormatting
                           && child.Type == OfficeArtFopt) {
                    children.Add(LegacyPptWriter
                        .BuildPreservedPictureFoptRecord(child,
                            edit.PictureFormatting!));
                    patchedPictureFormatting = true;
                    if (!hasPictureTertiaryFopt) {
                        byte[]? tertiary = LegacyPptWriter
                            .BuildPreservedPictureTertiaryFoptRecord(null,
                                edit.PictureFormatting!);
                        if (tertiary != null) children.Add(tertiary);
                        patchedPictureRecolor = true;
                    }
                } else if (!patchedPictureRecolor
                           && child.Type == OfficeArtTertiaryFopt) {
                    byte[]? tertiary = LegacyPptWriter
                        .BuildPreservedPictureTertiaryFoptRecord(child,
                            edit.PictureFormatting!);
                    if (tertiary != null) children.Add(tertiary);
                    patchedPictureRecolor = true;
                } else if (!patchedText
                           && child.Type == OfficeArtClientTextbox) {
                    bool rewritten = edit.Interactions?
                        .RewriteTextInteractions == true
                        ? TryRewriteTextInteractions(child,
                            edit.OriginalText, edit.Text,
                            edit.Interactions.Interactions.TextInteractions,
                            out byte[] textbox)
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
                               || edit.RewritePlaceholder)) {
                    sawClientData = true;
                    byte[] clientData = child.CopyRecordBytes();
                    if (edit.Interactions?.RewriteShapeInteractions == true
                        && !TryRewriteClientDataInteractions(child,
                            edit.Interactions.Interactions.ShapeInteractions,
                            append: !appendedShapeInteractions,
                            out clientData)) {
                        bytes = shapeContainer.CopyRecordBytes();
                        return false;
                    }
                    LegacyPptRecord rewrittenClientData =
                        LegacyPptRecordReader.ReadSingle(clientData, 0,
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
                    || edit.RewriteAnimation || edit.RewritePlaceholder)) {
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
                patchedAnimation = true;
                patchedPlaceholder = true;
                if (clientChildren.Count > 0) {
                    children.Add(BuildRecord(version: 0x0F, instance: 0,
                        OfficeArtClientData, Concat(clientChildren)));
                }
            }
            if (!patchedAnchor || !patchedPictureFormatting
                || !patchedPictureRecolor || !patchedText
                || !patchedShapeInteractions || !patchedAnimation
                || !patchedPlaceholder) {
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
