using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void CopyDocumentSettings(WordDocument source, RtfDocument destination) {
        if (source.Settings.DefaultTabStop > 0) {
            destination.Settings.DefaultTabWidthTwips = source.Settings.DefaultTabStop;
        }

        destination.Settings.ViewScale = source.Settings.ZoomPercentage;
        CopyDocumentProtection(source.Settings.ProtectionType, destination.Settings);

        if (source.Sections.Any(section => section.DifferentOddAndEvenPages)) {
            destination.Settings.FacingPages = true;
        }

        if (source.Settings.MirrorMargins ||
            source.Sections.Any(section => section.Margins.Type == WordMargin.Mirrored)) {
            destination.Settings.MirrorMargins = true;
        }
    }

    private static void ApplyDocumentSettings(RtfDocument source, WordDocument destination) {
        if (source.Settings.DefaultTabWidthTwips.HasValue) {
            destination.Settings.DefaultTabStop = source.Settings.DefaultTabWidthTwips.Value;
        }

        if (source.Settings.ViewScale.HasValue) {
            destination.Settings.ZoomPercentage = source.Settings.ViewScale.Value;
        }

        ApplyDocumentProtection(source.Settings, destination);

        if (source.Settings.FacingPages == true) {
            destination.DifferentOddAndEvenPages = true;
        }

        if (source.Settings.MirrorMargins == true) {
            destination.Settings.MirrorMargins = true;
        }
    }

    private static void CopyDocumentProtection(DocumentProtectionValues? source, RtfDocumentSettings destination) {
        if (!source.HasValue) {
            return;
        }

        if (source.Value == DocumentProtectionValues.Forms) {
            destination.FormProtection = true;
        } else if (source.Value == DocumentProtectionValues.TrackedChanges) {
            destination.RevisionProtection = true;
        } else if (source.Value == DocumentProtectionValues.Comments) {
            destination.AnnotationProtection = true;
        } else if (source.Value == DocumentProtectionValues.ReadOnly) {
            destination.ReadOnlyProtection = true;
        }
    }

    private static void ApplyDocumentProtection(RtfDocumentSettings source, WordDocument destination) {
        DocumentProtectionValues? protection = ToWordDocumentProtection(source);
        if (!protection.HasValue) {
            return;
        }

        Settings? settings = destination._wordprocessingDocument.MainDocumentPart?.DocumentSettingsPart?.Settings;
        if (settings == null) {
            return;
        }

        settings.RemoveAllChildren<DocumentProtection>();
        settings.Append(new DocumentProtection { Edit = protection.Value });
    }

    private static DocumentProtectionValues? ToWordDocumentProtection(RtfDocumentSettings source) {
        if (source.FormProtection == true) {
            return DocumentProtectionValues.Forms;
        }

        if (source.RevisionProtection == true) {
            return DocumentProtectionValues.TrackedChanges;
        }

        if (source.AnnotationProtection == true) {
            return DocumentProtectionValues.Comments;
        }

        if (source.ReadOnlyProtection == true) {
            return DocumentProtectionValues.ReadOnly;
        }

        return null;
    }
}
