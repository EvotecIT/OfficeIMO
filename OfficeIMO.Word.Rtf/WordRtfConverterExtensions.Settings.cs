namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void CopyDocumentSettings(WordDocument source, RtfDocument destination) {
        if (source.Settings.DefaultTabStop > 0) {
            destination.Settings.DefaultTabWidthTwips = source.Settings.DefaultTabStop;
        }

        destination.Settings.ViewScale = source.Settings.ZoomPercentage;
        CopyDocumentProtection(source.Settings.ProtectionType, destination.Settings);

        if (source.Sections.Any(section => section.DocumentOddEvenSettingEnabled)) {
            destination.Settings.FacingPages = true;
        }

        if (source.Settings.MirrorMargins ||
            source.Sections.Any(section => section.PageSetup.Margins.Type == WordMargin.Mirrored)) {
            destination.Settings.MirrorMargins = true;
        }
    }

    private static void ApplyDocumentSettings(RtfDocument source, WordDocument destination) {
        if (source.Settings.DefaultFontId is int fontId && TryGetFontName(source, fontId, out string? fontName)) {
            destination.Settings.FontFamily = fontName;
        }

        if (source.Settings.DefaultTabWidthTwips.HasValue) {
            destination.Settings.DefaultTabStop = source.Settings.DefaultTabWidthTwips.Value;
        }

        if (source.Settings.ViewScale.HasValue) {
            destination.Settings.ZoomPercentage = source.Settings.ViewScale.Value;
        }

        ApplyDocumentProtection(source.Settings, destination);

        if (source.Settings.FacingPages == true) {
            OfficeXmlNode settings = destination.SettingsNode();
            settings.RemoveChildren(WordXml.W, "evenAndOddHeaders");
            settings.Append(settings.CreateElement(WordXml.W, "evenAndOddHeaders"));
        }

        if (source.Settings.MirrorMargins == true) {
            destination.Settings.MirrorMargins = true;
        }
    }

    private static void CopyDocumentProtection(WordDocumentProtectionType? source, RtfDocumentSettings destination) {
        if (!source.HasValue) {
            return;
        }

        if (source.Value == WordDocumentProtectionType.Forms) {
            destination.FormProtection = true;
        } else if (source.Value == WordDocumentProtectionType.TrackedChanges) {
            destination.RevisionProtection = true;
        } else if (source.Value == WordDocumentProtectionType.Comments) {
            destination.AnnotationProtection = true;
        } else if (source.Value == WordDocumentProtectionType.ReadOnly) {
            destination.ReadOnlyProtection = true;
        }
    }

    private static void ApplyDocumentProtection(RtfDocumentSettings source, WordDocument destination) {
        string? protection = ToWordDocumentProtection(source);
        if (protection == null) {
            return;
        }

        OfficeXmlNode settings = destination.SettingsNode();
        if (settings.IsNone) {
            return;
        }

        settings.RemoveChildren(WordXml.W, "documentProtection");
        OfficeXmlNode documentProtection = settings.CreateElement(WordXml.W, "documentProtection");
        documentProtection.SetAttribute(WordXml.W, "edit", protection);
        settings.Append(documentProtection);
    }

    // The w:edit token of w:documentProtection.
    private static string? ToWordDocumentProtection(RtfDocumentSettings source) {
        if (source.FormProtection == true) {
            return "forms";
        }

        if (source.RevisionProtection == true) {
            return "trackedChanges";
        }

        if (source.AnnotationProtection == true) {
            return "comments";
        }

        if (source.ReadOnlyProtection == true) {
            return "readOnly";
        }

        return null;
    }
}
