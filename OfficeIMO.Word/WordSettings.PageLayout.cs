using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordSettings {
        /// <summary>Gets or sets whether the document gutter is placed at the top of each page.</summary>
        public bool GutterAtTop {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart?.DocumentSettingsPart?.Settings;
                var gutterAtTop = settings?.GetFirstChild<GutterAtTop>();
                return gutterAtTop != null && (gutterAtTop.Val?.Value ?? true);
            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart?.DocumentSettingsPart?.Settings;
                if (settings == null) return;
                var gutterAtTop = settings.GetFirstChild<GutterAtTop>();
                if (gutterAtTop == null) {
                    gutterAtTop = new GutterAtTop();
                    settings.AddChild(gutterAtTop, true);
                }
                gutterAtTop.Val = value;
            }
        }

        /// <summary>Gets or sets whether odd and even pages use mirrored inside/outside margins.</summary>
        public bool MirrorMargins {
            get {
                var settings = _document._wordprocessingDocument.MainDocumentPart?.DocumentSettingsPart?.Settings;
                var mirrorMargins = settings?.GetFirstChild<MirrorMargins>();
                return mirrorMargins != null && (mirrorMargins.Val?.Value ?? true);
            }
            set {
                var settings = _document._wordprocessingDocument.MainDocumentPart?.DocumentSettingsPart?.Settings;
                if (settings == null) return;
                var mirrorMargins = settings.GetFirstChild<MirrorMargins>();
                if (value) {
                    if (mirrorMargins == null) {
                        mirrorMargins = new MirrorMargins();
                        settings.AddChild(mirrorMargins, true);
                    }
                    mirrorMargins.Val = true;
                } else {
                    mirrorMargins?.Remove();
                }
            }
        }
    }
}
