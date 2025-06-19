using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace OfficeIMO.Word {
    public class WordCheckBox : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        internal readonly SdtRun _sdtRun;

        internal WordCheckBox(WordDocument document, Paragraph paragraph, SdtRun sdtRun) {
            _document = document;
            _paragraph = paragraph;
            _sdtRun = sdtRun;
        }

        public bool Checked {
            get {
                var cb = _sdtRun.SdtProperties.GetFirstChild<W14.SdtContentCheckBox>();
                var chk = cb?.GetFirstChild<W14.Checked>();
                return chk != null && chk.Val == W14.OnOffValues.One;
            }
            set {
                var cb = _sdtRun.SdtProperties.GetFirstChild<W14.SdtContentCheckBox>();
                if (cb == null) return;
                var chk = cb.GetFirstChild<W14.Checked>() ?? cb.AppendChild(new W14.Checked());
                chk.Val = value ? W14.OnOffValues.One : W14.OnOffValues.Zero;
                var run = _sdtRun.SdtContentRun.GetFirstChild<Run>();
                var text = run?.GetFirstChild<Text>();
                if (text != null) {
                    text.Text = value ? "☒" : "☐";
                }
            }
        }

        public void Remove() {
            _sdtRun.Remove();
        }

        public static WordParagraph AddCheckBox(WordParagraph paragraph, bool isChecked = false, string alias = null) {
            SdtRun sdtRun = new SdtRun();
            SdtProperties sdtProperties = new SdtProperties();
            SdtId sdtId = new SdtId() { Val = new Random().Next() };
            var checkBox = new W14.SdtContentCheckBox();
            var checkedElement = new W14.Checked() { Val = isChecked ? W14.OnOffValues.One : W14.OnOffValues.Zero };
            var checkedState = new W14.CheckedState() { Font = "MS Gothic", Val = "2612" };
            var uncheckedState = new W14.UncheckedState() { Font = "MS Gothic", Val = "2610" };
            checkBox.Append(checkedElement);
            checkBox.Append(checkedState);
            checkBox.Append(uncheckedState);
            sdtProperties.Append(sdtId);
            if (!string.IsNullOrEmpty(alias)) {
                sdtProperties.Append(new SdtAlias() { Val = alias });
            }
            sdtProperties.Append(checkBox);

            SdtContentRun contentRun = new SdtContentRun();
            Run run = new Run();
            RunProperties runProps = new RunProperties();
            RunFonts fonts = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "MS Gothic", HighAnsi = "MS Gothic", EastAsia = "MS Gothic" };
            runProps.Append(fonts);
            Text text = new Text(isChecked ? "☒" : "☐");
            run.Append(runProps);
            run.Append(text);
            contentRun.Append(run);

            sdtRun.Append(sdtProperties);
            sdtRun.Append(contentRun);

            paragraph._paragraph.Append(sdtRun);
            paragraph._stdRun = sdtRun;
            return paragraph;
        }
    }
}
