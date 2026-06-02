using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Hyperlink = DocumentFormat.OpenXml.Wordprocessing.Hyperlink;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using SdtContentPicture = DocumentFormat.OpenXml.Wordprocessing.SdtContentPicture;
using TabStop = DocumentFormat.OpenXml.Wordprocessing.TabStop;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using V = DocumentFormat.OpenXml.Vml;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Get PageBreaks within Paragraph
        /// </summary>
        public WordBreak? PageBreak {
            get {
                if (_run != null) {
                    var brake = _run.ChildElements.OfType<Break>().FirstOrDefault();
                    if (brake != null && brake.Type != null && brake.Type.Value == BreakValues.Page) {
                        return new WordBreak(_document, _paragraph, _run);
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Get Breaks within Paragraph
        /// </summary>
        public WordBreak? Break {
            get {
                if (_run != null) {
                    var brake = _run.ChildElements.OfType<Break>().FirstOrDefault();
                    if (brake != null) {
                        return new WordBreak(_document, _paragraph, _run);
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the <see cref="WordTabChar"/> representing a tab character in the current run, or <c>null</c> if none is present.
        /// </summary>
        public WordTabChar? Tab {
            get {
                if (_run != null) {
                    var tabChar = _run.ChildElements.OfType<TabChar>().FirstOrDefault();
                    if (tabChar != null) {
                        return new WordTabChar(_document, _paragraph, _run);
                    }
                }

                return null;
            }
        }
        /// <summary>
        /// Gets a value indicating whether the run within the paragraph contains a tab character.
        /// </summary>
        public bool IsTab => Tab is not null;
        /// <summary>
        /// Gets all tab stops defined on the paragraph.
        /// </summary>
        public List<WordTabStop> TabStops {
            get {
                List<WordTabStop> list = new List<WordTabStop>();
                if (_paragraph is not null && _paragraphProperties is not null) {
                    if (_paragraphProperties.Tabs is not null) {
                        foreach (TabStop tab in _paragraphProperties.Tabs) {
                            list.Add(new WordTabStop(this, tab));
                        }
                    }
                }
                return list;
            }
        }
    }
}
