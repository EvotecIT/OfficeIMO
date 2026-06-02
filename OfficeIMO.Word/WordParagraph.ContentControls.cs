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
        internal WordStructuredDocumentTag? StructuredDocumentTag {
            get {
                if (_stdRun != null) {
                    return new WordStructuredDocumentTag(_document, _paragraph, _stdRun);
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the checkbox contained in this paragraph, if present.
        /// </summary>
        public WordCheckBox? CheckBox {
            get {
                if (_stdRun != null && _stdRun.SdtProperties?.Elements<DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox>().Any() == true) {
                    return new WordCheckBox(_document, _paragraph, _stdRun);
                }

                return null;
            }
        }


        /// <summary>
        /// Gets the date picker contained in this paragraph, if present.
        /// </summary>
        public WordDatePicker? DatePicker {
            get {
                if (_stdRun != null && _stdRun.SdtProperties?.Elements<SdtContentDate>().Any() == true) {
                    return new WordDatePicker(_document, _paragraph, _stdRun);
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the dropdown list contained in this paragraph, if present.
        /// </summary>
        public WordDropDownList? DropDownList {
            get {
                if (_stdRun != null && _stdRun.SdtProperties?.Elements<SdtContentDropDownList>().Any() == true) {
                    return new WordDropDownList(_document, _paragraph, _stdRun);
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the combo box contained in this paragraph, if present.
        /// </summary>
        public WordComboBox? ComboBox {
            get {
                if (_stdRun != null && _stdRun.SdtProperties?.Elements<SdtContentComboBox>().Any() == true) {
                    return new WordComboBox(_document, _paragraph, _stdRun);
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the picture content control contained in this paragraph, if present.
        /// </summary>
        public WordPictureControl? PictureControl {
            get {
                if (_stdRun != null && _stdRun.SdtProperties?.Elements<SdtContentPicture>().Any() == true) {
                    return new WordPictureControl(_document, _paragraph, _stdRun);
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the repeating section contained in this paragraph, if present.
        /// </summary>
        public WordRepeatingSection? RepeatingSection =>
            _stdRun is not null && _stdRun.SdtProperties?.Elements<W15.SdtRepeatedSection>().Any() is true
                ? new WordRepeatingSection(_document, _paragraph, _stdRun)
                : null;
        /// <summary>
        /// Gets a value indicating whether the paragraph holds a structured document tag.
        /// </summary>
        public bool IsStructuredDocumentTag => StructuredDocumentTag is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a checkbox control.
        /// </summary>
        public bool IsCheckBox => CheckBox is not null;


        /// <summary>
        /// Gets a value indicating whether the paragraph contains a date picker control.
        /// </summary>
        public bool IsDatePicker => DatePicker is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a dropdown list control.
        /// </summary>
        public bool IsDropDownList => DropDownList is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a combo box control.
        /// </summary>
        public bool IsComboBox => ComboBox is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a picture control.
        /// </summary>
        public bool IsPictureControl => PictureControl is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a repeating section control.
        /// </summary>
        public bool IsRepeatingSection => RepeatingSection is not null;
    }
}
