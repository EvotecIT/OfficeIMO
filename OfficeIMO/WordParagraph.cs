using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;

namespace OfficeIMO {
    public class WordParagraph {
        internal Paragraph _paragraph = null;
        internal RunProperties _runProperties = null;
        internal Text _text = null;
        internal Run _run = null;

        public string Text {
            get {
                return _text.Text;
            }
            set {
                _text.Text = value;
            }
        }
        public bool Bold   // property
        {
            get
            {
                if (_runProperties.Bold != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            set
            {
                if (value != true)
                {
                    _runProperties.Bold = null;
                }
                else
                {
                    _runProperties.Bold = new Bold();
                }
            }
        }
        public bool Italic   // property
        {
            get
            {
                if (_runProperties.Italic != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            set
            {
                if (value != true)
                {
                    _runProperties.Italic = null;
                }
                else
                {
                    _runProperties.Italic = new Italic();
                }
            }
        }
        public UnderlineValues? Underline   // property
        {
            get
            {
                if (_runProperties.Underline.Val != null)
                {

                    return _runProperties.Underline.Val;
                }
                else {
                    return null;
                }
            }
            set
            {
                if (_runProperties.Underline == null) {
                    _runProperties.Underline = new Underline();
                } else {

                }
                _runProperties.Underline.Val = value;
            }
        }

        public WordParagraph(bool newParagraph = true) {
            this._run = new Run();
            this._runProperties = new RunProperties();
            this._text = new Text();
            // this ensures spaces are preserved between runs
            this._text.Space = SpaceProcessingModeValues.Preserve;
            this._run.AppendChild(_runProperties);
            this._run.AppendChild(_text);
            if (newParagraph) {
                this._paragraph = new Paragraph();
                this._paragraph.AppendChild(_run);
            }
        }

        public WordParagraph(string text) {
            WordParagraph paragraph = new WordParagraph();
            paragraph.Text = text;
        }

        public WordParagraph InsertText(string text) {
            //DocumentFormat.OpenXml.Wordprocessing.Run run = new DocumentFormat.OpenXml.Wordprocessing.Run();
            //DocumentFormat.OpenXml.Wordprocessing.Text textProp = new DocumentFormat.OpenXml.Wordprocessing.Text();
            //textProp.Text = text;
            //run.AppendChild(textProp);
            //this.paragraph.Append(run);
            //return this;

           // WordParagraph paragraph = new WordParagraph();
            //this.paragraph.Text = text;
            //return this.paragraph;
            this._text.Text = text;
            return this;
        }

        public WordParagraph AppendText(string text) {
            WordParagraph wordParagraph = new WordParagraph(false);
            wordParagraph.Text = text;
            this._paragraph.Append(wordParagraph._run);
            return wordParagraph;
        }

        public void InsertText(string text, WordFormatting formatting) {
          //  if (this._paragraph == null) {
           //     this._paragraph = new Paragraph();
           // }

           // this._paragraph.Append(formatting._runProperties);


            //var run = this.paragraph.ChildElements.First<Run>();
            //run.RunProperties = runProperties;

            return;
        }

    }
}