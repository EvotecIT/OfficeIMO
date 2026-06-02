using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Reflection;
using System.Linq;
using System.Xml.Linq;
using MathParagraph = DocumentFormat.OpenXml.Math.Paragraph;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using V = DocumentFormat.OpenXml.Vml;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace OfficeIMO.Word {
    /// <summary>
    /// Contains public methods for editing paragraphs.
    /// </summary>
    public partial class WordParagraph {
        /// <summary>
        /// Add a table after this paragraph and return the table.
        /// </summary>
        /// <param name="rows">The number of rows in the table.</param>
        /// <param name="columns">The number of columns in the table.</param>
        /// <param name="tableStyle">The optional style to be applied to the new table, defaults to TableGrid.</param>
        /// <returns>The new added table.</returns>
        public WordTable AddTableAfter(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            WordTable wordTable = new WordTable(this._document, this, rows, columns, tableStyle, "After");
            return wordTable;
        }

        /// <summary>
        /// Add a table before this paragraph and return the table.
        /// </summary>
        /// <param name="rows">The number of rows in this table</param>
        /// <param name="columns">The number of columns in this table</param>
        /// <param name="tableStyle">The style of the table being added.</param>
        /// <returns>The new added table.</returns>
        public WordTable AddTableBefore(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            WordTable wordTable = new WordTable(this._document, this, rows, columns, tableStyle, "Before");
            return wordTable;
        }
        /// <summary>
        /// Embeds an object with a preview image.
        /// </summary>
        /// <param name="filePath">Path to the object file.</param>
        /// <param name="imageFilePath">Preview image path.</param>
        /// <param name="width">Optional width in points.</param>
        /// <param name="height">Optional height in points.</param>
        /// <returns>The paragraph containing the embedded object.</returns>
        public WordParagraph AddEmbeddedObject(string filePath, string imageFilePath, double? width = null, double? height = null) {
            var wordEmbeddedObject = new WordEmbeddedObject(this, this._document, filePath, imageFilePath, "", width, height);
            return this;
        }

        /// <summary>
        /// Embeds an object with custom options.
        /// </summary>
        /// <param name="filePath">Path to the object file.</param>
        /// <param name="options">Embedding options.</param>
        /// <returns>The paragraph containing the embedded object.</returns>
        public WordParagraph AddEmbeddedObject(string filePath, WordEmbeddedObjectOptions options) {
            var wordEmbeddedObject = new WordEmbeddedObject(this, this._document, filePath, options);
            return this;
        }

        /// <summary>
        /// Provides ability for configuration of Tabs in a paragraph
        /// by adding one or more TabStops
        /// </summary>
        /// <param name="position">The position of the tabs in the paragraph.</param>
        /// <param name="alignment">The optional alignment for the tabs.</param>
        /// <param name="leader">The optional rune to use before the tabs.</param>
        /// <returns>The added tabs.</returns>
        public WordTabStop AddTabStop(int position, TabStopValues? alignment = null, TabStopLeaderCharValues? leader = null) {
            alignment ??= TabStopValues.Left;
            leader ??= TabStopLeaderCharValues.None;
            var wordTab = new WordTabStop(this);
            wordTab.AddTab(position, alignment, leader);
            return wordTab;
        }

        /// <summary>
        /// Adds a Tab to a paragraph
        /// </summary>
        /// <returns>The paragraph this is called on.</returns>
        public WordParagraph AddTab() {
            var wordParagraph = WordTabChar.AddTab(this._document, this);
            return wordParagraph;
        }

        /// <summary>
        /// Add a list after this paragraph.
        /// </summary>
        /// <param name="style">The style of this list.</param>
        /// <returns>The new list.</returns>
        public WordList AddList(WordListStyle style) {
            WordList wordList = new WordList(this._document, this);
            wordList.AddList(style);
            return wordList;
        }

        /// <summary>
        /// Adds the chart to the document. The type of chart is determined by the type of data passed in.
        /// You can use multiple:
        /// .AddBar() to add a bar chart
        /// .AddLine() to add a line chart
        /// .AddPie() to add a pie chart
        /// .AddArea() to add an area chart
        /// .AddScatter() to add a scatter chart
        /// .AddRadar() to add a radar chart
        /// .AddBar3D() to add a 3-D bar chart.
        /// .AddPie3D() to add a 3-D pie chart.
        /// .AddLine3D() to add a 3-D line chart.
        /// You can't mix and match the types of charts, except bar and line which can coexist in a combo chart.
        /// </summary>
        /// <param name="title">The title.</param>
        /// <param name="roundedCorners">if set to <c>true</c> [rounded corners].</param>
        /// <param name="width">The width.</param>
        /// <param name="height">The height.</param>
        /// <returns>WordChart</returns>
        public WordChart AddChart(string title = "", bool roundedCorners = false, int width = 600, int height = 600) {
            var paragraph = this.AddParagraph();
            var chartInstance = new WordChart(this._document, paragraph, title, roundedCorners, width, height);
            return chartInstance;
        }

        /// <summary>
        /// Creates a chart ready for combining bar and line series.
        /// Use <see cref="WordChart.AddChartAxisX"/> to supply category labels
        /// and then call <see cref="WordChart.AddBar(string,int,OfficeIMO.Drawing.OfficeColor)"/> or
        /// <see cref="WordChart.AddLine"/> to add data. <c>AddChartAxisX</c> must be called before adding any
        /// series so that both chart types share the same axes.
        /// </summary>
        public WordChart AddComboChart(string title = "", bool roundedCorners = false, int width = 600, int height = 600) {
            return AddChart(title, roundedCorners, width, height);
        }

        /// <summary>
        /// Inserts a SmartArt diagram at the current position.
        /// </summary>
        /// <param name="type">Layout of SmartArt to create.</param>
        /// <returns>The created <see cref="WordSmartArt"/>.</returns>
        public WordSmartArt AddSmartArt(SmartArtType type) {
            var paragraph = this.AddParagraph();
            return new WordSmartArt(this._document, paragraph, type);
        }

        /// <summary>
        /// Add a foot note for the current paragraph.
        /// </summary>
        /// <param name="text">The text of the note.</param>
        /// <returns>The footnote.</returns>
        public WordParagraph AddFootNote(string text) {
            var footerWordParagraph = new WordParagraph(this._document, true, true) {
                Text = text
            };

            var wordFootNote = WordFootNote.AddFootNote(this._document, this, footerWordParagraph);
            return wordFootNote;
        }

        /// <summary>
        /// Add an end note to the document for this paragraph.
        /// </summary>
        /// <param name="text">The text of the end note.</param>
        /// <returns>The end note.</returns>
        public WordParagraph AddEndNote(string text) {
            var endNoteWordParagraph = new WordParagraph(this._document, true, true);
            endNoteWordParagraph.Text = text;

            var wordEndNote = WordEndNote.AddEndNote(this._document, this, endNoteWordParagraph);
            return wordEndNote;

        }

        /// <summary>
        /// Add a text box to the document.
        /// </summary>
        /// <param name="text">The text inside the text box.</param>
        /// <param name="wrapTextImage">The text image wrapping settings.</param>
        public WordTextBox AddTextBox(string text, WrapTextImage wrapTextImage) {
            WordTextBox wordTextBox = new WordTextBox(this._document, this, text, wrapTextImage);
            return wordTextBox;
        }

        /// <summary>
        /// Adds a legacy VML text box to the paragraph.
        /// </summary>
        public WordTextBox AddTextBoxVml(string text) {
            var run = this.VerifyRun();
            var shape = new V.Shape() {
                Id = "TextBox" + Guid.NewGuid().ToString("N"),
                Style = "mso-wrap-style:square"
            };

            var textbox = new V.TextBox();
            var content = new TextBoxContent(new Paragraph(new Run(new Text(text))));
            textbox.Append(content);
            shape.Append(textbox);

            Picture pict = new Picture();
            pict.Append(shape);
            run.Append(pict);

            return new WordTextBox(this._document, this._paragraph, run);
        }
    }
}