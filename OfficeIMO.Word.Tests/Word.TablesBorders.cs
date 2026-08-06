using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = OfficeIMO.Drawing.OfficeColor;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithTablesAndBorders() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesBorders.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                WordTable wordTable = document.AddTable(4, 4, WordTableStyle.TableNormal);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";


                wordTable.Rows[1].Cells[1].Borders.LeftStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.LeftColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.LeftColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.LeftSize = 24;
                wordTable.Rows[1].Cells[1].Borders.LeftSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftStyle == WordBorderStyle.Dotted);
                Assert.Equal(Color.Gold, wordTable.Rows[1].Cells[1].Borders.LeftColor);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.LeftSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.LeftSpace);






                wordTable.Rows[1].Cells[1].Borders.RightStyle = WordBorderStyle.Double;
                wordTable.Rows[1].Cells[1].Borders.RightColorHex = OfficeIMO.Drawing.OfficeColor.Blue.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColorHex == OfficeIMO.Drawing.OfficeColor.Blue.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.RightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.RightSize = 4;
                wordTable.Rows[1].Cells[1].Borders.RightSpace = 5U;

                Assert.Equal(WordBorderStyle.Double, wordTable.Rows[1].Cells[1].Borders.RightStyle);
                Assert.Equal(Color.Gold, wordTable.Rows[1].Cells[1].Borders.RightColor);
                Assert.Equal(4U, wordTable.Rows[1].Cells[1].Borders.RightSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.RightSpace);




                wordTable.Rows[1].Cells[1].Borders.TopStyle = WordBorderStyle.CirclesRectangles;
                wordTable.Rows[1].Cells[1].Borders.TopColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.TopColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopSize = 6;
                wordTable.Rows[1].Cells[1].Borders.TopSpace = 5U;

                Assert.Equal(WordBorderStyle.CirclesRectangles, wordTable.Rows[1].Cells[1].Borders.TopStyle);
                Assert.Equal(Color.Gold, wordTable.Rows[1].Cells[1].Borders.TopColor);
                Assert.Equal(6U, wordTable.Rows[1].Cells[1].Borders.TopSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.TopSpace);



                wordTable.Rows[1].Cells[1].Borders.BottomStyle = WordBorderStyle.Safari;
                wordTable.Rows[1].Cells[1].Borders.BottomColorHex = OfficeIMO.Drawing.OfficeColor.Azure.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColorHex == OfficeIMO.Drawing.OfficeColor.Azure.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.BottomColor = Color.Cyan;
                wordTable.Rows[1].Cells[1].Borders.BottomSize = 8;
                wordTable.Rows[1].Cells[1].Borders.BottomSpace = 5U;

                Assert.Equal(WordBorderStyle.Safari, wordTable.Rows[1].Cells[1].Borders.BottomStyle);
                Assert.Equal(Color.Cyan, wordTable.Rows[1].Cells[1].Borders.BottomColor);
                Assert.Equal(8U, wordTable.Rows[1].Cells[1].Borders.BottomSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.BottomSpace);

                wordTable.Rows[1].Cells[1].Borders.StartStyle = WordBorderStyle.DashSmallGap;
                wordTable.Rows[1].Cells[1].Borders.StartColorHex = OfficeIMO.Drawing.OfficeColor.Orange.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColorHex == OfficeIMO.Drawing.OfficeColor.Orange.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.StartColor = Color.Yellow;
                wordTable.Rows[1].Cells[1].Borders.StartSize = 24;
                wordTable.Rows[1].Cells[1].Borders.StartSpace = 10U;

                Assert.Equal(WordBorderStyle.DashSmallGap, wordTable.Rows[1].Cells[1].Borders.StartStyle);
                Assert.Equal(Color.Yellow, wordTable.Rows[1].Cells[1].Borders.StartColor);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.StartSize);
                Assert.Equal(10U, wordTable.Rows[1].Cells[1].Borders.StartSpace);

                wordTable.Rows[1].Cells[1].Borders.EndStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.EndColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.EndColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.EndSize = 24;
                //wordTable.Rows[1].Cells[1].Borders.EndSpace = 5U;

                Assert.Equal(WordBorderStyle.Dotted, wordTable.Rows[1].Cells[1].Borders.EndStyle);
                Assert.Equal(Color.Gold, wordTable.Rows[1].Cells[1].Borders.EndColor);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.EndSize);
                Assert.Null(wordTable.Rows[1].Cells[1].Borders.EndSpace);


                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace = 5U;

                Assert.Equal(WordBorderStyle.Dotted, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle);
                Assert.Equal(Color.Gold, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace);


                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor = Color.Aqua;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize = 16;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace = 1U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle == WordBorderStyle.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor == Color.Aqua);
                Assert.Equal(16U, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize);
                Assert.Equal(1U, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace);



                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesBorders.docx"))) {

                var wordTable = document.Tables[0];

                wordTable.Rows[1].Cells[1].Borders.LeftStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.LeftColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.LeftColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.LeftSize = 24;
                wordTable.Rows[1].Cells[1].Borders.LeftSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftStyle == WordBorderStyle.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColor == Color.Gold);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.LeftSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.LeftSpace);






                wordTable.Rows[1].Cells[1].Borders.RightStyle = WordBorderStyle.Double;
                wordTable.Rows[1].Cells[1].Borders.RightColorHex = OfficeIMO.Drawing.OfficeColor.Blue.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColorHex == OfficeIMO.Drawing.OfficeColor.Blue.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.RightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.RightSize = 4;
                wordTable.Rows[1].Cells[1].Borders.RightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightStyle == WordBorderStyle.Double);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColor == Color.Gold);
                Assert.Equal(4U, wordTable.Rows[1].Cells[1].Borders.RightSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.RightSpace);




                wordTable.Rows[1].Cells[1].Borders.TopStyle = WordBorderStyle.CirclesRectangles;
                wordTable.Rows[1].Cells[1].Borders.TopColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.TopColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopSize = 6;
                wordTable.Rows[1].Cells[1].Borders.TopSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopStyle == WordBorderStyle.CirclesRectangles);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColor == Color.Gold);
                Assert.Equal(6U, wordTable.Rows[1].Cells[1].Borders.TopSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.TopSpace);



                wordTable.Rows[1].Cells[1].Borders.BottomStyle = WordBorderStyle.Safari;
                wordTable.Rows[1].Cells[1].Borders.BottomColorHex = OfficeIMO.Drawing.OfficeColor.Azure.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColorHex == OfficeIMO.Drawing.OfficeColor.Azure.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.BottomColor = Color.Cyan;
                wordTable.Rows[1].Cells[1].Borders.BottomSize = 8;
                wordTable.Rows[1].Cells[1].Borders.BottomSpace = 5U;

                  Assert.Equal(WordBorderStyle.Safari, wordTable.Rows[1].Cells[1].Borders.BottomStyle);
                  Assert.Equal(Color.Cyan, wordTable.Rows[1].Cells[1].Borders.BottomColor);
                  Assert.Equal(8U, wordTable.Rows[1].Cells[1].Borders.BottomSize);
                  Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.BottomSpace);

                wordTable.Rows[1].Cells[1].Borders.StartStyle = WordBorderStyle.DashSmallGap;
                wordTable.Rows[1].Cells[1].Borders.StartColorHex = OfficeIMO.Drawing.OfficeColor.Orange.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColorHex == OfficeIMO.Drawing.OfficeColor.Orange.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.StartColor = Color.Yellow;
                wordTable.Rows[1].Cells[1].Borders.StartSize = 24;
                wordTable.Rows[1].Cells[1].Borders.StartSpace = 10U;

                  Assert.Equal(WordBorderStyle.DashSmallGap, wordTable.Rows[1].Cells[1].Borders.StartStyle);
                  Assert.Equal(Color.Yellow, wordTable.Rows[1].Cells[1].Borders.StartColor);
                  Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.StartSize);
                  Assert.Equal(10U, wordTable.Rows[1].Cells[1].Borders.StartSpace);

                wordTable.Rows[1].Cells[1].Borders.EndStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.EndColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.EndColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.EndSize = 24;
                //wordTable.Rows[1].Cells[1].Borders.EndSpace = 5U;

                  Assert.Equal(WordBorderStyle.Dotted, wordTable.Rows[1].Cells[1].Borders.EndStyle);
                  Assert.Equal(Color.Gold, wordTable.Rows[1].Cells[1].Borders.EndColor);
                  Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.EndSize);
                  Assert.Null(wordTable.Rows[1].Cells[1].Borders.EndSpace);


                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace = 5U;

                  Assert.Equal(WordBorderStyle.Dotted, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle);
                  Assert.Equal(Color.Gold, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor);
                  Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize);
                  Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace);


                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor = Color.Aqua;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize = 16;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace = 1U;

                  Assert.Equal(WordBorderStyle.Dotted, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle);
                  Assert.Equal(Color.Aqua, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor);
                  Assert.Equal(16U, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize);
                  Assert.Equal(1U, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesBorders.docx"))) {

                var wordTable = document.Tables[0];

                wordTable.Rows[1].Cells[1].Borders.LeftStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.LeftColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.LeftColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.LeftSize = 24;
                wordTable.Rows[1].Cells[1].Borders.LeftSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftStyle == WordBorderStyle.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColor == Color.Gold);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.LeftSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.LeftSpace);






                wordTable.Rows[1].Cells[1].Borders.RightStyle = WordBorderStyle.Double;
                wordTable.Rows[1].Cells[1].Borders.RightColorHex = OfficeIMO.Drawing.OfficeColor.Blue.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColorHex == OfficeIMO.Drawing.OfficeColor.Blue.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.RightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.RightSize = 4;
                wordTable.Rows[1].Cells[1].Borders.RightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightStyle == WordBorderStyle.Double);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColor == Color.Gold);
                Assert.Equal(4U, wordTable.Rows[1].Cells[1].Borders.RightSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.RightSpace);




                wordTable.Rows[1].Cells[1].Borders.TopStyle = WordBorderStyle.CirclesRectangles;
                wordTable.Rows[1].Cells[1].Borders.TopColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.TopColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopSize = 6;
                wordTable.Rows[1].Cells[1].Borders.TopSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopStyle == WordBorderStyle.CirclesRectangles);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColor == Color.Gold);
                Assert.Equal(6U, wordTable.Rows[1].Cells[1].Borders.TopSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.TopSpace);



                wordTable.Rows[1].Cells[1].Borders.BottomStyle = WordBorderStyle.Safari;
                wordTable.Rows[1].Cells[1].Borders.BottomColorHex = OfficeIMO.Drawing.OfficeColor.Azure.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColorHex == OfficeIMO.Drawing.OfficeColor.Azure.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.BottomColor = Color.Cyan;
                wordTable.Rows[1].Cells[1].Borders.BottomSize = 8;
                wordTable.Rows[1].Cells[1].Borders.BottomSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomStyle == WordBorderStyle.Safari);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColor == Color.Cyan);
                Assert.Equal(8U, wordTable.Rows[1].Cells[1].Borders.BottomSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.BottomSpace);

                wordTable.Rows[1].Cells[1].Borders.StartStyle = WordBorderStyle.DashSmallGap;
                wordTable.Rows[1].Cells[1].Borders.StartColorHex = OfficeIMO.Drawing.OfficeColor.Orange.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColorHex == OfficeIMO.Drawing.OfficeColor.Orange.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.StartColor = Color.Yellow;
                wordTable.Rows[1].Cells[1].Borders.StartSize = 24;
                wordTable.Rows[1].Cells[1].Borders.StartSpace = 10U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartStyle == WordBorderStyle.DashSmallGap);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColor == Color.Yellow);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.StartSize);
                Assert.Equal(10U, wordTable.Rows[1].Cells[1].Borders.StartSpace);

                wordTable.Rows[1].Cells[1].Borders.EndStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.EndColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.EndColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.EndSize = 24;
                //wordTable.Rows[1].Cells[1].Borders.EndSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndStyle == WordBorderStyle.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColor == Color.Gold);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.EndSize);
                Assert.Null(wordTable.Rows[1].Cells[1].Borders.EndSpace);


                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle == WordBorderStyle.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor == Color.Gold);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace);


                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor = Color.Aqua;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize = 16;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace = 1U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle == WordBorderStyle.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor == Color.Aqua);
                Assert.Equal(16U, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize);
                Assert.Equal(1U, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace);

                wordTable.Rows[1].Cells[1].Borders.InsideVerticalStyle = WordBorderStyle.DecoBlocks;
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalColorHex = OfficeIMO.Drawing.OfficeColor.YellowGreen.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideVerticalColorHex == OfficeIMO.Drawing.OfficeColor.YellowGreen.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalColor = Color.DarkSlateBlue;
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalSize = 15;
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalSpace = 3U;

                  Assert.Equal(WordBorderStyle.DecoBlocks, wordTable.Rows[1].Cells[1].Borders.InsideVerticalStyle);
                  Assert.Equal(Color.DarkSlateBlue, wordTable.Rows[1].Cells[1].Borders.InsideVerticalColor);
                  Assert.Equal(15U, wordTable.Rows[1].Cells[1].Borders.InsideVerticalSize);
                  Assert.Equal(3U, wordTable.Rows[1].Cells[1].Borders.InsideVerticalSpace);

                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalStyle = WordBorderStyle.DecoBlocks;
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColorHex = OfficeIMO.Drawing.OfficeColor.YellowGreen.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColorHex == OfficeIMO.Drawing.OfficeColor.YellowGreen.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColor = Color.DarkSlateBlue;
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSize = 15;
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSpace = 3U;

                  Assert.Equal(WordBorderStyle.DecoBlocks, wordTable.Rows[1].Cells[1].Borders.InsideHorizontalStyle);
                  Assert.Equal(Color.DarkSlateBlue, wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColor);
                  Assert.Equal(15U, wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSize);
                  Assert.Equal(3U, wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSpace);

                document.Save();
            }
        }


        [Fact]
        public void Test_LoadingWordDocumentWithTablesAndBorders() {
            string filePath = Path.Combine(_directoryDocuments, "DocumentWithTables.docx");
            using (WordDocument document = WordDocument.Load(filePath)) {
                var wordTable = document.Tables[0];

                wordTable.Rows[1].Cells[1].Borders.LeftStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.LeftColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.LeftColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.LeftSize = 24;
                wordTable.Rows[1].Cells[1].Borders.LeftSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftStyle == WordBorderStyle.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColor == Color.Gold);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.LeftSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.LeftSpace);






                wordTable.Rows[1].Cells[1].Borders.RightStyle = WordBorderStyle.Double;
                wordTable.Rows[1].Cells[1].Borders.RightColorHex = OfficeIMO.Drawing.OfficeColor.Blue.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColorHex == OfficeIMO.Drawing.OfficeColor.Blue.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.RightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.RightSize = 4;
                wordTable.Rows[1].Cells[1].Borders.RightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightStyle == WordBorderStyle.Double);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColor == Color.Gold);
                Assert.Equal(4U, wordTable.Rows[1].Cells[1].Borders.RightSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.RightSpace);




                wordTable.Rows[1].Cells[1].Borders.TopStyle = WordBorderStyle.CirclesRectangles;
                wordTable.Rows[1].Cells[1].Borders.TopColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.TopColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopSize = 6;
                wordTable.Rows[1].Cells[1].Borders.TopSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopStyle == WordBorderStyle.CirclesRectangles);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColor == Color.Gold);
                Assert.Equal(6U, wordTable.Rows[1].Cells[1].Borders.TopSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.TopSpace);



                wordTable.Rows[1].Cells[1].Borders.BottomStyle = WordBorderStyle.Safari;
                wordTable.Rows[1].Cells[1].Borders.BottomColorHex = OfficeIMO.Drawing.OfficeColor.Azure.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColorHex == OfficeIMO.Drawing.OfficeColor.Azure.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.BottomColor = Color.Cyan;
                wordTable.Rows[1].Cells[1].Borders.BottomSize = 8;
                wordTable.Rows[1].Cells[1].Borders.BottomSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomStyle == WordBorderStyle.Safari);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColor == Color.Cyan);
                Assert.Equal(8U, wordTable.Rows[1].Cells[1].Borders.BottomSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.BottomSpace);

                wordTable.Rows[1].Cells[1].Borders.StartStyle = WordBorderStyle.DashSmallGap;
                wordTable.Rows[1].Cells[1].Borders.StartColorHex = OfficeIMO.Drawing.OfficeColor.Orange.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColorHex == OfficeIMO.Drawing.OfficeColor.Orange.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.StartColor = Color.Yellow;
                wordTable.Rows[1].Cells[1].Borders.StartSize = 24;
                wordTable.Rows[1].Cells[1].Borders.StartSpace = 10U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartStyle == WordBorderStyle.DashSmallGap);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColor == Color.Yellow);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.StartSize);
                Assert.Equal(10U, wordTable.Rows[1].Cells[1].Borders.StartSpace);

                wordTable.Rows[1].Cells[1].Borders.EndStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.EndColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.EndColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.EndSize = 24;
                //wordTable.Rows[1].Cells[1].Borders.EndSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndStyle == WordBorderStyle.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColor == Color.Gold);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.EndSize);
                Assert.Null(wordTable.Rows[1].Cells[1].Borders.EndSpace);


                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle == WordBorderStyle.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor == Color.Gold);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace);


                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor = Color.Aqua;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize = 16;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace = 1U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle == WordBorderStyle.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor == Color.Aqua);
                Assert.Equal(16U, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize);
                Assert.Equal(1U, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "DocumentWithTables.docx"))) {
                var wordTable = document.Tables[0];

                wordTable.Rows[1].Cells[1].Borders.LeftStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.LeftColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.LeftColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.LeftSize = 24;
                wordTable.Rows[1].Cells[1].Borders.LeftSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftStyle == WordBorderStyle.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColor == Color.Gold);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.LeftSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.LeftSpace);






                wordTable.Rows[1].Cells[1].Borders.RightStyle = WordBorderStyle.Double;
                wordTable.Rows[1].Cells[1].Borders.RightColorHex = OfficeIMO.Drawing.OfficeColor.Blue.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColorHex == OfficeIMO.Drawing.OfficeColor.Blue.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.RightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.RightSize = 4;
                wordTable.Rows[1].Cells[1].Borders.RightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightStyle == WordBorderStyle.Double);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColor == Color.Gold);
                Assert.Equal(4U, wordTable.Rows[1].Cells[1].Borders.RightSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.RightSpace);




                wordTable.Rows[1].Cells[1].Borders.TopStyle = WordBorderStyle.CirclesRectangles;
                wordTable.Rows[1].Cells[1].Borders.TopColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.TopColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopSize = 6;
                wordTable.Rows[1].Cells[1].Borders.TopSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopStyle == WordBorderStyle.CirclesRectangles);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColor == Color.Gold);
                Assert.Equal(6U, wordTable.Rows[1].Cells[1].Borders.TopSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.TopSpace);



                wordTable.Rows[1].Cells[1].Borders.BottomStyle = WordBorderStyle.Safari;
                wordTable.Rows[1].Cells[1].Borders.BottomColorHex = OfficeIMO.Drawing.OfficeColor.Azure.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColorHex == OfficeIMO.Drawing.OfficeColor.Azure.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.BottomColor = Color.Cyan;
                wordTable.Rows[1].Cells[1].Borders.BottomSize = 8;
                wordTable.Rows[1].Cells[1].Borders.BottomSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomStyle == WordBorderStyle.Safari);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColor == Color.Cyan);
                Assert.Equal(8U, wordTable.Rows[1].Cells[1].Borders.BottomSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.BottomSpace);

                wordTable.Rows[1].Cells[1].Borders.StartStyle = WordBorderStyle.DashSmallGap;
                wordTable.Rows[1].Cells[1].Borders.StartColorHex = OfficeIMO.Drawing.OfficeColor.Orange.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColorHex == OfficeIMO.Drawing.OfficeColor.Orange.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.StartColor = Color.Yellow;
                wordTable.Rows[1].Cells[1].Borders.StartSize = 24;
                wordTable.Rows[1].Cells[1].Borders.StartSpace = 10U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartStyle == WordBorderStyle.DashSmallGap);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColor == Color.Yellow);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.StartSize);
                Assert.Equal(10U, wordTable.Rows[1].Cells[1].Borders.StartSpace);

                wordTable.Rows[1].Cells[1].Borders.EndStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.EndColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.EndColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.EndSize = 24;
                //wordTable.Rows[1].Cells[1].Borders.EndSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndStyle == WordBorderStyle.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColor == Color.Gold);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.EndSize);
                Assert.Null(wordTable.Rows[1].Cells[1].Borders.EndSpace);


                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle == WordBorderStyle.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor == Color.Gold);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace);


                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle = WordBorderStyle.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex = OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor = Color.Aqua;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize = 16;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace = 1U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle == WordBorderStyle.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor == Color.Aqua);
                Assert.Equal(16U, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize);
                Assert.Equal(1U, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace);


                wordTable.Rows[1].Cells[1].Borders.InsideVerticalStyle = WordBorderStyle.DecoBlocks;
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalColorHex = OfficeIMO.Drawing.OfficeColor.YellowGreen.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideVerticalColorHex == OfficeIMO.Drawing.OfficeColor.YellowGreen.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalColor = Color.DarkSlateBlue;
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalSize = 15;
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalSpace = 3U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideVerticalStyle == WordBorderStyle.DecoBlocks);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideVerticalColor == Color.DarkSlateBlue);
                Assert.Equal(15U, wordTable.Rows[1].Cells[1].Borders.InsideVerticalSize);
                Assert.Equal(3U, wordTable.Rows[1].Cells[1].Borders.InsideVerticalSpace);

                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalStyle = WordBorderStyle.DecoBlocks;
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColorHex = OfficeIMO.Drawing.OfficeColor.YellowGreen.ToRgbHex();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColorHex == OfficeIMO.Drawing.OfficeColor.YellowGreen.ToRgbHex());
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColor = Color.DarkSlateBlue;
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSize = 15;
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSpace = 3U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideHorizontalStyle == WordBorderStyle.DecoBlocks);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColor == Color.DarkSlateBlue);
                Assert.Equal(15U, wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSize);
                Assert.Equal(3U, wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSpace);

                document.Save();
            }
        }
    }
}
