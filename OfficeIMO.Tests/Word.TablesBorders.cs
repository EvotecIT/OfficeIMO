using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;
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


                wordTable.Rows[1].Cells[1].Borders.LeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.LeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.LeftColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.LeftSize = 24;
                wordTable.Rows[1].Cells[1].Borders.LeftSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftStyle == BorderValues.Dotted);
                Assert.Equal(Color.Gold, wordTable.Rows[1].Cells[1].Borders.LeftColor);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.LeftSize?.Value);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.LeftSpace?.Value);






                wordTable.Rows[1].Cells[1].Borders.RightStyle = BorderValues.Double;
                wordTable.Rows[1].Cells[1].Borders.RightColorHex = SixLabors.ImageSharp.Color.Blue.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColorHex == SixLabors.ImageSharp.Color.Blue.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.RightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.RightSize = 4;
                wordTable.Rows[1].Cells[1].Borders.RightSpace = 5U;

                Assert.Equal(BorderValues.Double, wordTable.Rows[1].Cells[1].Borders.RightStyle);
                Assert.Equal(Color.Gold, wordTable.Rows[1].Cells[1].Borders.RightColor);
                Assert.Equal(4U, wordTable.Rows[1].Cells[1].Borders.RightSize?.Value);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.RightSpace?.Value);




                wordTable.Rows[1].Cells[1].Borders.TopStyle = BorderValues.CirclesRectangles;
                wordTable.Rows[1].Cells[1].Borders.TopColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopSize = 6;
                wordTable.Rows[1].Cells[1].Borders.TopSpace = 5U;

                Assert.Equal(BorderValues.CirclesRectangles, wordTable.Rows[1].Cells[1].Borders.TopStyle);
                Assert.Equal(Color.Gold, wordTable.Rows[1].Cells[1].Borders.TopColor);
                Assert.Equal(6U, wordTable.Rows[1].Cells[1].Borders.TopSize?.Value);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.TopSpace?.Value);



                wordTable.Rows[1].Cells[1].Borders.BottomStyle = BorderValues.Safari;
                wordTable.Rows[1].Cells[1].Borders.BottomColorHex = SixLabors.ImageSharp.Color.Azure.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColorHex == SixLabors.ImageSharp.Color.Azure.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.BottomColor = Color.Cyan;
                wordTable.Rows[1].Cells[1].Borders.BottomSize = 8;
                wordTable.Rows[1].Cells[1].Borders.BottomSpace = 5U;

                Assert.Equal(BorderValues.Safari, wordTable.Rows[1].Cells[1].Borders.BottomStyle);
                Assert.Equal(Color.Cyan, wordTable.Rows[1].Cells[1].Borders.BottomColor);
                Assert.Equal(8U, wordTable.Rows[1].Cells[1].Borders.BottomSize?.Value);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.BottomSpace?.Value);

                wordTable.Rows[1].Cells[1].Borders.StartStyle = BorderValues.DashSmallGap;
                wordTable.Rows[1].Cells[1].Borders.StartColorHex = SixLabors.ImageSharp.Color.Orange.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColorHex == SixLabors.ImageSharp.Color.Orange.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.StartColor = Color.Yellow;
                wordTable.Rows[1].Cells[1].Borders.StartSize = 24;
                wordTable.Rows[1].Cells[1].Borders.StartSpace = 10U;

                Assert.Equal(BorderValues.DashSmallGap, wordTable.Rows[1].Cells[1].Borders.StartStyle);
                Assert.Equal(Color.Yellow, wordTable.Rows[1].Cells[1].Borders.StartColor);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.StartSize?.Value);
                Assert.Equal(10U, wordTable.Rows[1].Cells[1].Borders.StartSpace?.Value);

                wordTable.Rows[1].Cells[1].Borders.EndStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.EndColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.EndColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.EndSize = 24;
                //wordTable.Rows[1].Cells[1].Borders.EndSpace = 5U;

                Assert.Equal(BorderValues.Dotted, wordTable.Rows[1].Cells[1].Borders.EndStyle);
                Assert.Equal(Color.Gold, wordTable.Rows[1].Cells[1].Borders.EndColor);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.EndSize?.Value);
                Assert.Null(wordTable.Rows[1].Cells[1].Borders.EndSpace);


                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace = 5U;

                Assert.Equal(BorderValues.Dotted, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle);
                Assert.Equal(Color.Gold, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor);
                Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize?.Value);
                Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace?.Value);


                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor = Color.Aqua;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize = 16;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace = 1U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor == Color.Aqua);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize!.Value == 16);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace!.Value == 1U);



                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesBorders.docx"))) {

                var wordTable = document.Tables[0];

                wordTable.Rows[1].Cells[1].Borders.LeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.LeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.LeftColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.LeftSize = 24;
                wordTable.Rows[1].Cells[1].Borders.LeftSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSize!.Value == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSpace!.Value == 5U);






                wordTable.Rows[1].Cells[1].Borders.RightStyle = BorderValues.Double;
                wordTable.Rows[1].Cells[1].Borders.RightColorHex = SixLabors.ImageSharp.Color.Blue.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColorHex == SixLabors.ImageSharp.Color.Blue.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.RightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.RightSize = 4;
                wordTable.Rows[1].Cells[1].Borders.RightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightStyle == BorderValues.Double);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSize!.Value == 4);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSpace!.Value == 5U);




                wordTable.Rows[1].Cells[1].Borders.TopStyle = BorderValues.CirclesRectangles;
                wordTable.Rows[1].Cells[1].Borders.TopColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopSize = 6;
                wordTable.Rows[1].Cells[1].Borders.TopSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopStyle == BorderValues.CirclesRectangles);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSize!.Value == 6);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSpace!.Value == 5U);



                wordTable.Rows[1].Cells[1].Borders.BottomStyle = BorderValues.Safari;
                wordTable.Rows[1].Cells[1].Borders.BottomColorHex = SixLabors.ImageSharp.Color.Azure.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColorHex == SixLabors.ImageSharp.Color.Azure.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.BottomColor = Color.Cyan;
                wordTable.Rows[1].Cells[1].Borders.BottomSize = 8;
                wordTable.Rows[1].Cells[1].Borders.BottomSpace = 5U;

                  Assert.Equal(BorderValues.Safari, wordTable.Rows[1].Cells[1].Borders.BottomStyle);
                  Assert.Equal(Color.Cyan, wordTable.Rows[1].Cells[1].Borders.BottomColor);
                  Assert.Equal(8U, wordTable.Rows[1].Cells[1].Borders.BottomSize?.Value);
                  Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.BottomSpace?.Value);

                wordTable.Rows[1].Cells[1].Borders.StartStyle = BorderValues.DashSmallGap;
                wordTable.Rows[1].Cells[1].Borders.StartColorHex = SixLabors.ImageSharp.Color.Orange.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColorHex == SixLabors.ImageSharp.Color.Orange.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.StartColor = Color.Yellow;
                wordTable.Rows[1].Cells[1].Borders.StartSize = 24;
                wordTable.Rows[1].Cells[1].Borders.StartSpace = 10U;

                  Assert.Equal(BorderValues.DashSmallGap, wordTable.Rows[1].Cells[1].Borders.StartStyle);
                  Assert.Equal(Color.Yellow, wordTable.Rows[1].Cells[1].Borders.StartColor);
                  Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.StartSize?.Value);
                  Assert.Equal(10U, wordTable.Rows[1].Cells[1].Borders.StartSpace?.Value);

                wordTable.Rows[1].Cells[1].Borders.EndStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.EndColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.EndColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.EndSize = 24;
                //wordTable.Rows[1].Cells[1].Borders.EndSpace = 5U;

                  Assert.Equal(BorderValues.Dotted, wordTable.Rows[1].Cells[1].Borders.EndStyle);
                  Assert.Equal(Color.Gold, wordTable.Rows[1].Cells[1].Borders.EndColor);
                  Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.EndSize?.Value);
                  Assert.Null(wordTable.Rows[1].Cells[1].Borders.EndSpace);


                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace = 5U;

                  Assert.Equal(BorderValues.Dotted, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle);
                  Assert.Equal(Color.Gold, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor);
                  Assert.Equal(24U, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize?.Value);
                  Assert.Equal(5U, wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace?.Value);


                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor = Color.Aqua;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize = 16;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace = 1U;

                  Assert.Equal(BorderValues.Dotted, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle);
                  Assert.Equal(Color.Aqua, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor);
                  Assert.Equal(16U, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize?.Value);
                  Assert.Equal(1U, wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace?.Value);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesBorders.docx"))) {

                var wordTable = document.Tables[0];

                wordTable.Rows[1].Cells[1].Borders.LeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.LeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.LeftColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.LeftSize = 24;
                wordTable.Rows[1].Cells[1].Borders.LeftSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSize!.Value == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSpace!.Value == 5U);






                wordTable.Rows[1].Cells[1].Borders.RightStyle = BorderValues.Double;
                wordTable.Rows[1].Cells[1].Borders.RightColorHex = SixLabors.ImageSharp.Color.Blue.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColorHex == SixLabors.ImageSharp.Color.Blue.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.RightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.RightSize = 4;
                wordTable.Rows[1].Cells[1].Borders.RightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightStyle == BorderValues.Double);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSize!.Value == 4);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSpace!.Value == 5U);




                wordTable.Rows[1].Cells[1].Borders.TopStyle = BorderValues.CirclesRectangles;
                wordTable.Rows[1].Cells[1].Borders.TopColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopSize = 6;
                wordTable.Rows[1].Cells[1].Borders.TopSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopStyle == BorderValues.CirclesRectangles);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSize!.Value == 6);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSpace!.Value == 5U);



                wordTable.Rows[1].Cells[1].Borders.BottomStyle = BorderValues.Safari;
                wordTable.Rows[1].Cells[1].Borders.BottomColorHex = SixLabors.ImageSharp.Color.Azure.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColorHex == SixLabors.ImageSharp.Color.Azure.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.BottomColor = Color.Cyan;
                wordTable.Rows[1].Cells[1].Borders.BottomSize = 8;
                wordTable.Rows[1].Cells[1].Borders.BottomSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomStyle == BorderValues.Safari);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColor == Color.Cyan);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSize!.Value == 8);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSpace!.Value == 5U);

                wordTable.Rows[1].Cells[1].Borders.StartStyle = BorderValues.DashSmallGap;
                wordTable.Rows[1].Cells[1].Borders.StartColorHex = SixLabors.ImageSharp.Color.Orange.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColorHex == SixLabors.ImageSharp.Color.Orange.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.StartColor = Color.Yellow;
                wordTable.Rows[1].Cells[1].Borders.StartSize = 24;
                wordTable.Rows[1].Cells[1].Borders.StartSpace = 10U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartStyle == BorderValues.DashSmallGap);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColor == Color.Yellow);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSize!.Value == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSpace!.Value == 10U);

                wordTable.Rows[1].Cells[1].Borders.EndStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.EndColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.EndColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.EndSize = 24;
                //wordTable.Rows[1].Cells[1].Borders.EndSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndSize!.Value == 24);
                Assert.Null(wordTable.Rows[1].Cells[1].Borders.EndSpace);


                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize!.Value == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace!.Value == 5U);


                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor = Color.Aqua;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize = 16;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace = 1U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor == Color.Aqua);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize!.Value == 16);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace!.Value == 1U);

                wordTable.Rows[1].Cells[1].Borders.InsideVerticalStyle = BorderValues.DecoBlocks;
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalColorHex = SixLabors.ImageSharp.Color.YellowGreen.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideVerticalColorHex == SixLabors.ImageSharp.Color.YellowGreen.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalColor = Color.DarkSlateBlue;
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalSize = 15;
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalSpace = 3U;

                  Assert.Equal(BorderValues.DecoBlocks, wordTable.Rows[1].Cells[1].Borders.InsideVerticalStyle);
                  Assert.Equal(Color.DarkSlateBlue, wordTable.Rows[1].Cells[1].Borders.InsideVerticalColor);
                  Assert.Equal(15U, wordTable.Rows[1].Cells[1].Borders.InsideVerticalSize?.Value);
                  Assert.Equal(3U, wordTable.Rows[1].Cells[1].Borders.InsideVerticalSpace?.Value);

                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalStyle = BorderValues.DecoBlocks;
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColorHex = SixLabors.ImageSharp.Color.YellowGreen.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColorHex == SixLabors.ImageSharp.Color.YellowGreen.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColor = Color.DarkSlateBlue;
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSize = 15;
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSpace = 3U;

                  Assert.Equal(BorderValues.DecoBlocks, wordTable.Rows[1].Cells[1].Borders.InsideHorizontalStyle);
                  Assert.Equal(Color.DarkSlateBlue, wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColor);
                  Assert.Equal(15U, wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSize?.Value);
                  Assert.Equal(3U, wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSpace?.Value);

                document.Save();
            }
        }


        [Fact]
        public void Test_LoadingWordDocumentWithTablesAndBorders() {
            string filePath = Path.Combine(_directoryDocuments, "DocumentWithTables.docx");
            using (WordDocument document = WordDocument.Load(filePath)) {
                var wordTable = document.Tables[0];

                wordTable.Rows[1].Cells[1].Borders.LeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.LeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.LeftColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.LeftSize = 24;
                wordTable.Rows[1].Cells[1].Borders.LeftSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSize!.Value == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSpace!.Value == 5U);






                wordTable.Rows[1].Cells[1].Borders.RightStyle = BorderValues.Double;
                wordTable.Rows[1].Cells[1].Borders.RightColorHex = SixLabors.ImageSharp.Color.Blue.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColorHex == SixLabors.ImageSharp.Color.Blue.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.RightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.RightSize = 4;
                wordTable.Rows[1].Cells[1].Borders.RightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightStyle == BorderValues.Double);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSize!.Value == 4);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSpace!.Value == 5U);




                wordTable.Rows[1].Cells[1].Borders.TopStyle = BorderValues.CirclesRectangles;
                wordTable.Rows[1].Cells[1].Borders.TopColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopSize = 6;
                wordTable.Rows[1].Cells[1].Borders.TopSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopStyle == BorderValues.CirclesRectangles);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSize!.Value == 6);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSpace!.Value == 5U);



                wordTable.Rows[1].Cells[1].Borders.BottomStyle = BorderValues.Safari;
                wordTable.Rows[1].Cells[1].Borders.BottomColorHex = SixLabors.ImageSharp.Color.Azure.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColorHex == SixLabors.ImageSharp.Color.Azure.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.BottomColor = Color.Cyan;
                wordTable.Rows[1].Cells[1].Borders.BottomSize = 8;
                wordTable.Rows[1].Cells[1].Borders.BottomSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomStyle == BorderValues.Safari);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColor == Color.Cyan);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSize!.Value == 8);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSpace!.Value == 5U);

                wordTable.Rows[1].Cells[1].Borders.StartStyle = BorderValues.DashSmallGap;
                wordTable.Rows[1].Cells[1].Borders.StartColorHex = SixLabors.ImageSharp.Color.Orange.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColorHex == SixLabors.ImageSharp.Color.Orange.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.StartColor = Color.Yellow;
                wordTable.Rows[1].Cells[1].Borders.StartSize = 24;
                wordTable.Rows[1].Cells[1].Borders.StartSpace = 10U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartStyle == BorderValues.DashSmallGap);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColor == Color.Yellow);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSize!.Value == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSpace!.Value == 10U);

                wordTable.Rows[1].Cells[1].Borders.EndStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.EndColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.EndColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.EndSize = 24;
                //wordTable.Rows[1].Cells[1].Borders.EndSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndSize!.Value == 24);
                Assert.Null(wordTable.Rows[1].Cells[1].Borders.EndSpace);


                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize!.Value == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace!.Value == 5U);


                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor = Color.Aqua;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize = 16;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace = 1U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor == Color.Aqua);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize!.Value == 16);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace!.Value == 1U);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "DocumentWithTables.docx"))) {
                var wordTable = document.Tables[0];

                wordTable.Rows[1].Cells[1].Borders.LeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.LeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.LeftColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.LeftSize = 24;
                wordTable.Rows[1].Cells[1].Borders.LeftSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSize!.Value == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSpace!.Value == 5U);






                wordTable.Rows[1].Cells[1].Borders.RightStyle = BorderValues.Double;
                wordTable.Rows[1].Cells[1].Borders.RightColorHex = SixLabors.ImageSharp.Color.Blue.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColorHex == SixLabors.ImageSharp.Color.Blue.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.RightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.RightSize = 4;
                wordTable.Rows[1].Cells[1].Borders.RightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightStyle == BorderValues.Double);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSize!.Value == 4);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSpace!.Value == 5U);




                wordTable.Rows[1].Cells[1].Borders.TopStyle = BorderValues.CirclesRectangles;
                wordTable.Rows[1].Cells[1].Borders.TopColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopSize = 6;
                wordTable.Rows[1].Cells[1].Borders.TopSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopStyle == BorderValues.CirclesRectangles);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSize!.Value == 6);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSpace!.Value == 5U);



                wordTable.Rows[1].Cells[1].Borders.BottomStyle = BorderValues.Safari;
                wordTable.Rows[1].Cells[1].Borders.BottomColorHex = SixLabors.ImageSharp.Color.Azure.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColorHex == SixLabors.ImageSharp.Color.Azure.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.BottomColor = Color.Cyan;
                wordTable.Rows[1].Cells[1].Borders.BottomSize = 8;
                wordTable.Rows[1].Cells[1].Borders.BottomSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomStyle == BorderValues.Safari);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColor == Color.Cyan);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSize!.Value == 8);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSpace!.Value == 5U);

                wordTable.Rows[1].Cells[1].Borders.StartStyle = BorderValues.DashSmallGap;
                wordTable.Rows[1].Cells[1].Borders.StartColorHex = SixLabors.ImageSharp.Color.Orange.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColorHex == SixLabors.ImageSharp.Color.Orange.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.StartColor = Color.Yellow;
                wordTable.Rows[1].Cells[1].Borders.StartSize = 24;
                wordTable.Rows[1].Cells[1].Borders.StartSpace = 10U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartStyle == BorderValues.DashSmallGap);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColor == Color.Yellow);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSize!.Value == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSpace!.Value == 10U);

                wordTable.Rows[1].Cells[1].Borders.EndStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.EndColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.EndColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.EndSize = 24;
                //wordTable.Rows[1].Cells[1].Borders.EndSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndSize!.Value == 24);
                Assert.Null(wordTable.Rows[1].Cells[1].Borders.EndSpace);


                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize!.Value == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace!.Value == 5U);


                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor = Color.Aqua;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize = 16;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace = 1U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor == Color.Aqua);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize!.Value == 16);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace!.Value == 1U);


                wordTable.Rows[1].Cells[1].Borders.InsideVerticalStyle = BorderValues.DecoBlocks;
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalColorHex = SixLabors.ImageSharp.Color.YellowGreen.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideVerticalColorHex == SixLabors.ImageSharp.Color.YellowGreen.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalColor = Color.DarkSlateBlue;
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalSize = 15;
                wordTable.Rows[1].Cells[1].Borders.InsideVerticalSpace = 3U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideVerticalStyle == BorderValues.DecoBlocks);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideVerticalColor == Color.DarkSlateBlue);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideVerticalSize!.Value == 15);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideVerticalSpace!.Value == 3U);

                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalStyle = BorderValues.DecoBlocks;
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColorHex = SixLabors.ImageSharp.Color.YellowGreen.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColorHex == SixLabors.ImageSharp.Color.YellowGreen.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColor = Color.DarkSlateBlue;
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSize = 15;
                wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSpace = 3U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideHorizontalStyle == BorderValues.DecoBlocks);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideHorizontalColor == Color.DarkSlateBlue);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSize!.Value == 15);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.InsideHorizontalSpace!.Value == 3U);

                document.Save();
            }
        }
    }
}
