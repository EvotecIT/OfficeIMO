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
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSpace == 5U);






                wordTable.Rows[1].Cells[1].Borders.RightStyle = BorderValues.Double;
                wordTable.Rows[1].Cells[1].Borders.RightColorHex = SixLabors.ImageSharp.Color.Blue.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColorHex == SixLabors.ImageSharp.Color.Blue.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.RightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.RightSize = 4;
                wordTable.Rows[1].Cells[1].Borders.RightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightStyle == BorderValues.Double);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSize == 4);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSpace == 5U);




                wordTable.Rows[1].Cells[1].Borders.TopStyle = BorderValues.CirclesRectangles;
                wordTable.Rows[1].Cells[1].Borders.TopColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopSize = 6;
                wordTable.Rows[1].Cells[1].Borders.TopSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopStyle == BorderValues.CirclesRectangles);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSize == 6);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSpace == 5U);



                wordTable.Rows[1].Cells[1].Borders.BottomStyle = BorderValues.Safari;
                wordTable.Rows[1].Cells[1].Borders.BottomColorHex = SixLabors.ImageSharp.Color.Azure.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColorHex == SixLabors.ImageSharp.Color.Azure.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.BottomColor = Color.Cyan;
                wordTable.Rows[1].Cells[1].Borders.BottomSize = 8;
                wordTable.Rows[1].Cells[1].Borders.BottomSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomStyle == BorderValues.Safari);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColor == Color.Cyan);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSize == 8);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSpace == 5U);

                wordTable.Rows[1].Cells[1].Borders.StartStyle = BorderValues.DashSmallGap;
                wordTable.Rows[1].Cells[1].Borders.StartColorHex = SixLabors.ImageSharp.Color.Orange.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColorHex == SixLabors.ImageSharp.Color.Orange.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.StartColor = Color.Yellow;
                wordTable.Rows[1].Cells[1].Borders.StartSize = 24;
                wordTable.Rows[1].Cells[1].Borders.StartSpace = 10U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartStyle == BorderValues.DashSmallGap);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColor == Color.Yellow);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSpace == 10U);

                wordTable.Rows[1].Cells[1].Borders.EndStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.EndColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.EndColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.EndSize = 24;
                //wordTable.Rows[1].Cells[1].Borders.EndSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndSpace == null);


                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace == 5U);


                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor = Color.Aqua;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize = 16;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace = 1U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor == Color.Aqua);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize == 16);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace == 1U);



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
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSpace == 5U);






                wordTable.Rows[1].Cells[1].Borders.RightStyle = BorderValues.Double;
                wordTable.Rows[1].Cells[1].Borders.RightColorHex = SixLabors.ImageSharp.Color.Blue.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColorHex == SixLabors.ImageSharp.Color.Blue.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.RightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.RightSize = 4;
                wordTable.Rows[1].Cells[1].Borders.RightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightStyle == BorderValues.Double);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSize == 4);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSpace == 5U);




                wordTable.Rows[1].Cells[1].Borders.TopStyle = BorderValues.CirclesRectangles;
                wordTable.Rows[1].Cells[1].Borders.TopColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopSize = 6;
                wordTable.Rows[1].Cells[1].Borders.TopSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopStyle == BorderValues.CirclesRectangles);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSize == 6);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSpace == 5U);



                wordTable.Rows[1].Cells[1].Borders.BottomStyle = BorderValues.Safari;
                wordTable.Rows[1].Cells[1].Borders.BottomColorHex = SixLabors.ImageSharp.Color.Azure.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColorHex == SixLabors.ImageSharp.Color.Azure.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.BottomColor = Color.Cyan;
                wordTable.Rows[1].Cells[1].Borders.BottomSize = 8;
                wordTable.Rows[1].Cells[1].Borders.BottomSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomStyle == BorderValues.Safari);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColor == Color.Cyan);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSize == 8);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSpace == 5U);

                wordTable.Rows[1].Cells[1].Borders.StartStyle = BorderValues.DashSmallGap;
                wordTable.Rows[1].Cells[1].Borders.StartColorHex = SixLabors.ImageSharp.Color.Orange.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColorHex == SixLabors.ImageSharp.Color.Orange.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.StartColor = Color.Yellow;
                wordTable.Rows[1].Cells[1].Borders.StartSize = 24;
                wordTable.Rows[1].Cells[1].Borders.StartSpace = 10U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartStyle == BorderValues.DashSmallGap);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColor == Color.Yellow);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSpace == 10U);

                wordTable.Rows[1].Cells[1].Borders.EndStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.EndColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.EndColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.EndSize = 24;
                //wordTable.Rows[1].Cells[1].Borders.EndSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndSpace == null);


                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace == 5U);


                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor = Color.Aqua;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize = 16;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace = 1U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor == Color.Aqua);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize == 16);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace == 1U);

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
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSpace == 5U);






                wordTable.Rows[1].Cells[1].Borders.RightStyle = BorderValues.Double;
                wordTable.Rows[1].Cells[1].Borders.RightColorHex = SixLabors.ImageSharp.Color.Blue.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColorHex == SixLabors.ImageSharp.Color.Blue.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.RightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.RightSize = 4;
                wordTable.Rows[1].Cells[1].Borders.RightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightStyle == BorderValues.Double);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSize == 4);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSpace == 5U);




                wordTable.Rows[1].Cells[1].Borders.TopStyle = BorderValues.CirclesRectangles;
                wordTable.Rows[1].Cells[1].Borders.TopColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopSize = 6;
                wordTable.Rows[1].Cells[1].Borders.TopSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopStyle == BorderValues.CirclesRectangles);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSize == 6);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSpace == 5U);



                wordTable.Rows[1].Cells[1].Borders.BottomStyle = BorderValues.Safari;
                wordTable.Rows[1].Cells[1].Borders.BottomColorHex = SixLabors.ImageSharp.Color.Azure.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColorHex == SixLabors.ImageSharp.Color.Azure.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.BottomColor = Color.Cyan;
                wordTable.Rows[1].Cells[1].Borders.BottomSize = 8;
                wordTable.Rows[1].Cells[1].Borders.BottomSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomStyle == BorderValues.Safari);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColor == Color.Cyan);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSize == 8);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSpace == 5U);

                wordTable.Rows[1].Cells[1].Borders.StartStyle = BorderValues.DashSmallGap;
                wordTable.Rows[1].Cells[1].Borders.StartColorHex = SixLabors.ImageSharp.Color.Orange.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColorHex == SixLabors.ImageSharp.Color.Orange.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.StartColor = Color.Yellow;
                wordTable.Rows[1].Cells[1].Borders.StartSize = 24;
                wordTable.Rows[1].Cells[1].Borders.StartSpace = 10U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartStyle == BorderValues.DashSmallGap);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColor == Color.Yellow);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSpace == 10U);

                wordTable.Rows[1].Cells[1].Borders.EndStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.EndColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.EndColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.EndSize = 24;
                //wordTable.Rows[1].Cells[1].Borders.EndSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndSpace == null);


                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace == 5U);


                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor = Color.Aqua;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize = 16;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace = 1U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor == Color.Aqua);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize == 16);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace == 1U);

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
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSpace == 5U);






                wordTable.Rows[1].Cells[1].Borders.RightStyle = BorderValues.Double;
                wordTable.Rows[1].Cells[1].Borders.RightColorHex = SixLabors.ImageSharp.Color.Blue.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColorHex == SixLabors.ImageSharp.Color.Blue.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.RightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.RightSize = 4;
                wordTable.Rows[1].Cells[1].Borders.RightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightStyle == BorderValues.Double);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSize == 4);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSpace == 5U);




                wordTable.Rows[1].Cells[1].Borders.TopStyle = BorderValues.CirclesRectangles;
                wordTable.Rows[1].Cells[1].Borders.TopColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopSize = 6;
                wordTable.Rows[1].Cells[1].Borders.TopSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopStyle == BorderValues.CirclesRectangles);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSize == 6);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSpace == 5U);



                wordTable.Rows[1].Cells[1].Borders.BottomStyle = BorderValues.Safari;
                wordTable.Rows[1].Cells[1].Borders.BottomColorHex = SixLabors.ImageSharp.Color.Azure.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColorHex == SixLabors.ImageSharp.Color.Azure.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.BottomColor = Color.Cyan;
                wordTable.Rows[1].Cells[1].Borders.BottomSize = 8;
                wordTable.Rows[1].Cells[1].Borders.BottomSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomStyle == BorderValues.Safari);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColor == Color.Cyan);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSize == 8);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSpace == 5U);

                wordTable.Rows[1].Cells[1].Borders.StartStyle = BorderValues.DashSmallGap;
                wordTable.Rows[1].Cells[1].Borders.StartColorHex = SixLabors.ImageSharp.Color.Orange.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColorHex == SixLabors.ImageSharp.Color.Orange.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.StartColor = Color.Yellow;
                wordTable.Rows[1].Cells[1].Borders.StartSize = 24;
                wordTable.Rows[1].Cells[1].Borders.StartSpace = 10U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartStyle == BorderValues.DashSmallGap);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColor == Color.Yellow);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSpace == 10U);

                wordTable.Rows[1].Cells[1].Borders.EndStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.EndColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.EndColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.EndSize = 24;
                //wordTable.Rows[1].Cells[1].Borders.EndSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndSpace == null);


                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace == 5U);


                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor = Color.Aqua;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize = 16;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace = 1U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor == Color.Aqua);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize == 16);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace == 1U);

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
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.LeftSpace == 5U);






                wordTable.Rows[1].Cells[1].Borders.RightStyle = BorderValues.Double;
                wordTable.Rows[1].Cells[1].Borders.RightColorHex = SixLabors.ImageSharp.Color.Blue.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColorHex == SixLabors.ImageSharp.Color.Blue.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.RightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.RightSize = 4;
                wordTable.Rows[1].Cells[1].Borders.RightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightStyle == BorderValues.Double);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSize == 4);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.RightSpace == 5U);




                wordTable.Rows[1].Cells[1].Borders.TopStyle = BorderValues.CirclesRectangles;
                wordTable.Rows[1].Cells[1].Borders.TopColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopSize = 6;
                wordTable.Rows[1].Cells[1].Borders.TopSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopStyle == BorderValues.CirclesRectangles);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSize == 6);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopSpace == 5U);



                wordTable.Rows[1].Cells[1].Borders.BottomStyle = BorderValues.Safari;
                wordTable.Rows[1].Cells[1].Borders.BottomColorHex = SixLabors.ImageSharp.Color.Azure.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColorHex == SixLabors.ImageSharp.Color.Azure.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.BottomColor = Color.Cyan;
                wordTable.Rows[1].Cells[1].Borders.BottomSize = 8;
                wordTable.Rows[1].Cells[1].Borders.BottomSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomStyle == BorderValues.Safari);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomColor == Color.Cyan);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSize == 8);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.BottomSpace == 5U);

                wordTable.Rows[1].Cells[1].Borders.StartStyle = BorderValues.DashSmallGap;
                wordTable.Rows[1].Cells[1].Borders.StartColorHex = SixLabors.ImageSharp.Color.Orange.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColorHex == SixLabors.ImageSharp.Color.Orange.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.StartColor = Color.Yellow;
                wordTable.Rows[1].Cells[1].Borders.StartSize = 24;
                wordTable.Rows[1].Cells[1].Borders.StartSpace = 10U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartStyle == BorderValues.DashSmallGap);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartColor == Color.Yellow);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.StartSpace == 10U);

                wordTable.Rows[1].Cells[1].Borders.EndStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.EndColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.EndColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.EndSize = 24;
                //wordTable.Rows[1].Cells[1].Borders.EndSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.EndSpace == null);


                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor = Color.Gold;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize = 24;
                wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace = 5U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightColor == Color.Gold);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSize == 24);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopLeftToBottomRightSpace == 5U);


                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle = BorderValues.Dotted;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex = SixLabors.ImageSharp.Color.OrangeRed.ToHexColor();
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColorHex == SixLabors.ImageSharp.Color.OrangeRed.ToHexColor());
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor = Color.Aqua;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize = 16;
                wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace = 1U;

                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftStyle == BorderValues.Dotted);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftColor == Color.Aqua);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSize == 16);
                Assert.True(wordTable.Rows[1].Cells[1].Borders.TopRightToBottomLeftSpace == 1U);

                document.Save();
            }
        }
    }
}
