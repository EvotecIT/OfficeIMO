using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithBorders() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithBorders.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.Sections[0].Borders.LeftStyle = BorderValues.BabyPacifier;
                document.Sections[0].Borders.LeftColor = OfficeIMO.Drawing.OfficeColor.Aqua;
                document.Sections[0].Borders.LeftSpace = 10;
                document.Sections[0].Borders.LeftSize = 24;

                document.Sections[0].Borders.RightStyle = BorderValues.BirdsFlight;
                document.Sections[0].Borders.RightColor = OfficeIMO.Drawing.OfficeColor.Red;

                document.Sections[0].Borders.TopStyle = BorderValues.SharksTeeth;
                document.Sections[0].Borders.TopColor = OfficeIMO.Drawing.OfficeColor.GreenYellow;

                document.Sections[0].Borders.BottomStyle = BorderValues.Thick;
                document.Sections[0].Borders.BottomColor = OfficeIMO.Drawing.OfficeColor.Blue;
                document.Sections[0].Borders.BottomSpace = 15;
                document.Sections[0].Borders.BottomSize = 18;



                Assert.True(document.Sections[0].Borders.LeftStyle == BorderValues.BabyPacifier);
                Assert.True(document.Sections[0].Borders.LeftColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Aqua.ToHexColor());
                Assert.Equal(10U, document.Sections[0].Borders.LeftSpace!.Value);
                Assert.Equal(24U, document.Sections[0].Borders.LeftSize!.Value);

                Assert.True(document.Sections[0].Borders.RightStyle == BorderValues.BirdsFlight);
                Assert.True(document.Sections[0].Borders.RightColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Red.ToHexColor());
                Assert.True(document.Sections[0].Borders.RightSpace == null);
                Assert.True(document.Sections[0].Borders.RightSize == null);

                Assert.True(document.Sections[0].Borders.TopStyle == BorderValues.SharksTeeth);
                Assert.True(document.Sections[0].Borders.TopColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.GreenYellow.ToHexColor());
                Assert.True(document.Sections[0].Borders.TopSpace == null);
                Assert.True(document.Sections[0].Borders.TopSize == null);

                Assert.True(document.Sections[0].Borders.BottomStyle == BorderValues.Thick);
                Assert.True(document.Sections[0].Borders.BottomColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Blue.ToHexColor());
                Assert.Equal(15U, document.Sections[0].Borders.BottomSpace!.Value);
                Assert.Equal(18U, document.Sections[0].Borders.BottomSize!.Value);


                document.AddSection();

                document.Sections[1].Borders.LeftStyle = BorderValues.BabyRattle;
                document.Sections[1].Borders.LeftColor = OfficeIMO.Drawing.OfficeColor.LightYellow;
                document.Sections[1].Borders.LeftShadow = true;
                document.Sections[1].Borders.LeftFrame = true;


                document.Sections[1].Borders.RightStyle = BorderValues.ChainLink;
                document.Sections[1].Borders.RightColor = OfficeIMO.Drawing.OfficeColor.GreenYellow;
                document.Sections[1].Borders.RightShadow = true;
                document.Sections[1].Borders.RightFrame = false;

                document.Sections[1].Borders.TopStyle = BorderValues.Dashed;
                document.Sections[1].Borders.TopColor = OfficeIMO.Drawing.OfficeColor.OrangeRed;

                document.Sections[1].Borders.BottomStyle = BorderValues.DashSmallGap;
                document.Sections[1].Borders.BottomColor = OfficeIMO.Drawing.OfficeColor.DarkOliveGreen;
                document.Sections[1].Borders.BottomShadow = false;


                Assert.True(document.Sections[0].Borders.LeftStyle == BorderValues.BabyPacifier);
                Assert.True(document.Sections[0].Borders.LeftColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Aqua.ToHexColor());
                Assert.Equal(10U, document.Sections[0].Borders.LeftSpace!.Value);
                Assert.Equal(24U, document.Sections[0].Borders.LeftSize!.Value);

                Assert.True(document.Sections[0].Borders.RightStyle == BorderValues.BirdsFlight);
                Assert.True(document.Sections[0].Borders.RightColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Red.ToHexColor());
                Assert.True(document.Sections[0].Borders.RightSpace == null);
                Assert.True(document.Sections[0].Borders.RightSize == null);

                Assert.True(document.Sections[0].Borders.TopStyle == BorderValues.SharksTeeth);
                Assert.True(document.Sections[0].Borders.TopColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.GreenYellow.ToHexColor());
                Assert.True(document.Sections[0].Borders.TopSpace == null);
                Assert.True(document.Sections[0].Borders.TopSize == null);

                Assert.True(document.Sections[0].Borders.BottomStyle == BorderValues.Thick);
                Assert.True(document.Sections[0].Borders.BottomColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Blue.ToHexColor());
                Assert.Equal(15U, document.Sections[0].Borders.BottomSpace!.Value);
                Assert.Equal(18U, document.Sections[0].Borders.BottomSize!.Value);


                Assert.True(document.Sections[1].Borders.LeftStyle == BorderValues.BabyRattle);
                Assert.True(document.Sections[1].Borders.LeftColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.LightYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.LeftShadow == true);
                Assert.True(document.Sections[1].Borders.LeftFrame == true);

                Assert.True(document.Sections[1].Borders.RightStyle == BorderValues.ChainLink);
                Assert.True(document.Sections[1].Borders.RightColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.GreenYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.RightShadow == true);
                Assert.True(document.Sections[1].Borders.RightFrame == false);

                Assert.True(document.Sections[1].Borders.TopStyle == BorderValues.Dashed);
                Assert.True(document.Sections[1].Borders.TopColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToHexColor());
                Assert.True(document.Sections[1].Borders.TopShadow == null);
                Assert.True(document.Sections[1].Borders.TopFrame == null);

                Assert.True(document.Sections[1].Borders.BottomStyle == BorderValues.DashSmallGap);
                Assert.True(document.Sections[1].Borders.BottomColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.DarkOliveGreen.ToHexColor());
                Assert.True(document.Sections[1].Borders.BottomShadow == false);
                Assert.True(document.Sections[1].Borders.BottomFrame == null);

                Assert.True(document.Borders.LeftStyle == BorderValues.BabyPacifier);
                Assert.True(document.Borders.LeftColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Aqua.ToHexColor());
                Assert.True(document.Borders.RightStyle == BorderValues.BirdsFlight);
                Assert.True(document.Borders.RightColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Red.ToHexColor());
                Assert.True(document.Borders.TopStyle == BorderValues.SharksTeeth);
                Assert.True(document.Borders.TopColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.GreenYellow.ToHexColor());
                Assert.True(document.Borders.BottomStyle == BorderValues.Thick);
                Assert.True(document.Borders.BottomColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Blue.ToHexColor());


                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithBorders.docx"))) {

                Assert.True(document.Sections[0].Borders.LeftStyle == BorderValues.BabyPacifier);
                Assert.True(document.Sections[0].Borders.LeftColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Aqua.ToHexColor());
                Assert.Equal(10U, document.Sections[0].Borders.LeftSpace!.Value);
                Assert.Equal(24U, document.Sections[0].Borders.LeftSize!.Value);

                Assert.True(document.Sections[0].Borders.RightStyle == BorderValues.BirdsFlight);
                Assert.True(document.Sections[0].Borders.RightColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Red.ToHexColor());
                Assert.True(document.Sections[0].Borders.RightSpace == null);
                Assert.True(document.Sections[0].Borders.RightSize == null);

                Assert.True(document.Sections[0].Borders.TopStyle == BorderValues.SharksTeeth);
                Assert.True(document.Sections[0].Borders.TopColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.GreenYellow.ToHexColor());
                Assert.True(document.Sections[0].Borders.TopSpace == null);
                Assert.True(document.Sections[0].Borders.TopSize == null);

                Assert.True(document.Sections[0].Borders.BottomStyle == BorderValues.Thick);
                Assert.True(document.Sections[0].Borders.BottomColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Blue.ToHexColor());
                Assert.Equal(15U, document.Sections[0].Borders.BottomSpace!.Value);
                Assert.Equal(18U, document.Sections[0].Borders.BottomSize!.Value);


                Assert.True(document.Sections[1].Borders.LeftStyle == BorderValues.BabyRattle);
                Assert.True(document.Sections[1].Borders.LeftColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.LightYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.RightStyle == BorderValues.ChainLink);
                Assert.True(document.Sections[1].Borders.RightColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.GreenYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.TopStyle == BorderValues.Dashed);
                Assert.True(document.Sections[1].Borders.TopColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToHexColor());
                Assert.True(document.Sections[1].Borders.BottomStyle == BorderValues.DashSmallGap);
                Assert.True(document.Sections[1].Borders.BottomColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.DarkOliveGreen.ToHexColor());


                document.AddSection();
                document.Sections[2].Borders.LeftStyle = BorderValues.DotDash;
                document.Sections[2].Borders.LeftColor = OfficeIMO.Drawing.OfficeColor.OrangeRed;
                document.Sections[2].Borders.RightStyle = BorderValues.DotDotDash;
                document.Sections[2].Borders.RightColor = OfficeIMO.Drawing.OfficeColor.Goldenrod;
                document.Sections[2].Borders.TopStyle = BorderValues.DashDotStroked;
                document.Sections[2].Borders.TopColor = OfficeIMO.Drawing.OfficeColor.DarkKhaki;
                document.Sections[2].Borders.BottomStyle = BorderValues.BasicWhiteDashes;
                document.Sections[2].Borders.BottomColor = OfficeIMO.Drawing.OfficeColor.LightSkyBlue;

                Assert.True(document.Sections[0].Borders.LeftStyle == BorderValues.BabyPacifier);
                Assert.True(document.Sections[0].Borders.LeftColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Aqua.ToHexColor());
                Assert.True(document.Sections[0].Borders.RightStyle == BorderValues.BirdsFlight);
                Assert.True(document.Sections[0].Borders.RightColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Red.ToHexColor());
                Assert.True(document.Sections[0].Borders.TopStyle == BorderValues.SharksTeeth);
                Assert.True(document.Sections[0].Borders.TopColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.GreenYellow.ToHexColor());
                Assert.True(document.Sections[0].Borders.BottomStyle == BorderValues.Thick);
                Assert.True(document.Sections[0].Borders.BottomColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Blue.ToHexColor());

                Assert.True(document.Sections[1].Borders.LeftStyle == BorderValues.BabyRattle);
                Assert.True(document.Sections[1].Borders.LeftColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.LightYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.RightStyle == BorderValues.ChainLink);
                Assert.True(document.Sections[1].Borders.RightColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.GreenYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.TopStyle == BorderValues.Dashed);
                Assert.True(document.Sections[1].Borders.TopColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToHexColor());
                Assert.True(document.Sections[1].Borders.BottomStyle == BorderValues.DashSmallGap);
                Assert.True(document.Sections[1].Borders.BottomColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.DarkOliveGreen.ToHexColor());

                Assert.True(document.Sections[2].Borders.LeftStyle == BorderValues.DotDash);
                Assert.True(document.Sections[2].Borders.LeftColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToHexColor());
                Assert.True(document.Sections[2].Borders.RightStyle == BorderValues.DotDotDash);
                Assert.True(document.Sections[2].Borders.RightColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Goldenrod.ToHexColor());
                Assert.True(document.Sections[2].Borders.TopStyle == BorderValues.DashDotStroked);
                Assert.True(document.Sections[2].Borders.TopColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.DarkKhaki.ToHexColor());
                Assert.True(document.Sections[2].Borders.BottomStyle == BorderValues.BasicWhiteDashes);
                Assert.True(document.Sections[2].Borders.BottomColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.LightSkyBlue.ToHexColor());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithBorders.docx"))) {

                Assert.True(document.Sections[0].Borders.LeftStyle == BorderValues.BabyPacifier);
                Assert.True(document.Sections[0].Borders.LeftColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Aqua.ToHexColor());
                Assert.True(document.Sections[0].Borders.RightStyle == BorderValues.BirdsFlight);
                Assert.True(document.Sections[0].Borders.RightColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Red.ToHexColor());
                Assert.True(document.Sections[0].Borders.TopStyle == BorderValues.SharksTeeth);
                Assert.True(document.Sections[0].Borders.TopColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.GreenYellow.ToHexColor());
                Assert.True(document.Sections[0].Borders.BottomStyle == BorderValues.Thick);
                Assert.True(document.Sections[0].Borders.BottomColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Blue.ToHexColor());

                Assert.True(document.Sections[1].Borders.LeftStyle == BorderValues.BabyRattle);
                Assert.True(document.Sections[1].Borders.LeftColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.LightYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.RightStyle == BorderValues.ChainLink);
                Assert.True(document.Sections[1].Borders.RightColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.GreenYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.TopStyle == BorderValues.Dashed);
                Assert.True(document.Sections[1].Borders.TopColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToHexColor());
                Assert.True(document.Sections[1].Borders.BottomStyle == BorderValues.DashSmallGap);
                Assert.True(document.Sections[1].Borders.BottomColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.DarkOliveGreen.ToHexColor());

                Assert.True(document.Sections[2].Borders.LeftStyle == BorderValues.DotDash);
                Assert.True(document.Sections[2].Borders.LeftColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToHexColor());
                Assert.True(document.Sections[2].Borders.RightStyle == BorderValues.DotDotDash);
                Assert.True(document.Sections[2].Borders.RightColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Goldenrod.ToHexColor());
                Assert.True(document.Sections[2].Borders.TopStyle == BorderValues.DashDotStroked);
                Assert.True(document.Sections[2].Borders.TopColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.DarkKhaki.ToHexColor());
                Assert.True(document.Sections[2].Borders.BottomStyle == BorderValues.BasicWhiteDashes);
                Assert.True(document.Sections[2].Borders.BottomColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.LightSkyBlue.ToHexColor());

                document.Borders.LeftStyle = BorderValues.DotDash;
                document.Borders.LeftColor = OfficeIMO.Drawing.OfficeColor.OrangeRed;
                document.Borders.RightStyle = BorderValues.DotDotDash;
                document.Borders.RightColor = OfficeIMO.Drawing.OfficeColor.Goldenrod;
                document.Borders.TopStyle = BorderValues.DashDotStroked;
                document.Borders.TopColor = OfficeIMO.Drawing.OfficeColor.DarkKhaki;
                document.Borders.BottomStyle = BorderValues.BasicWhiteDashes;
                document.Borders.BottomColor = OfficeIMO.Drawing.OfficeColor.LightSkyBlue;

                Assert.True(document.Sections[0].Borders.LeftStyle == BorderValues.DotDash);
                Assert.True(document.Sections[0].Borders.LeftColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToHexColor());
                Assert.True(document.Sections[0].Borders.RightStyle == BorderValues.DotDotDash);
                Assert.True(document.Sections[0].Borders.RightColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.Goldenrod.ToHexColor());
                Assert.True(document.Sections[0].Borders.TopStyle == BorderValues.DashDotStroked);
                Assert.True(document.Sections[0].Borders.TopColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.DarkKhaki.ToHexColor());
                Assert.True(document.Sections[0].Borders.BottomStyle == BorderValues.BasicWhiteDashes);
                Assert.True(document.Sections[0].Borders.BottomColor.ToHexColor() == OfficeIMO.Drawing.OfficeColor.LightSkyBlue.ToHexColor());

                Assert.True(document.Sections[0].Borders.LeftColorHex == OfficeIMO.Drawing.OfficeColor.OrangeRed.ToHexColor());
                Assert.True(document.Sections[0].Borders.RightColorHex == OfficeIMO.Drawing.OfficeColor.Goldenrod.ToHexColor());
                Assert.True(document.Sections[0].Borders.TopColorHex == OfficeIMO.Drawing.OfficeColor.DarkKhaki.ToHexColor());
                Assert.True(document.Sections[0].Borders.BottomColorHex == OfficeIMO.Drawing.OfficeColor.LightSkyBlue.ToHexColor());

                document.Borders.LeftColorHex = OfficeIMO.Drawing.OfficeColor.Yellow.ToHexColor();
                document.Borders.RightColorHex = OfficeIMO.Drawing.OfficeColor.DarkOliveGreen.ToHexColor();
                document.Borders.TopColorHex = OfficeIMO.Drawing.OfficeColor.LightSkyBlue.ToHexColor();
                document.Borders.BottomColorHex = OfficeIMO.Drawing.OfficeColor.Beige.ToHexColor();

                Assert.True(document.Sections[0].Borders.LeftColorHex == OfficeIMO.Drawing.OfficeColor.Yellow.ToHexColor());
                Assert.True(document.Sections[0].Borders.RightColorHex == OfficeIMO.Drawing.OfficeColor.DarkOliveGreen.ToHexColor());
                Assert.True(document.Sections[0].Borders.TopColorHex == OfficeIMO.Drawing.OfficeColor.LightSkyBlue.ToHexColor());
                Assert.True(document.Sections[0].Borders.BottomColorHex == OfficeIMO.Drawing.OfficeColor.Beige.ToHexColor());

                document.Save();
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithBordersBuiltin() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithBordersBuiltin.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.Sections[0].SetBorders(WordBorder.Box);
                Assert.True(document.Sections[0].Borders.Type == WordBorder.Box);
                document.AddSection().SetBorders(WordBorder.Shadow);
                Assert.True(document.Sections[1].Borders.Type == WordBorder.Shadow);
                document.AddSection().SetBorders(WordBorder.Shadow);
                Assert.True(document.Sections[2].Borders.Type == WordBorder.Shadow);
                document.AddSection();
                Assert.True(document.Sections[3].Borders.Type == WordBorder.Shadow);
                document.Sections[3].SetBorders(WordBorder.None);
                Assert.True(document.Sections[3].Borders.Type == WordBorder.None);
                document.Save(false);

                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithBordersBuiltin.docx"))) {
                Assert.True(document.Sections[0].Borders.Type == WordBorder.Box);
                Assert.True(document.Sections[1].Borders.Type == WordBorder.Shadow);
                Assert.True(document.Sections[2].Borders.Type == WordBorder.Shadow);
                Assert.True(document.Sections[3].Borders.Type == WordBorder.None);

                document.AddSection().SetBorders(WordBorder.Box);
                Assert.True(document.Sections[4].Borders.Type == WordBorder.Box);
                document.Sections[2].Borders.Type = WordBorder.None;
                Assert.True(document.Sections[2].Borders.Type == WordBorder.None);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithBordersBuiltin.docx"))) {
                Assert.True(document.Sections[0].Borders.Type == WordBorder.Box);
                Assert.True(document.Sections[1].Borders.Type == WordBorder.Shadow);
                Assert.True(document.Sections[2].Borders.Type == WordBorder.None);
                Assert.True(document.Sections[3].Borders.Type == WordBorder.None);
                Assert.True(document.Sections[4].Borders.Type == WordBorder.Box);

                document.Save();
            }
        }
    }
}
