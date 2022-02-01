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
                document.Sections[0].Borders.LeftColor = System.Drawing.Color.Aqua;

                document.Sections[0].Borders.RightStyle = BorderValues.BirdsFlight;
                document.Sections[0].Borders.RightColor = System.Drawing.Color.Red;

                document.Sections[0].Borders.TopStyle = BorderValues.SharksTeeth;
                document.Sections[0].Borders.TopColor = System.Drawing.Color.GreenYellow;

                document.Sections[0].Borders.BottomStyle = BorderValues.Thick;
                document.Sections[0].Borders.BottomColor = System.Drawing.Color.Blue;


                Assert.True(document.Sections[0].Borders.LeftStyle == BorderValues.BabyPacifier);
                Assert.True(document.Sections[0].Borders.LeftColor.ToHexColor() == System.Drawing.Color.Aqua.ToHexColor());
                Assert.True(document.Sections[0].Borders.RightStyle == BorderValues.BirdsFlight);
                Assert.True(document.Sections[0].Borders.RightColor.ToHexColor() == System.Drawing.Color.Red.ToHexColor());
                Assert.True(document.Sections[0].Borders.TopStyle == BorderValues.SharksTeeth);
                Assert.True(document.Sections[0].Borders.TopColor.ToHexColor() == System.Drawing.Color.GreenYellow.ToHexColor());
                Assert.True(document.Sections[0].Borders.BottomStyle == BorderValues.Thick);
                Assert.True(document.Sections[0].Borders.BottomColor.ToHexColor() == System.Drawing.Color.Blue.ToHexColor());

                document.AddSection();

                document.Sections[1].Borders.LeftStyle = BorderValues.BabyRattle;
                document.Sections[1].Borders.LeftColor = System.Drawing.Color.LightYellow;

                document.Sections[1].Borders.RightStyle = BorderValues.ChainLink;
                document.Sections[1].Borders.RightColor = System.Drawing.Color.GreenYellow;

                document.Sections[1].Borders.TopStyle = BorderValues.Dashed;
                document.Sections[1].Borders.TopColor = System.Drawing.Color.OrangeRed;

                document.Sections[1].Borders.BottomStyle = BorderValues.DashSmallGap;
                document.Sections[1].Borders.BottomColor = System.Drawing.Color.DarkOliveGreen;


                Assert.True(document.Sections[0].Borders.LeftStyle == BorderValues.BabyPacifier);
                Assert.True(document.Sections[0].Borders.LeftColor.ToHexColor() == System.Drawing.Color.Aqua.ToHexColor());
                Assert.True(document.Sections[0].Borders.RightStyle == BorderValues.BirdsFlight);
                Assert.True(document.Sections[0].Borders.RightColor.ToHexColor() == System.Drawing.Color.Red.ToHexColor());
                Assert.True(document.Sections[0].Borders.TopStyle == BorderValues.SharksTeeth);
                Assert.True(document.Sections[0].Borders.TopColor.ToHexColor() == System.Drawing.Color.GreenYellow.ToHexColor());
                Assert.True(document.Sections[0].Borders.BottomStyle == BorderValues.Thick);
                Assert.True(document.Sections[0].Borders.BottomColor.ToHexColor() == System.Drawing.Color.Blue.ToHexColor());


                Assert.True(document.Sections[1].Borders.LeftStyle == BorderValues.BabyRattle);
                Assert.True(document.Sections[1].Borders.LeftColor.ToHexColor() == System.Drawing.Color.LightYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.RightStyle == BorderValues.ChainLink);
                Assert.True(document.Sections[1].Borders.RightColor.ToHexColor() == System.Drawing.Color.GreenYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.TopStyle == BorderValues.Dashed);
                Assert.True(document.Sections[1].Borders.TopColor.ToHexColor() == System.Drawing.Color.OrangeRed.ToHexColor());
                Assert.True(document.Sections[1].Borders.BottomStyle == BorderValues.DashSmallGap);
                Assert.True(document.Sections[1].Borders.BottomColor.ToHexColor() == System.Drawing.Color.DarkOliveGreen.ToHexColor());

                Assert.True(document.Borders.LeftStyle == BorderValues.BabyPacifier);
                Assert.True(document.Borders.LeftColor.ToHexColor() == System.Drawing.Color.Aqua.ToHexColor());
                Assert.True(document.Borders.RightStyle == BorderValues.BirdsFlight);
                Assert.True(document.Borders.RightColor.ToHexColor() == System.Drawing.Color.Red.ToHexColor());
                Assert.True(document.Borders.TopStyle == BorderValues.SharksTeeth);
                Assert.True(document.Borders.TopColor.ToHexColor() == System.Drawing.Color.GreenYellow.ToHexColor());
                Assert.True(document.Borders.BottomStyle == BorderValues.Thick);
                Assert.True(document.Borders.BottomColor.ToHexColor() == System.Drawing.Color.Blue.ToHexColor());


                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithBorders.docx"))) {

                Assert.True(document.Sections[0].Borders.LeftStyle == BorderValues.BabyPacifier);
                Assert.True(document.Sections[0].Borders.LeftColor.ToHexColor() == System.Drawing.Color.Aqua.ToHexColor());
                Assert.True(document.Sections[0].Borders.RightStyle == BorderValues.BirdsFlight);
                Assert.True(document.Sections[0].Borders.RightColor.ToHexColor() == System.Drawing.Color.Red.ToHexColor());
                Assert.True(document.Sections[0].Borders.TopStyle == BorderValues.SharksTeeth);
                Assert.True(document.Sections[0].Borders.TopColor.ToHexColor() == System.Drawing.Color.GreenYellow.ToHexColor());
                Assert.True(document.Sections[0].Borders.BottomStyle == BorderValues.Thick);
                Assert.True(document.Sections[0].Borders.BottomColor.ToHexColor() == System.Drawing.Color.Blue.ToHexColor());


                Assert.True(document.Sections[1].Borders.LeftStyle == BorderValues.BabyRattle);
                Assert.True(document.Sections[1].Borders.LeftColor.ToHexColor() == System.Drawing.Color.LightYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.RightStyle == BorderValues.ChainLink);
                Assert.True(document.Sections[1].Borders.RightColor.ToHexColor() == System.Drawing.Color.GreenYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.TopStyle == BorderValues.Dashed);
                Assert.True(document.Sections[1].Borders.TopColor.ToHexColor() == System.Drawing.Color.OrangeRed.ToHexColor());
                Assert.True(document.Sections[1].Borders.BottomStyle == BorderValues.DashSmallGap);
                Assert.True(document.Sections[1].Borders.BottomColor.ToHexColor() == System.Drawing.Color.DarkOliveGreen.ToHexColor());


                document.AddSection();
                document.Sections[2].Borders.LeftStyle = BorderValues.DotDash;
                document.Sections[2].Borders.LeftColor = System.Drawing.Color.OrangeRed;
                document.Sections[2].Borders.RightStyle = BorderValues.DotDotDash;
                document.Sections[2].Borders.RightColor = System.Drawing.Color.Goldenrod;
                document.Sections[2].Borders.TopStyle = BorderValues.DashDotStroked;
                document.Sections[2].Borders.TopColor = System.Drawing.Color.DarkKhaki;
                document.Sections[2].Borders.BottomStyle = BorderValues.BasicWhiteDashes;
                document.Sections[2].Borders.BottomColor = System.Drawing.Color.LightSkyBlue;

                Assert.True(document.Sections[0].Borders.LeftStyle == BorderValues.BabyPacifier);
                Assert.True(document.Sections[0].Borders.LeftColor.ToHexColor() == System.Drawing.Color.Aqua.ToHexColor());
                Assert.True(document.Sections[0].Borders.RightStyle == BorderValues.BirdsFlight);
                Assert.True(document.Sections[0].Borders.RightColor.ToHexColor() == System.Drawing.Color.Red.ToHexColor());
                Assert.True(document.Sections[0].Borders.TopStyle == BorderValues.SharksTeeth);
                Assert.True(document.Sections[0].Borders.TopColor.ToHexColor() == System.Drawing.Color.GreenYellow.ToHexColor());
                Assert.True(document.Sections[0].Borders.BottomStyle == BorderValues.Thick);
                Assert.True(document.Sections[0].Borders.BottomColor.ToHexColor() == System.Drawing.Color.Blue.ToHexColor());

                Assert.True(document.Sections[1].Borders.LeftStyle == BorderValues.BabyRattle);
                Assert.True(document.Sections[1].Borders.LeftColor.ToHexColor() == System.Drawing.Color.LightYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.RightStyle == BorderValues.ChainLink);
                Assert.True(document.Sections[1].Borders.RightColor.ToHexColor() == System.Drawing.Color.GreenYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.TopStyle == BorderValues.Dashed);
                Assert.True(document.Sections[1].Borders.TopColor.ToHexColor() == System.Drawing.Color.OrangeRed.ToHexColor());
                Assert.True(document.Sections[1].Borders.BottomStyle == BorderValues.DashSmallGap);
                Assert.True(document.Sections[1].Borders.BottomColor.ToHexColor() == System.Drawing.Color.DarkOliveGreen.ToHexColor());

                Assert.True(document.Sections[2].Borders.LeftStyle == BorderValues.DotDash);
                Assert.True(document.Sections[2].Borders.LeftColor.ToHexColor() == System.Drawing.Color.OrangeRed.ToHexColor());
                Assert.True(document.Sections[2].Borders.RightStyle == BorderValues.DotDotDash);
                Assert.True(document.Sections[2].Borders.RightColor.ToHexColor() == System.Drawing.Color.Goldenrod.ToHexColor());
                Assert.True(document.Sections[2].Borders.TopStyle == BorderValues.DashDotStroked);
                Assert.True(document.Sections[2].Borders.TopColor.ToHexColor() == System.Drawing.Color.DarkKhaki.ToHexColor());
                Assert.True(document.Sections[2].Borders.BottomStyle == BorderValues.BasicWhiteDashes);
                Assert.True(document.Sections[2].Borders.BottomColor.ToHexColor() == System.Drawing.Color.LightSkyBlue.ToHexColor());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithBorders.docx"))) {

                Assert.True(document.Sections[0].Borders.LeftStyle == BorderValues.BabyPacifier);
                Assert.True(document.Sections[0].Borders.LeftColor.ToHexColor() == System.Drawing.Color.Aqua.ToHexColor());
                Assert.True(document.Sections[0].Borders.RightStyle == BorderValues.BirdsFlight);
                Assert.True(document.Sections[0].Borders.RightColor.ToHexColor() == System.Drawing.Color.Red.ToHexColor());
                Assert.True(document.Sections[0].Borders.TopStyle == BorderValues.SharksTeeth);
                Assert.True(document.Sections[0].Borders.TopColor.ToHexColor() == System.Drawing.Color.GreenYellow.ToHexColor());
                Assert.True(document.Sections[0].Borders.BottomStyle == BorderValues.Thick);
                Assert.True(document.Sections[0].Borders.BottomColor.ToHexColor() == System.Drawing.Color.Blue.ToHexColor());

                Assert.True(document.Sections[1].Borders.LeftStyle == BorderValues.BabyRattle);
                Assert.True(document.Sections[1].Borders.LeftColor.ToHexColor() == System.Drawing.Color.LightYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.RightStyle == BorderValues.ChainLink);
                Assert.True(document.Sections[1].Borders.RightColor.ToHexColor() == System.Drawing.Color.GreenYellow.ToHexColor());
                Assert.True(document.Sections[1].Borders.TopStyle == BorderValues.Dashed);
                Assert.True(document.Sections[1].Borders.TopColor.ToHexColor() == System.Drawing.Color.OrangeRed.ToHexColor());
                Assert.True(document.Sections[1].Borders.BottomStyle == BorderValues.DashSmallGap);
                Assert.True(document.Sections[1].Borders.BottomColor.ToHexColor() == System.Drawing.Color.DarkOliveGreen.ToHexColor());

                Assert.True(document.Sections[2].Borders.LeftStyle == BorderValues.DotDash);
                Assert.True(document.Sections[2].Borders.LeftColor.ToHexColor() == System.Drawing.Color.OrangeRed.ToHexColor());
                Assert.True(document.Sections[2].Borders.RightStyle == BorderValues.DotDotDash);
                Assert.True(document.Sections[2].Borders.RightColor.ToHexColor() == System.Drawing.Color.Goldenrod.ToHexColor());
                Assert.True(document.Sections[2].Borders.TopStyle == BorderValues.DashDotStroked);
                Assert.True(document.Sections[2].Borders.TopColor.ToHexColor() == System.Drawing.Color.DarkKhaki.ToHexColor());
                Assert.True(document.Sections[2].Borders.BottomStyle == BorderValues.BasicWhiteDashes);
                Assert.True(document.Sections[2].Borders.BottomColor.ToHexColor() == System.Drawing.Color.LightSkyBlue.ToHexColor());

                document.Borders.LeftStyle = BorderValues.DotDash;
                document.Borders.LeftColor = System.Drawing.Color.OrangeRed;
                document.Borders.RightStyle = BorderValues.DotDotDash;
                document.Borders.RightColor = System.Drawing.Color.Goldenrod;
                document.Borders.TopStyle = BorderValues.DashDotStroked;
                document.Borders.TopColor = System.Drawing.Color.DarkKhaki;
                document.Borders.BottomStyle = BorderValues.BasicWhiteDashes;
                document.Borders.BottomColor = System.Drawing.Color.LightSkyBlue;

                Assert.True(document.Sections[0].Borders.LeftStyle == BorderValues.DotDash);
                Assert.True(document.Sections[0].Borders.LeftColor.ToHexColor() == System.Drawing.Color.OrangeRed.ToHexColor());
                Assert.True(document.Sections[0].Borders.RightStyle == BorderValues.DotDotDash);
                Assert.True(document.Sections[0].Borders.RightColor.ToHexColor() == System.Drawing.Color.Goldenrod.ToHexColor());
                Assert.True(document.Sections[0].Borders.TopStyle == BorderValues.DashDotStroked);
                Assert.True(document.Sections[0].Borders.TopColor.ToHexColor() == System.Drawing.Color.DarkKhaki.ToHexColor());
                Assert.True(document.Sections[0].Borders.BottomStyle == BorderValues.BasicWhiteDashes);
                Assert.True(document.Sections[0].Borders.BottomColor.ToHexColor() == System.Drawing.Color.LightSkyBlue.ToHexColor());

                document.Save();
            }
        }
    }
}
