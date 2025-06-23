using System;
using System.IO;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

using Path = System.IO.Path;

namespace OfficeIMO.Tests {
    public partial class Word {

        [Fact]
        public void Test_CreatingWordDocumentWithImages() {
            var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithImages.docx");
            using var document = WordDocument.Create(filePath);

            document.BuiltinDocumentProperties.Title = "This is sparta";
            document.BuiltinDocumentProperties.Creator = "Przemek";

            var paragraph = document.AddParagraph("This paragraph starts with some text");
            paragraph.Text = "0th This paragraph started with some other text and was overwritten and made bold.";

            // lets add image to paragraph
            paragraph.AddImage(Path.Combine(_directoryWithImages, "PrzemyslawKlysAndKulkozaurr.jpg"), 22, 22);
            Assert.True(paragraph.Image.WrapText == WrapTextImage.InLineWithText);

            paragraph.Image.WrapText = WrapTextImage.BehindText;

            Assert.True(paragraph.Image.WrapText == WrapTextImage.BehindText);

            paragraph.Image.WrapText = WrapTextImage.InLineWithText;

            Assert.True(paragraph.Image.WrapText == WrapTextImage.InLineWithText);

            var paragraph5 = paragraph.AddText("and more text");
            paragraph5.Bold = true;

            Assert.Equal(22d, paragraph.Image.Height!.Value, 15);
            Assert.Equal(22d, paragraph.Image.Width!.Value, 15);

            document.AddParagraph("This adds another picture with 500x500");

            var filePathImage = Path.Combine(_directoryWithImages, "Kulek.jpg");
            var paragraph2 = document.AddParagraph();
            paragraph2.AddImage(filePathImage, 500, 500);
            //paragraph2.Image.BlackWiteMode = BlackWhiteModeValues.GrayWhite;
            paragraph2.Image.Rotation = 180;
            paragraph2.Image.Shape = ShapeTypeValues.ActionButtonMovie;

            Assert.Equal(500d, paragraph2.Image.Height!.Value, 15);
            Assert.Equal(500d, paragraph2.Image.Width!.Value, 15);

            document.AddParagraph("This adds another picture with 100x100");

            var paragraph3 = document.AddParagraph();
            paragraph3.AddImage(filePathImage, 100, 100);

            // we add paragraph with an image
            var paragraph4 = document.AddParagraph();
            paragraph4.AddImage(filePathImage);

            // we can also overwrite height later on
            paragraph4.Image.Height = 50;
            paragraph4.Image.Width = 50;
            // this doesn't work
            paragraph4.Image.HorizontalFlip = true;


            Assert.Equal(50d, paragraph4.Image.Height.Value, 15);
            Assert.Equal(50d, paragraph4.Image.Width.Value, 15);

            // or we can get any image and overwrite it's size
            document.Images[0].Height = 200;
            document.Images[0].Width = 200;

            Assert.Equal(200d, document.Images[0].Height.Value, 15);
            Assert.Equal(200d, document.Images[0].Width.Value, 15);

            var fileToSave = Path.Combine(_directoryDocuments, "CreatedDocumentWithImagesPrzemyslawKlysAndKulkozaurr.jpg");
            document.Images[0].SaveToFile(fileToSave);

            var fileInfo = new FileInfo(fileToSave);

            Assert.True(fileInfo.Length > 0);
            Assert.True(File.Exists(fileToSave));

            Assert.Equal(4, document.Images.Count);
            Assert.Equal(4, document.Sections[0].Images.Count);
            document.Save(false);

            Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
        }


        [Fact]
        public void Test_CreatingWordDocumentWithImagesHeadersAndFooters() {
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

            var filePath = Path.Combine(_directoryWithFiles, "CreatingWordDocumentWithImagesHeadersAndFooters.docx");
            using var document = WordDocument.Create(filePath);

            var image = Path.Combine(_directoryWithImages, "PrzemyslawKlysAndKulkozaurr.jpg");

            document.AddHeadersAndFooters();
            document.DifferentOddAndEvenPages = true;

            Assert.True(document.Header.Default.Images.Count == 0);
            Assert.True(document.Header.Even.Images.Count == 0);
            Assert.True(document.Images.Count == 0);
            Assert.True(document.Footer.Default.Images.Count == 0);
            Assert.True(document.Footer.Even.Images.Count == 0);


            var header = document.Header.Default;
            // add image to header, directly to paragraph
            header.AddParagraph().AddImage(image, 100, 100);

            var footer = document.Footer.Default;
            // add image to footer, directly to paragraph
            footer.AddParagraph().AddImage(image, 200, 200);

            Assert.True(document.Header.Default.Images.Count == 1);
            Assert.True(document.Header.Even.Images.Count == 0);
            Assert.True(document.Images.Count == 0);
            Assert.True(document.Footer.Default.Images.Count == 1);
            Assert.True(document.Footer.Even.Images.Count == 0);

            const string fileNameImage = "Kulek.jpg";
            var filePathImage = System.IO.Path.Combine(imagePaths, fileNameImage);
            var paragraph = document.AddParagraph();
            using (var imageStream = System.IO.File.OpenRead(filePathImage)) {
                paragraph.AddImage(imageStream, fileNameImage, 300, 300);
            }
            Assert.True(document.Header.Default.Images.Count == 1);
            Assert.True(document.Header.Even.Images.Count == 0);

            Assert.True(document.Images.Count == 1);
            Assert.True(document.Footer.Default.Images.Count == 1);
            Assert.True(document.Footer.Even.Images.Count == 0);

            Assert.True(document.Images[0].FileName == fileNameImage);
            Assert.True(document.Images[0].Rotation == null);
            Assert.True(document.Images[0].Width == 300);
            Assert.True(document.Images[0].Height == 300);

            const string fileNameImageEvotec = "EvotecLogo.png";
            var filePathImageEvotec = System.IO.Path.Combine(imagePaths, fileNameImageEvotec);
            var paragraphHeader = document.Header.Even.AddParagraph();
            using (var imageStream = System.IO.File.OpenRead(filePathImageEvotec)) {
                paragraphHeader.AddImage(imageStream, fileNameImageEvotec, 300, 300, WrapTextImage.InLineWithText, "This is a test");
                Assert.True(paragraphHeader.Image.CompressionQuality == BlipCompressionValues.Print);
            }

            Assert.True(document.Header.Default.Images.Count == 1);
            Assert.True(document.Header.Even.Images.Count == 1);
            Assert.True(document.Header.Even.Images[0].FileName == fileNameImageEvotec);
            Assert.True(document.Header.Even.Images[0].Description == "This is a test");

            document.Header.Even.Images[0].Description = "Different description";
            Assert.True(document.Header.Even.Images[0].VerticalFlip == null);
            Assert.True(document.Header.Even.Images[0].HorizontalFlip == null);
            document.Header.Even.Images[0].VerticalFlip = true;

            Assert.True(document.Header.Even.Images[0].Description == "Different description");
            Assert.True(document.Header.Even.Images[0].VerticalFlip == true);
            Assert.True(document.Header.Even.Images[0].CompressionQuality == BlipCompressionValues.Print);
            document.Header.Even.Images[0].CompressionQuality = BlipCompressionValues.HighQualityPrint;
            Assert.True(document.Header.Even.Images[0].CompressionQuality == BlipCompressionValues.HighQualityPrint);

            document.Save();
        }



        [Fact]
        public void Test_LoadingWordDocumentWithImages() {
            var documentsPaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Documents");
            var filePath = Path.Combine(documentsPaths, "DocumentWithImagesWraps.docx");
            using (var document = WordDocument.Load(filePath)) {
                Assert.True(document.Paragraphs.Count == 36);
                Assert.True(document.Images.Count == 4);
                Assert.True(document.Header.Default.Images.Count == 1);
                Assert.True(document.Footer.Default.Images.Count == 0);

                Assert.True(document.Images[0].WrapText == WrapTextImage.InLineWithText);
                Assert.True(document.Images[1].WrapText == WrapTextImage.Square);
                Assert.True(document.Images[2].WrapText == WrapTextImage.InFrontOfText);
                Assert.True(document.Images[3].WrapText == WrapTextImage.BehindText);
                Assert.True(document.Header.Default.Images[0].WrapText == WrapTextImage.InLineWithText);

                Assert.True(document.Images[0].Shape == ShapeTypeValues.Rectangle);
                Assert.True(document.Images[1].Shape == ShapeTypeValues.Rectangle);
                Assert.True(document.Images[2].Shape == ShapeTypeValues.Rectangle);
                Assert.True(document.Images[3].Shape == ShapeTypeValues.Rectangle);

                document.Images[0].Shape = ShapeTypeValues.Cloud;

                Assert.True(document.Images[0].Shape == ShapeTypeValues.Cloud);
                document.Save(false);
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithImagesWraps() {
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithImagesWraps.docx");
            using var document = WordDocument.Create(filePath);

            const string fileNameImage = "Kulek.jpg";
            var filePathImage = System.IO.Path.Combine(imagePaths, fileNameImage);

            var paragraph1 = document.AddParagraph("This is a test document with images wraps");
            paragraph1.AddImage(filePathImage, 100, 100, WrapTextImage.InLineWithText);

            var paragraph2 = document.AddParagraph("This is a test document with images wraps");
            paragraph2.AddImage(filePathImage, 100, 100, WrapTextImage.BehindText);

            var paragraph3 = document.AddParagraph("This is a test document with images wraps");
            paragraph3.AddImage(filePathImage, 100, 100, WrapTextImage.InFrontOfText);

            var paragraph4 = document.AddParagraph("This is a test document with images wraps");
            paragraph4.AddImage(filePathImage, 100, 100, WrapTextImage.TopAndBottom);

            var paragraph5 = document.AddParagraph("This is a test document with images wraps");
            paragraph5.AddImage(filePathImage, 100, 100, WrapTextImage.Square);

            var paragraph6 = document.AddParagraph("This is a test document with images wraps");
            paragraph6.AddImage(filePathImage, 100, 100, WrapTextImage.Tight);

            var paragraph7 = document.AddParagraph("This is a test document with images wraps");
            paragraph7.AddImage(filePathImage, 100, 100, WrapTextImage.Through);

            Assert.True(document.Paragraphs.Count == 7);
            Assert.True(document.Paragraphs[0].Image.WrapText == WrapTextImage.InLineWithText);
            Assert.True(document.Paragraphs[1].Image.WrapText == WrapTextImage.BehindText);
            Assert.True(document.Paragraphs[2].Image.WrapText == WrapTextImage.InFrontOfText);
            Assert.True(document.Paragraphs[3].Image.WrapText == WrapTextImage.TopAndBottom);
            Assert.True(document.Paragraphs[4].Image.WrapText == WrapTextImage.Square);
            Assert.True(document.Paragraphs[5].Image.WrapText == WrapTextImage.Tight);
            Assert.True(document.Paragraphs[6].Image.WrapText == WrapTextImage.Through);

            document.Save(false);

            Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
        }

        [Fact]
        public void Test_CreatingWordDocumentWithFixedImages() {
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            var filePath = Path.Combine(_directoryWithFiles, "CreatingWordDocumentWithFixedImages.docx");
            using var document = WordDocument.Create(filePath);

            const string fileNameImage = "Kulek.jpg";
            var filePathImage = System.IO.Path.Combine(imagePaths, fileNameImage);

            var paragraph1 = document.AddParagraph("This is a test document with images wraps");
            paragraph1.AddImage(filePathImage, 100, 100, WrapTextImage.InLineWithText);

            var paragraph2 = document.AddParagraph("This is a test document with images wraps");
            paragraph2.AddImage(filePathImage, 100, 100, WrapTextImage.Square);

            Assert.True(document.Paragraphs.Count == 2);
            Assert.True(document.Paragraphs[0].Image.WrapText == WrapTextImage.InLineWithText);
            Assert.True(document.Paragraphs[1].Image.WrapText == WrapTextImage.Square);
            Assert.Throws<System.InvalidOperationException>(() => document.Paragraphs[0].Image.horizontalPosition);
            Assert.Throws<System.InvalidOperationException>(() => document.Paragraphs[0].Image.verticalPosition);

            int emus = 914400;
            HorizontalRelativePositionValues hRelativeFromGood = HorizontalRelativePositionValues.RightMargin;
            HorizontalRelativePositionValues hRelativeFromFail = HorizontalRelativePositionValues.LeftMargin;
            VerticalRelativePositionValues vRelativeFromGood = VerticalRelativePositionValues.Page;
            VerticalRelativePositionValues vRelativeFromFail = VerticalRelativePositionValues.Line;

            HorizontalPosition horizontalPosition1 = new HorizontalPosition() {
                RelativeFrom = hRelativeFromGood,
                PositionOffset = new PositionOffset { Text = $"{emus}" }
            };

            VerticalPosition verticalPosition1 = new VerticalPosition() {
                RelativeFrom = vRelativeFromGood,
                PositionOffset = new PositionOffset { Text = $"{emus}" }
            };

            Assert.Throws<System.InvalidOperationException>(() => document.Paragraphs[0].Image.horizontalPosition = horizontalPosition1);
            Assert.Throws<System.InvalidOperationException>(() => document.Paragraphs[0].Image.verticalPosition = verticalPosition1);

            PositionOffset positionOffsetGood = new PositionOffset { Text = $"{emus}" };
            PositionOffset positionOffsetFail = new PositionOffset { Text = $"{2 * emus}" };
            Assert.NotEqual(positionOffsetFail.Text, positionOffsetGood.Text);

            document.Paragraphs[1].Image.horizontalPosition = horizontalPosition1;
            Assert.Equal(positionOffsetGood.Text, document.Paragraphs[1].Image.horizontalPosition.PositionOffset.Text);
            Assert.NotEqual(positionOffsetFail.Text, document.Paragraphs[1].Image.horizontalPosition.PositionOffset.Text);
            Assert.Equal(hRelativeFromGood, document.Paragraphs[1].Image.horizontalPosition.RelativeFrom.Value);
            Assert.NotEqual(hRelativeFromFail, document.Paragraphs[1].Image.horizontalPosition.RelativeFrom.Value);


            document.Paragraphs[1].Image.verticalPosition = verticalPosition1;
            Assert.Equal(positionOffsetGood.Text, document.Paragraphs[1].Image.verticalPosition.PositionOffset.Text);
            Assert.NotEqual(positionOffsetFail.Text, document.Paragraphs[1].Image.verticalPosition.PositionOffset.Text);
            Assert.Equal(vRelativeFromGood, document.Paragraphs[1].Image.verticalPosition.RelativeFrom.Value);
            Assert.NotEqual(vRelativeFromFail, document.Paragraphs[1].Image.verticalPosition.RelativeFrom.Value);

            document.Save(false);

            Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
        }

        [Fact]
        public void Test_CreatingWordDocumentWithImagesInTable() {
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithImagesInTable.docx");
            using var document = WordDocument.Create(filePath);

            const string fileNameImage = "Kulek.jpg";
            var filePathImage = System.IO.Path.Combine(imagePaths, fileNameImage);

            var table = document.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].AddImage(filePathImage, 200, 200);

            // not really necessary to add new paragraph since one is already there by default
            var paragraph = table.Rows[0].Cells[1].AddParagraph();
            paragraph.AddImage(filePathImage, 200, 200);

            document.AddHeadersAndFooters();

            var tableInHeader = document.Header.Default.AddTable(2, 2);
            tableInHeader.Rows[0].Cells[0].Paragraphs[0].AddImage(filePathImage, 200, 200);

            // not really necessary to add new paragraph since one is already there by default
            var paragraphInHeader = tableInHeader.Rows[0].Cells[1].AddParagraph();
            paragraphInHeader.AddImage(filePathImage, 200, 200);

            Assert.True(document.Tables.Count == 1);
            Assert.True(document.Tables[0].Rows.Count == 2);
            Assert.True(document.Tables[0].Rows[0].Cells.Count == 2);

            Assert.True(document.Header.Default.Tables.Count == 1);

            document.Save(false);

            Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
        }

        [Fact]
        public void Test_CreatingWordDocumentWithImagesInline() {
            var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithImagesInline.docx");
            using var document = WordDocument.Create(filePath);

            var paragraph = document.AddParagraph("This paragraph starts with some text");
            paragraph.Text = "0th This paragraph started with some other text and was overwritten and made bold.";

            // lets add image to paragraph
            var imageParagraph = paragraph.AddImage(Path.Combine(_directoryWithImages, "PrzemyslawKlysAndKulkozaurr.jpg"), 22, 22, WrapTextImage.InLineWithText);

            Assert.True(document.Images[0].WrapText == WrapTextImage.InLineWithText);

            Assert.True(imageParagraph.Image.WrapText == WrapTextImage.InLineWithText);

            imageParagraph.Image.WrapText = WrapTextImage.Square;

            Assert.True(imageParagraph.Image.WrapText == WrapTextImage.Square);

            var paragraph5 = paragraph.AddText("and more text");
            paragraph5.Bold = true;


            document.Save(false);

            Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
        }

        [Fact]
        public void Test_AddImageFromBase64() {
            var filePath = Path.Combine(_directoryWithFiles, "DocumentBase64Image.docx");
            using var document = WordDocument.Create(filePath);

            var bytes = File.ReadAllBytes(Path.Combine(_directoryWithImages, "Kulek.jpg"));
            var base64 = Convert.ToBase64String(bytes);

            var paragraph = document.AddParagraph();
            paragraph.AddImageFromBase64(base64, "Kulek.jpg", 50, 50);

            Assert.Single(document.Images);
            document.Save(false);
        }

        [Fact]
        public void Test_ImageTransparency() {
            var filePath = Path.Combine(_directoryWithFiles, "DocumentImageTransparency.docx");
            using var document = WordDocument.Create(filePath);

            var paragraph = document.AddParagraph();
            paragraph.AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"), 50, 50);
            paragraph.Image.Transparency = 25;

            document.Save(false);

            using (var reloaded = WordDocument.Load(filePath)) {
                Assert.Equal(25, reloaded.Images[0].Transparency);
            }

            using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(filePath, false)) {
                var blip = doc.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().First();
                var alpha = blip.GetFirstChild<DocumentFormat.OpenXml.Drawing.AlphaModulationFixed>();
                Assert.Equal(75000, alpha.Amount.Value);
            }
        }

        [Fact]
        public void Test_ImageTransparency50() {
            var filePath = Path.Combine(_directoryWithFiles, "DocumentImageTransparency50.docx");
            using var document = WordDocument.Create(filePath);

            var paragraph = document.AddParagraph();
            paragraph.AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"), 50, 50);
            paragraph.Image.Transparency = 50;

            document.Save(false);

            using (var reloaded = WordDocument.Load(filePath)) {
                Assert.Equal(50, reloaded.Images[0].Transparency);
            }

            using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(filePath, false)) {
                var blip = doc.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().First();
                var alpha = blip.GetFirstChild<DocumentFormat.OpenXml.Drawing.AlphaModulationFixed>();
                Assert.Equal(50000, alpha.Amount.Value);
            }
        }

        [Fact]
        public void Test_ImageTransparencyNotSet() {
            var filePath = Path.Combine(_directoryWithFiles, "DocumentImageTransparencyNone.docx");
            using var document = WordDocument.Create(filePath);

            var paragraph = document.AddParagraph();
            paragraph.AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"), 50, 50);

            document.Save(false);

            using (var reloaded = WordDocument.Load(filePath)) {
                Assert.Null(reloaded.Images[0].Transparency);
            }

            using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(filePath, false)) {
                var blip = doc.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().First();
                var alpha = blip.GetFirstChild<DocumentFormat.OpenXml.Drawing.AlphaModulationFixed>();
                Assert.Null(alpha);
            }
        }
      
        [Fact]
        public void Test_ImageCropping() {
            var filePath = Path.Combine(_directoryWithFiles, "DocumentImageCrop.docx");
            using (var document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"), 100, 100);

                paragraph.Image.CropTopCentimeters = 1;
                paragraph.Image.CropBottomCentimeters = 2;
                paragraph.Image.CropLeftCentimeters = 3;
                paragraph.Image.CropRightCentimeters = 4;

                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(1, document.Images[0].CropTopCentimeters);
                Assert.Equal(2, document.Images[0].CropBottomCentimeters);
                Assert.Equal(3, document.Images[0].CropLeftCentimeters);
                Assert.Equal(4, document.Images[0].CropRightCentimeters);
            }
        }

        [Fact]
        public void Test_ImageFillModeAndLocalDpi() {
            var filePath = Path.Combine(_directoryWithFiles, "DocumentImageFillMode.docx");
            using (var document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"), 50, 50);
                paragraph.Image.FillMode = ImageFillMode.Tile;
                paragraph.Image.UseLocalDpi = true;
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(ImageFillMode.Tile, document.Images[0].FillMode);
                Assert.True(document.Images[0].UseLocalDpi);
            }

            using (var pkg = WordprocessingDocument.Open(filePath, false)) {
                var blipFill = pkg.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.BlipFill>().First();
                Assert.NotNull(blipFill.GetFirstChild<DocumentFormat.OpenXml.Drawing.Tile>());
                var blip = blipFill.Blip;
                var ext = blip.GetFirstChild<BlipExtensionList>()?.OfType<BlipExtension>().FirstOrDefault(e => e.Uri == "{28A0092B-C50C-407E-A947-70E740481C1C}");
                Assert.NotNull(ext?.GetFirstChild<DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi>());
            }
        }

        [Fact]
        public void Test_AddExternalImage() {
            var filePath = Path.Combine(_directoryWithFiles, "DocumentExternalImage.docx");
            using (var document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddImage(new Uri("http://example.com/image.png"), 50, 50);
                Assert.Single(document.Images);
                Assert.True(document.Images[0].IsExternal);
                Assert.Equal(new Uri("http://example.com/image.png"), document.Images[0].ExternalUri);
                Assert.Throws<InvalidOperationException>(() => document.Images[0].SaveToFile("tmp.png"));
                document.Images[0].Remove();
                Assert.Empty(document.Images);
                document.Save(false);
            }

            using (var pkg = WordprocessingDocument.Open(filePath, false)) {
                // ensure document opens correctly after removing the external image
                Assert.NotNull(pkg.MainDocumentPart.Document);
            }
        }

        [Fact]
        public void Test_ImageNonVisualProperties() {
            var filePath = Path.Combine(_directoryWithFiles, "DocumentImageNvProps.docx");
            using (var document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"), 50, 50);
                paragraph.Image.Title = "MyTitle";
                paragraph.Image.Hidden = true;
                paragraph.Image.PreferRelativeResize = true;
                paragraph.Image.NoChangeAspect = true;
                paragraph.Image.FixedOpacity = 80;
                document.Save(false);
            }

            using (var pkg = WordprocessingDocument.Open(filePath, false)) {
                var pic = pkg.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().First();
                var nv = pic.NonVisualPictureProperties;
                Assert.Equal("MyTitle", nv.NonVisualDrawingProperties.Title);
                Assert.True(nv.NonVisualDrawingProperties.Hidden);
                Assert.True(nv.NonVisualPictureDrawingProperties.PreferRelativeResize);
                Assert.True(nv.NonVisualPictureDrawingProperties.PictureLocks.NoChangeAspect);
                var ar = pic.BlipFill.Blip.GetFirstChild<DocumentFormat.OpenXml.Drawing.AlphaReplace>();
                Assert.Equal(80000, ar.Alpha.Value);
            }
        }

        [Fact]
        public void Test_ImageEffects() {
            var filePath = Path.Combine(_directoryWithFiles, "DocumentImageEffects.docx");
            using (var document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"), 50, 50);
                var img = paragraph.Image;
                img.AlphaInversionColor = SixLabors.ImageSharp.Color.Red;
                img.BlackWhiteThreshold = 60;
                img.BlurRadius = 2000;
                img.BlurGrow = true;
                img.ColorChangeFrom = SixLabors.ImageSharp.Color.Parse("#97E4FE");
                img.ColorChangeTo = SixLabors.ImageSharp.Color.Parse("#FF3399");
                img.ColorReplacement = SixLabors.ImageSharp.Color.Lime;
                img.DuotoneColor1 = SixLabors.ImageSharp.Color.Black;
                img.DuotoneColor2 = SixLabors.ImageSharp.Color.White;
                img.GrayScale = true;
                img.LuminanceBrightness = 65;
                img.LuminanceContrast = 30;
                img.TintAmount = 50;
                img.TintHue = 300;
                document.Save(false);
            }

            using (var reloaded = WordDocument.Load(filePath)) {
                var img = reloaded.Images[0];
                Assert.Equal("ff0000", img.AlphaInversionColorHex);
                Assert.Equal(60, img.BlackWhiteThreshold);
                Assert.Equal(2000, img.BlurRadius);
                Assert.True(img.BlurGrow);
                Assert.Equal("97e4fe", img.ColorChangeFromHex);
                Assert.Equal("ff3399", img.ColorChangeToHex);
                Assert.Equal("00ff00", img.ColorReplacementHex);
                Assert.Equal("000000", img.DuotoneColor1Hex);
                Assert.Equal("ffffff", img.DuotoneColor2Hex);
                Assert.True(img.GrayScale);
                Assert.Equal(65, img.LuminanceBrightness);
                Assert.Equal(30, img.LuminanceContrast);
                Assert.Equal(50, img.TintAmount);
                Assert.Equal(300, img.TintHue);
            }
        }

    }

}
