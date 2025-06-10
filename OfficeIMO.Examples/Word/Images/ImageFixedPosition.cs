using System;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_AddingFixedImages(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with an Image in a fixed position.");
            var filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithImages.docx");
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

            using var document = WordDocument.Create(filePath);
            document.BuiltinDocumentProperties.Title = "Fixed image example";
            document.BuiltinDocumentProperties.Creator = "example";

            var paragraph1 = document.AddParagraph("First Paragraph");

            const string fileNameImage = "Kulek.jpg";
            var filePathImage = System.IO.Path.Combine(imagePaths, fileNameImage);
            // Add an image with a fixed position to paragraph. First we add the image, then we will
            // edit the position properties.
            //
            // Note: The image MUST be constructed with a WrapTextImage property that is NOT inline. Assigning
            // the WrapTextImage property later was not available at the time of making this example.
            paragraph1.AddImage(filePathImage, 100, 100, WrapTextImage.Square);

            Console.WriteLine("PRE position edit.");
            // Before editing, we can assess the RelativeFrom and PositionOffset properties of the image.
            //DocumentFormat.OpenXml.EnumValue<HorizontalRelativePositionValues> hRelativeFrom;
            //string hOffset, vOffset;
            //DocumentFormat.OpenXml.EnumValue<VerticalRelativePositionValues> vRelativeFrom;
            checkImageProps(paragraph1);

            // Begin editing the fixed position properties of the image. You may edit both, however it
            // is not necessary.

            // Note that the units for the PositionOffset are taken in EMU's. This is a conversion
            // for an offset of 1/4 inch.
            const double emusPerInch = 914400.0;
            double offsetInches = 0.25;
            // Non integer values will cause the document properties to be corrupted, cast
            // to an int for avoiding this.
            int offsetEmus = (int)(offsetInches * emusPerInch);

            // Edit the horizontal relative from property of the image. Both
            // the RelativeFrom property and PositionOffset are required.
            HorizontalPosition horizontalPosition1 = new HorizontalPosition() {
                RelativeFrom = HorizontalRelativePositionValues.Page,
                PositionOffset = new PositionOffset { Text = $"{offsetEmus}" }
            };
            paragraph1.Image.horizontalPosition = horizontalPosition1;

            // Edit the vertical relative from property of the image. Both
            // the RelativeFrom property and PositionOffset are required.
            VerticalPosition verticalPosition1 = new VerticalPosition() {
                RelativeFrom = VerticalRelativePositionValues.Page,
                PositionOffset = new PositionOffset { Text = $"{offsetEmus}" }
            };
            paragraph1.Image.verticalPosition = verticalPosition1;

            Console.WriteLine("POST position edit.");
            // After editing, lets reassess the properties.
            checkImageProps(paragraph1);

            // This will put the image in the upper top left corner of the document.

            document.Save(openWord);

            static void checkImageProps(WordParagraph paragraph1) {
                var hRelativeFrom = paragraph1.Image.horizontalPosition.RelativeFrom;
                var hOffset = paragraph1.Image.horizontalPosition.PositionOffset.Text;
                var vRelativeFrom = paragraph1.Image.verticalPosition.RelativeFrom;
                var vOffset = paragraph1.Image.verticalPosition.PositionOffset.Text;
                Console.WriteLine($"Horizontal RelativeFrom type: {hRelativeFrom.ToString()}");
                Console.WriteLine($"Horizontal PositionOffset value: {hOffset.ToString()}");
                Console.WriteLine($"Vertical RelativeFrom type: {vRelativeFrom.ToString()}");
                Console.WriteLine($"Vertical PositionOffset value: {vOffset.ToString()}");
            }
        }
    }
}
