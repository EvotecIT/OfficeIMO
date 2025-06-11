using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;

namespace OfficeIMO.Word {
    public class WordEmbeddedObject {
        private readonly WordDocument _document;
        private readonly Run _run;

        public WordEmbeddedObject(WordParagraph wordParagraph, WordDocument wordDocument, string fileName, string fileImage, string description, double? width = null, double? height = null) {


            _document = wordDocument;

            width ??= 64.8;
            height ??= 40.8;

            var embeddedObject = ConvertFileToEmbeddedObject(wordDocument, fileName, fileImage, width.Value, height.Value);

            Run run = new Run();
            run.Append(embeddedObject);
            wordParagraph._paragraph.AppendChild(run);

            _run = run;

            //var p = GenerateParagraph(idImagePart, idEmbeddedObjectPart);

            //wordDocument._document.MainDocumentPart.Document.Body.AppendChild(p);
        }

        internal WordEmbeddedObject(WordDocument wordDocument, Run run) {
            _document = wordDocument;
            _run = run;
        }

        //public Paragraph GenerateParagraph(string imageId, string embedId) {
        //    Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "008F6FFA", RsidRunAdditionDefault = "008F6FFA", ParagraphId = "324F144F", TextId = "77777777" };

        //    Run run1 = new Run();



        //    run1.Append(embeddedObject1);

        //    paragraph1.Append(run1);
        //    return paragraph1;
        //}

        private (string contentType, string programId) GetObjectInfo(string fileName) {
            string extension = System.IO.Path.GetExtension(fileName).ToLower();
            return extension switch {
                ".xlsx" => ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Excel.Sheet.12"),
                ".xls"  => ("application/vnd.ms-excel", "Excel.Sheet.8"),
                ".docx" => ("application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Word.Document.12"),
                ".doc"  => ("application/msword", "Word.Document.8"),
                ".pptx" => ("application/vnd.openxmlformats-officedocument.presentationml.presentation", "PowerPoint.Show.12"),
                ".ppt"  => ("application/vnd.ms-powerpoint", "PowerPoint.Show.8"),
                ".pdf"  => ("application/pdf", "AcroExch.Document.DC"),
                ".html" => ("text/html", "htmlfile"),
                ".htm"  => ("text/html", "htmlfile"),
                ".rtf"  => ("application/rtf", "Word.RTF.8"),
                _       => ("application/octet-stream", "Package")
            };
        }

        private EmbeddedObject ConvertFileToEmbeddedObject(WordDocument wordDocument, string fileName, string fileImage, double width, double height) {
            ImagePart imagePart = wordDocument._document.MainDocumentPart.AddImagePart(ImagePartType.Png);
            using (FileStream stream = new FileStream(fileImage, FileMode.Open)) {
                imagePart.FeedData(stream);
            }
            MainDocumentPart mainPart = wordDocument._document.MainDocumentPart;

            var (contentType, programId) = GetObjectInfo(fileName);
            //ProgId = "Package",
            //var contentType = "application/vnd.openxmlformats-officedocument.oleObject";
            //var programId = "Package";


            EmbeddedPackagePart embeddedObjectPart = mainPart.AddEmbeddedPackagePart(contentType);

            using (FileStream fileStream = new FileStream(fileName, FileMode.Open)) {
                embeddedObjectPart.FeedData(fileStream);
            }

            var idImagePart = mainPart.GetIdOfPart(imagePart);
            var idEmbeddedObjectPart = mainPart.GetIdOfPart(embeddedObjectPart);

            var embeddedObject = CreateEmbeddedObject(idImagePart, idEmbeddedObjectPart, programId, width, height);
            //var embeddedObject = GenerateEmbeddedObject(idImagePart, idEmbeddedObjectPart, programId, 49.2, 49.2);
            return embeddedObject;
        }


        private EmbeddedObject CreateEmbeddedObject(string imageId, string packageEmbedId, string programId, double width, double height) {
            EmbeddedObject embeddedObject1 = new EmbeddedObject() {
                DxaOriginal = "15962",
                DyaOriginal = "21179",
                AnchorId = "3C42CF0C"
            };

            V.Shapetype shapetype1 = new V.Shapetype() {
                Id = "_x0000_t75",
                CoordinateSize = "21600,21600",
                Filled = false,
                Stroked = false,
                OptionalNumber = 75,
                PreferRelative = true,
                EdgePath = "m@4@5l@4@11@9@11@9@5xe"
            };
            V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };

            V.Formulas formulas1 = new V.Formulas();
            V.Formula formula1 = new V.Formula() { Equation = "if lineDrawn pixelLineWidth 0" };
            V.Formula formula2 = new V.Formula() { Equation = "sum @0 1 0" };
            V.Formula formula3 = new V.Formula() { Equation = "sum 0 0 @1" };
            V.Formula formula4 = new V.Formula() { Equation = "prod @2 1 2" };
            V.Formula formula5 = new V.Formula() { Equation = "prod @3 21600 pixelWidth" };
            V.Formula formula6 = new V.Formula() { Equation = "prod @3 21600 pixelHeight" };
            V.Formula formula7 = new V.Formula() { Equation = "sum @0 0 1" };
            V.Formula formula8 = new V.Formula() { Equation = "prod @6 1 2" };
            V.Formula formula9 = new V.Formula() { Equation = "prod @7 21600 pixelWidth" };
            V.Formula formula10 = new V.Formula() { Equation = "sum @8 21600 0" };
            V.Formula formula11 = new V.Formula() { Equation = "prod @7 21600 pixelHeight" };
            V.Formula formula12 = new V.Formula() { Equation = "sum @10 21600 0" };

            formulas1.Append(formula1);
            formulas1.Append(formula2);
            formulas1.Append(formula3);
            formulas1.Append(formula4);
            formulas1.Append(formula5);
            formulas1.Append(formula6);
            formulas1.Append(formula7);
            formulas1.Append(formula8);
            formulas1.Append(formula9);
            formulas1.Append(formula10);
            formulas1.Append(formula11);
            formulas1.Append(formula12);

            V.Path path1 = new V.Path() {
                AllowGradientShape = true,
                ConnectionPointType = Ovml.ConnectValues.Rectangle,
                AllowExtrusion = false
            };
            Ovml.Lock lock1 = new Ovml.Lock() {
                Extension = V.ExtensionHandlingBehaviorValues.Edit,
                AspectRatio = true
            };

            shapetype1.Append(stroke1);
            shapetype1.Append(formulas1);
            shapetype1.Append(path1);
            shapetype1.Append(lock1);

            var style = "width:" + width + "pt;height:" + height + "pt";

            V.Shape shape1 = new V.Shape() {
                Id = "_x0000_i1029",
                Style = style,
                //Style = "width:798pt;height:1059pt",
                Ole = false,
                Type = "#_x0000_t75"
            };

            V.ImageData imageData1 = new V.ImageData() {
                Title = "",
                RelationshipId = imageId
            };

            shape1.Append(imageData1);

            Ovml.OleObject oleObject1 = new Ovml.OleObject() {
                Type = Ovml.OleValues.Embed,
                ProgId = programId,
                ShapeId = "_x0000_i1029",
                DrawAspect = Ovml.OleDrawAspectValues.Content,
                ObjectId = "_" + Guid.NewGuid().ToString("N"),
                Id = packageEmbedId
            };


            embeddedObject1.Append(shapetype1);
            embeddedObject1.Append(shape1);
            embeddedObject1.Append(oleObject1);
            return embeddedObject1;
        }

        public EmbeddedObject GenerateEmbeddedObject(string imageId, string packageEmbedId, string programId, double width, double height) {
            EmbeddedObject embeddedObject1 = new EmbeddedObject() { DxaOriginal = "1297", DyaOriginal = "816", AnchorId = "595268A8" };

            V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t75", CoordinateSize = "21600,21600", Filled = false, Stroked = false, OptionalNumber = 75, PreferRelative = true, EdgePath = "m@4@5l@4@11@9@11@9@5xe" };
            V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };

            V.Formulas formulas1 = new V.Formulas();
            V.Formula formula1 = new V.Formula() { Equation = "if lineDrawn pixelLineWidth 0" };
            V.Formula formula2 = new V.Formula() { Equation = "sum @0 1 0" };
            V.Formula formula3 = new V.Formula() { Equation = "sum 0 0 @1" };
            V.Formula formula4 = new V.Formula() { Equation = "prod @2 1 2" };
            V.Formula formula5 = new V.Formula() { Equation = "prod @3 21600 pixelWidth" };
            V.Formula formula6 = new V.Formula() { Equation = "prod @3 21600 pixelHeight" };
            V.Formula formula7 = new V.Formula() { Equation = "sum @0 0 1" };
            V.Formula formula8 = new V.Formula() { Equation = "prod @6 1 2" };
            V.Formula formula9 = new V.Formula() { Equation = "prod @7 21600 pixelWidth" };
            V.Formula formula10 = new V.Formula() { Equation = "sum @8 21600 0" };
            V.Formula formula11 = new V.Formula() { Equation = "prod @7 21600 pixelHeight" };
            V.Formula formula12 = new V.Formula() { Equation = "sum @10 21600 0" };

            formulas1.Append(formula1);
            formulas1.Append(formula2);
            formulas1.Append(formula3);
            formulas1.Append(formula4);
            formulas1.Append(formula5);
            formulas1.Append(formula6);
            formulas1.Append(formula7);
            formulas1.Append(formula8);
            formulas1.Append(formula9);
            formulas1.Append(formula10);
            formulas1.Append(formula11);
            formulas1.Append(formula12);
            V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle, AllowExtrusion = false };
            Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, AspectRatio = true };

            shapetype1.Append(stroke1);
            shapetype1.Append(formulas1);
            shapetype1.Append(path1);
            shapetype1.Append(lock1);

            V.Shape shape1 = new V.Shape() { Id = "_x0000_i1025", Style = "width:64.8pt;height:40.8pt", Ole = false, Type = "#_x0000_t75" };
            V.ImageData imageData1 = new V.ImageData() { Title = "", RelationshipId = imageId };

            shape1.Append(imageData1);
            Ovml.OleObject oleObject1 = new Ovml.OleObject() { Type = Ovml.OleValues.Embed, ProgId = "Package", ShapeId = "_x0000_i1025", DrawAspect = Ovml.OleDrawAspectValues.Content, ObjectId = "_1736440255", Id = packageEmbedId };

            embeddedObject1.Append(shapetype1);
            embeddedObject1.Append(shape1);
            embeddedObject1.Append(oleObject1);
            return embeddedObject1;
        }
    }
}
