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
        /// Add image from file with ability to provide width and height of the image
        /// The image will be resized given new dimensions
        /// </summary>
        /// <param name="filePathImage">Path to file to import to Word Document</param>
        /// <param name="width">Optional width of the image. If not given the actual image width will be used.</param>
        /// <param name="height">Optional height of the image. If not given the actual image height will be used.</param>
        /// <param name="wrapImageText">Optional text wrapping rule. If not given the image will be inserted inline to the text.</param>
        /// <param name="description">The description for this image.</param>
        /// <returns>The WordParagraph that AddImage was called on.</returns>
        public WordParagraph AddImage(string filePathImage, double? width = null, double? height = null, WrapTextImage wrapImageText = WrapTextImage.InLineWithText, string description = "") {
            var wordImage = new WordImage(_document, this, filePathImage, width, height, wrapImageText, description);
            var run = VerifyRun();
            run.Append(wordImage._Image);
            return this;
        }

        /// <summary>
        /// Inserts an image and returns the created <see cref="WordImage"/> for immediate configuration.
        /// This is a convenience alternative to <see cref="AddImage(string,double?,double?,WrapTextImage,string)"/>
        /// when you want to set properties on the inserted image without accessing <see cref="Image"/>.
        /// </summary>
        public WordImage InsertImage(string filePathImage, double? width = null, double? height = null,
            WrapTextImage wrapImageText = WrapTextImage.InLineWithText, string description = "") {
            var wordImage = new WordImage(_document, this, filePathImage, width, height, wrapImageText, description);
            var run = VerifyRun();
            run.Append(wordImage._Image);
            return wordImage;
        }
        /// <summary>
        /// Add image from Stream with ability to provide width and height of the image
        /// The image will be resized given new dimensions
        /// </summary>
        /// <param name="imageStream">The stream to load the image from.</param>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="width">Optional width of the image. If not given the actual image width will be used.</param>
        /// <param name="height">Optional height of the image. If not give the actual image height will be used.</param>
        /// <param name="wrapImageText">Optional text wrapping rule. If not given the image will be inserted inline to the text.</param>
        /// <param name="description">The description for this image.</param>
        /// <returns>The WordParagraph that AddImage was called on.</returns>
        public WordParagraph AddImage(Stream imageStream, string fileName, double? width, double? height, WrapTextImage wrapImageText = WrapTextImage.InLineWithText, string description = "") {
            var wordImage = new WordImage(_document, this, imageStream, fileName, width, height, wrapImageText, description);
            var run = VerifyRun();
            run.Append(wordImage._Image);
            return this;
        }

        /// <summary>
        /// Inserts an image from a stream and returns the created <see cref="WordImage"/>.
        /// </summary>
        public WordImage InsertImage(Stream imageStream, string fileName, double? width = null, double? height = null,
            WrapTextImage wrapImageText = WrapTextImage.InLineWithText, string description = "") {
            var wordImage = new WordImage(_document, this, imageStream, fileName, width, height, wrapImageText, description);
            var run = VerifyRun();
            run.Append(wordImage._Image);
            return wordImage;
        }

        /// <summary>
        /// Add an image that is stored outside the package.
        /// </summary>
        public WordParagraph AddImage(Uri imageUri, double width, double height, WrapTextImage wrapImageText = WrapTextImage.InLineWithText, string description = "") {
            var wordImage = new WordImage(_document, this, imageUri, width, height, wrapImageText, description);
            var run = VerifyRun();
            run.Append(wordImage._Image);
            return this;
        }

        /// <summary>
        /// Inserts an external image (by URI) and returns the created <see cref="WordImage"/>.
        /// </summary>
        public WordImage InsertImage(Uri imageUri, double width, double height,
            WrapTextImage wrapImageText = WrapTextImage.InLineWithText, string description = "") {
            var wordImage = new WordImage(_document, this, imageUri, width, height, wrapImageText, description);
            var run = VerifyRun();
            run.Append(wordImage._Image);
            return wordImage;
        }

        /// <summary>
        /// Add image from a Base64 encoded string.
        /// </summary>
        public WordParagraph AddImageFromBase64(string base64String, string fileName, double? width = null, double? height = null, WrapTextImage wrapImageText = WrapTextImage.InLineWithText, string description = "") {
            var wordImage = new WordImage(_document, this, base64String, fileName, width, height, wrapImageText, description);
            var run = VerifyRun();
            run.Append(wordImage._Image);
            return this;
        }

        /// <summary>
        /// Inserts an image from a Base64 payload and returns the created <see cref="WordImage"/>.
        /// </summary>
        public WordImage InsertImageFromBase64(string base64String, string fileName, double? width = null, double? height = null,
            WrapTextImage wrapImageText = WrapTextImage.InLineWithText, string description = "") {
            var wordImage = new WordImage(_document, this, base64String, fileName, width, height, wrapImageText, description);
            var run = VerifyRun();
            run.Append(wordImage._Image);
            return wordImage;
        }

        /// <summary>
        /// Add image from an embedded resource.
        /// </summary>
        /// <param name="assembly">Assembly that contains the resource.</param>
        /// <param name="resourceName">Full name of the embedded resource.</param>
        /// <param name="width">Optional width of the image.</param>
        /// <param name="height">Optional height of the image.</param>
        /// <param name="wrapImageText">Optional text wrapping rule.</param>
        /// <param name="description">The description for this image.</param>
        /// <returns>The WordParagraph that AddImage was called on.</returns>
        public WordParagraph AddImageFromResource(Assembly assembly, string resourceName, double? width = null, double? height = null, WrapTextImage wrapImageText = WrapTextImage.InLineWithText, string description = "") {
            assembly ??= Assembly.GetCallingAssembly();
            var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null) {
                throw new ArgumentException($"Resource '{resourceName}' was not found in assembly '{assembly.FullName}'.", nameof(resourceName));
            }
            using (stream) {
                var fileName = Path.GetFileName(resourceName);
                var wordImage = new WordImage(_document, this, stream, fileName, width, height, wrapImageText, description);
                var run = VerifyRun();
                run.Append(wordImage._Image);
            }
            return this;
        }
    }
}