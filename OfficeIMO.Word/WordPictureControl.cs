using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a picture content control within a paragraph.
    /// </summary>
    public class WordPictureControl : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        internal readonly SdtRun _sdtRun;

        internal WordPictureControl(WordDocument document, Paragraph paragraph, SdtRun sdtRun) {
            _document = document;
            _paragraph = paragraph;
            _sdtRun = sdtRun;
        }

        /// <summary>
        /// Gets the alias associated with this picture control.
        /// </summary>
        public string? Alias {
            get {
                var properties = _sdtRun.SdtProperties;
                var sdtAlias = properties?.OfType<SdtAlias>().FirstOrDefault();
                return sdtAlias?.Val;
            }
        }

        /// <summary>
        /// Gets or sets the tag value for this picture control.
        /// </summary>
        public string? Tag {
            get {
                var properties = _sdtRun.SdtProperties;
                var tag = properties?.OfType<Tag>().FirstOrDefault();
                return tag?.Val;
            }
            set {
                var properties = EnsureProperties();
                var tag = properties.OfType<Tag>().FirstOrDefault();
                if (value == null) {
                    tag?.Remove();
                    return;
                }
                if (tag == null) {
                    tag = new Tag();
                    properties.Append(tag);
                }
                tag.Val = value;
            }
        }

        /// <summary>
        /// Gets the image currently contained by the picture content control.
        /// </summary>
        public WordImage? Image {
            get {
                WordDrawing? drawing = _sdtRun.SdtContentRun?.Descendants<WordDrawing>().FirstOrDefault();
                return drawing == null ? null : new WordImage(_document, drawing);
            }
        }

        /// <summary>
        /// Replaces the picture content-control image with the image at the supplied path.
        /// </summary>
        /// <param name="filePath">Image file path.</param>
        /// <param name="width">Optional width of the image.</param>
        /// <param name="height">Optional height of the image.</param>
        public void SetImage(string filePath, double? width = null, double? height = null) {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));

            var newRun = new Run();
            var paragraph = new WordParagraph(_document, _paragraph, newRun);
            paragraph.AddImage(filePath, width, height);

            ReplaceContent(newRun);
        }

        /// <summary>
        /// Replaces the picture content-control image with image data from a stream.
        /// </summary>
        /// <param name="imageStream">Image stream.</param>
        /// <param name="fileName">File name used to infer image type.</param>
        /// <param name="width">Optional width of the image.</param>
        /// <param name="height">Optional height of the image.</param>
        public void SetImage(Stream imageStream, string fileName, double? width = null, double? height = null) {
            if (imageStream == null) throw new ArgumentNullException(nameof(imageStream));
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentNullException(nameof(fileName));

            var newRun = new Run();
            var paragraph = new WordParagraph(_document, _paragraph, newRun);
            paragraph.AddImage(imageStream, fileName, width, height);

            ReplaceContent(newRun);
        }

        /// <summary>
        /// Extracts the current picture content-control image as a form-map value.
        /// </summary>
        public WordContentControlPictureValue? ExtractValue() {
            WordImage? image = Image;
            if (image == null) {
                return null;
            }

            if (image.IsExternal) {
                return WordContentControlPictureValue.FromExternalImage(image.ExternalUri, image.FileName, image.ExternalRelationshipId);
            }

            return WordContentControlPictureValue.FromBytes(image.GetBytes(), image.FileName ?? "image.bin");
        }

        /// <summary>
        /// Removes the picture control from the paragraph.
        /// </summary>
        public void Remove() {
            _sdtRun.Remove();
        }

        private void ReplaceContent(Run run) {
            Image?.Remove();
            _sdtRun.SdtContentRun ??= new SdtContentRun();
            _sdtRun.SdtContentRun.RemoveAllChildren();
            _sdtRun.SdtContentRun.Append(run);
        }

        private SdtProperties EnsureProperties() {
            if (_sdtRun.SdtProperties == null) {
                _sdtRun.SdtProperties = new SdtProperties();
            }
            return _sdtRun.SdtProperties;
        }
    }
}
