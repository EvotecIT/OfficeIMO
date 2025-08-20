using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for images.
    /// </summary>
    public class ImageBuilder {
        private readonly WordFluentDocument _fluent;
        private WordImage? _image;
        private WordParagraph? _paragraph;

        internal ImageBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        public WordImage? Image => _image;

        public ImageBuilder Add(string path) {
            var paragraph = _fluent.Document.AddParagraph();
            paragraph.AddImage(path);
            _image = paragraph.Image;
            _paragraph = paragraph;
            return this;
        }

        public ImageBuilder Add(Stream stream, string fileName) {
            var paragraph = _fluent.Document.AddParagraph();
            paragraph.AddImage(stream, fileName, null, null);
            _image = paragraph.Image;
            _paragraph = paragraph;
            return this;
        }

        public ImageBuilder Add(byte[] bytes, string fileName) {
            using var ms = new MemoryStream(bytes);
            return Add(ms, fileName);
        }

        public ImageBuilder AddFromUrl(string url) {
            using HttpClient client = new HttpClient();
            var data = client.GetByteArrayAsync(url).GetAwaiter().GetResult();
            string fileName = GetFileName(url);
            return Add(data, fileName);
        }

        public async Task<ImageBuilder> AddFromUrlAsync(string url) {
            using HttpClient client = new HttpClient();
            var data = await client.GetByteArrayAsync(url);
            string fileName = GetFileName(url);
            return Add(data, fileName);
        }

        public ImageBuilder Size(double width, double? height = null) {
            if (_image != null) {
                _image.Width = width;
                if (height != null) {
                    _image.Height = height.Value;
                }
            }
            return this;
        }

        public ImageBuilder MaxWidth(double width) {
            if (_image != null) {
                if (_image.Width == null || _image.Width > width) {
                    _image.Width = width;
                }
            }
            return this;
        }

        public ImageBuilder Wrap(WrapTextImage wrapImage) {
            if (_image != null) {
                _image.WrapText = wrapImage;
            }
            return this;
        }

        public ImageBuilder Align(HorizontalAlignment alignment) {
            var justification = alignment switch {
                HorizontalAlignment.Center => JustificationValues.Center,
                HorizontalAlignment.Right => JustificationValues.Right,
                HorizontalAlignment.Justified => JustificationValues.Both,
                _ => JustificationValues.Left,
            };
            _paragraph?.SetAlignment(justification);
            return this;
        }

        private static string GetFileName(string url) {
            try {
                var uri = new Uri(url);
                var fileName = Path.GetFileName(uri.LocalPath);
                return string.IsNullOrEmpty(fileName) ? "image" : fileName;
            } catch (UriFormatException) {
                return "image";
            }
        }
    }
}
