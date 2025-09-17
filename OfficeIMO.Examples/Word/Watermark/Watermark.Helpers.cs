using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Watermark {
        private static WordHeader GetRequiredHeader(WordSection section) {
            return GetRequiredHeader(section, HeaderFooterValues.Default);
        }

        private static WordHeader GetRequiredHeader(WordSection section, HeaderFooterValues type) {
            if (section is null) {
                throw new ArgumentNullException(nameof(section));
            }

            if (section.Header == null) {
                section.AddHeadersAndFooters();
            }

            var headers = section.Header;
            if (headers == null) {
                throw new InvalidOperationException("Headers are not available after attempting to add them.");
            }

            string description;
            WordHeader? header;

            if (type == HeaderFooterValues.Default) {
                description = "default";
                header = headers.Default;
            } else if (type == HeaderFooterValues.Even) {
                description = "even";
                header = headers.Even;
            } else if (type == HeaderFooterValues.First) {
                description = "first";
                header = headers.First;
            } else {
                throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported header type.");
            }

            if (header == null) {
                throw new InvalidOperationException($"The {description} header is not available.");
            }

            return header;
        }
    }
}
