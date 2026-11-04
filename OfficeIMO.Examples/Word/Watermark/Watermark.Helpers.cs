using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Watermark {
        private static WordHeader GetRequiredHeader(WordSection section) {
            return GetRequiredHeader(section, HeaderFooterValues.Default);
        }

        private static WordHeader GetRequiredHeader(WordSection section, HeaderFooterValues type) {
            ArgumentNullException.ThrowIfNull(section);

            var setup = section.PageSetup;
            string description;
            WordHeader? header;

            if (type == HeaderFooterValues.Default) {
                description = "default";
                header = setup.GetHeader(WordHeaderFooterType.Default);
            } else if (type == HeaderFooterValues.Even) {
                description = "even";
                header = setup.GetHeader(WordHeaderFooterType.Even);
            } else if (type == HeaderFooterValues.First) {
                description = "first";
                header = setup.GetHeader(WordHeaderFooterType.First);
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
