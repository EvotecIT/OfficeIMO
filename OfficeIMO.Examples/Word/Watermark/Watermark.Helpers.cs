using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
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

            WordHeader? header;
            string description;

            if (type == HeaderFooterValues.Default) {
                header = section.Header.Default;
                description = "default";
            } else if (type == HeaderFooterValues.Even) {
                header = section.Header.Even;
                description = "even";
            } else if (type == HeaderFooterValues.First) {
                header = section.Header.First;
                description = "first";
            } else {
                throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported header type.");
            }

            return Guard.NotNull(header, $"Call AddHeadersAndFooters before accessing the {description} header.");
        }
    }
}
