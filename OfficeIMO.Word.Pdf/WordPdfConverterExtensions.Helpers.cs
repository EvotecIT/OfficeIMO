using OfficeIMO.Word;
using QuestPDF.Helpers;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static PageSize MapToPageSize(WordPageSize pageSize) {
            return pageSize switch {
                WordPageSize.Letter => PageSizes.Letter,
                WordPageSize.Legal => PageSizes.Legal,
                WordPageSize.Executive => PageSizes.Executive,
                WordPageSize.A3 => PageSizes.A3,
                WordPageSize.A4 => PageSizes.A4,
                WordPageSize.A5 => PageSizes.A5,
                WordPageSize.A6 => PageSizes.A6,
                WordPageSize.B5 => PageSizes.B5,
                _ => PageSizes.A4
            };
        }
    }
}
