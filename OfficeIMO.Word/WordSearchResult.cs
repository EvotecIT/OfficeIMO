using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Internal;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    /// <summary>
    /// Result of <see cref="WordDocument.Find(string, OfficeTextSearchOptions?)"/> and <see cref="WordDocument.Find(Regex)"/>:
    /// the paragraphs containing matches, grouped by document part, and the total match count.
    /// </summary>
    public sealed class WordSearchResult {
        private readonly List<WordParagraph> _bodyParagraphs = new();
        private readonly List<WordParagraph> _tableParagraphs = new();
        private readonly List<WordParagraph> _defaultHeaderParagraphs = new();
        private readonly List<WordParagraph> _evenPageHeaderParagraphs = new();
        private readonly List<WordParagraph> _firstPageHeaderParagraphs = new();
        private readonly List<WordParagraph> _defaultFooterParagraphs = new();
        private readonly List<WordParagraph> _evenPageFooterParagraphs = new();
        private readonly List<WordParagraph> _firstPageFooterParagraphs = new();
        private readonly List<WordParagraph> _footnoteParagraphs = new();
        private readonly List<WordParagraph> _endnoteParagraphs = new();

        private readonly Func<IReadOnlyList<WordRange>> _rangeFactory;
        private IReadOnlyList<WordRange>? _ranges;

        internal WordSearchResult(Func<IReadOnlyList<WordRange>> rangeFactory) {
            _rangeFactory = rangeFactory;
        }

        /// <summary>
        /// Every match as an editable <see cref="WordRange"/>, in document order. The ranges are located when this property is first read.
        /// </summary>
        /// <remarks>
        /// Use a range to read, replace, format, bookmark, comment or link the matched text. Replacing a range moves the
        /// ranges after it, so the list can be processed in order.
        /// </remarks>
        public IReadOnlyList<WordRange> Ranges => _ranges ??= _rangeFactory();

        private IReadOnlyList<WordParagraph>? _paragraphs;

        /// <summary>Gets the total number of matches.</summary>
        public int MatchCount { get; internal set; }

        /// <summary>
        /// Every paragraph containing one or more matches, in the order of the body, table, header, footer, footnote, and endnote lists.
        /// </summary>
        public IReadOnlyList<WordParagraph> Paragraphs => _paragraphs ??= _bodyParagraphs
            .Concat(_tableParagraphs)
            .Concat(_defaultHeaderParagraphs)
            .Concat(_evenPageHeaderParagraphs)
            .Concat(_firstPageHeaderParagraphs)
            .Concat(_defaultFooterParagraphs)
            .Concat(_evenPageFooterParagraphs)
            .Concat(_firstPageFooterParagraphs)
            .Concat(_footnoteParagraphs)
            .Concat(_endnoteParagraphs)
            .ToList();

        /// <summary>Body paragraphs containing one or more matches.</summary>
        public IReadOnlyList<WordParagraph> BodyParagraphs => _bodyParagraphs;
        /// <summary>Table-cell paragraphs containing one or more matches.</summary>
        public IReadOnlyList<WordParagraph> TableParagraphs => _tableParagraphs;
        /// <summary>Default-header paragraphs containing one or more matches.</summary>
        public IReadOnlyList<WordParagraph> DefaultHeaderParagraphs => _defaultHeaderParagraphs;
        /// <summary>Even-page-header paragraphs containing one or more matches.</summary>
        public IReadOnlyList<WordParagraph> EvenPageHeaderParagraphs => _evenPageHeaderParagraphs;
        /// <summary>First-page-header paragraphs containing one or more matches.</summary>
        public IReadOnlyList<WordParagraph> FirstPageHeaderParagraphs => _firstPageHeaderParagraphs;
        /// <summary>Default-footer paragraphs containing one or more matches.</summary>
        public IReadOnlyList<WordParagraph> DefaultFooterParagraphs => _defaultFooterParagraphs;
        /// <summary>Even-page-footer paragraphs containing one or more matches.</summary>
        public IReadOnlyList<WordParagraph> EvenPageFooterParagraphs => _evenPageFooterParagraphs;
        /// <summary>First-page-footer paragraphs containing one or more matches.</summary>
        public IReadOnlyList<WordParagraph> FirstPageFooterParagraphs => _firstPageFooterParagraphs;
        /// <summary>Ordinary footnote paragraphs containing one or more matches from a document-level search.</summary>
        public IReadOnlyList<WordParagraph> FootnoteParagraphs => _footnoteParagraphs;
        /// <summary>Ordinary endnote paragraphs containing one or more matches from a document-level search.</summary>
        public IReadOnlyList<WordParagraph> EndnoteParagraphs => _endnoteParagraphs;

        internal List<WordParagraph> MutableBodyParagraphs => _bodyParagraphs;
        internal List<WordParagraph> MutableTableParagraphs => _tableParagraphs;
        internal List<WordParagraph> MutableDefaultHeaderParagraphs => _defaultHeaderParagraphs;
        internal List<WordParagraph> MutableEvenPageHeaderParagraphs => _evenPageHeaderParagraphs;
        internal List<WordParagraph> MutableFirstPageHeaderParagraphs => _firstPageHeaderParagraphs;
        internal List<WordParagraph> MutableDefaultFooterParagraphs => _defaultFooterParagraphs;
        internal List<WordParagraph> MutableEvenPageFooterParagraphs => _evenPageFooterParagraphs;
        internal List<WordParagraph> MutableFirstPageFooterParagraphs => _firstPageFooterParagraphs;
        internal List<WordParagraph> MutableFootnoteParagraphs => _footnoteParagraphs;
        internal List<WordParagraph> MutableEndnoteParagraphs => _endnoteParagraphs;

        /// <summary>
        /// Searches the provided paragraphs for matches of the supplied regular expression and
        /// stores the paragraphs containing matches in the specified list.
        /// </summary>
        /// <param name="paragraphs">Paragraphs to search.</param>
        /// <param name="regex">Regular expression pattern.</param>
        /// <param name="destination">Collection to store paragraphs with matches.</param>
        internal void FindRegex(IEnumerable<WordParagraph> paragraphs, Regex regex, List<WordParagraph> destination) {
            if (paragraphs == null || regex == null) {
                return;
            }

            foreach (var group in paragraphs.GroupBy(paragraph => paragraph._paragraph, OfficeReferenceEqualityComparer<Paragraph>.Instance)) {
                var matches = regex.Matches(string.Concat(group.Select(paragraph => paragraph.Text)));
                if (matches.Count == 0) continue;

                MatchCount += matches.Count;
                int offset = matches[0].Index;
                WordParagraph matchStart = group.First();
                foreach (WordParagraph paragraph in group) {
                    matchStart = paragraph;
                    int length = paragraph.Text.Length;
                    if (offset < length) break;
                    offset -= length;
                }
                destination.Add(matchStart);
            }
        }
    }
}



/// <summary>
/// Provides helper string extensions used by the library.
/// </summary>
public static class StringExtensions {
    /// <summary>
    /// Replaces all occurrences of <paramref name="oldValue"/> with <paramref name="newValue"/> using the specified comparison type.
    /// </summary>
    /// <param name="str">Source string.</param>
    /// <param name="oldValue">Value to search for.</param>
    /// <param name="newValue">Value to replace with.</param>
    /// <param name="comparisonType">String comparison type.</param>
    /// <param name="count">Outputs the number of replacements performed.</param>
    /// <returns>The resulting string after replacements.</returns>
    public static string FindAndReplace(this string str, string oldValue, string newValue, StringComparison comparisonType, ref int count) {
        // Check inputs.
        if (str == null) {
            // Same as original .NET C# string.Replace behavior.
            throw new ArgumentNullException(nameof(str));
        }
        if (str.Length == 0) {
            // Same as original .NET C# string.Replace behavior.
            return str;
        }
        if (oldValue == null) {
            // Same as original .NET C# string.Replace behavior.
            throw new ArgumentNullException(nameof(oldValue));
        }
        if (oldValue.Length == 0) {
            // Same as original .NET C# string.Replace behavior.
            throw new ArgumentException("String cannot be of zero length.");
        }
        // Prepare string builder for storing the processed string.
        // Note: StringBuilder has a better performance than String by 30-40%.
        StringBuilder resultStringBuilder = new StringBuilder(str.Length);

        // Analyze the replacement: replace or remove.
        bool isReplacementNullOrEmpty = string.IsNullOrEmpty(newValue);

        // Replace all values.
        const int valueNotFound = -1;
        int foundAt;
        int startSearchFromIndex = 0;
        while ((foundAt = str.IndexOf(oldValue, startSearchFromIndex, comparisonType)) != valueNotFound) {

            // Append all characters until the found replacement.
            int charsUntilReplacement = foundAt - startSearchFromIndex;
            bool isNothingToAppend = charsUntilReplacement == 0;
            if (!isNothingToAppend) {
                resultStringBuilder.Append(str, startSearchFromIndex, charsUntilReplacement);
            }

            // Process the replacement.
            if (!isReplacementNullOrEmpty) {
                resultStringBuilder.Append(newValue);
            }

            // lets see how many times this was hit
            count++;

            // Prepare start index for the next search.
            // This needed to prevent infinite loop, otherwise method always start search 
            // from the start of the string. For example: if an oldValue == "EXAMPLE", newValue == "example"
            // and comparisonType == "any ignore case" will conquer to replacing:
            // "EXAMPLE" to "example" to "example" to "example" … infinite loop.
            startSearchFromIndex = foundAt + oldValue.Length;
            if (startSearchFromIndex == str.Length) {
                // It is end of the input string: no more space for the next search.
                // The input string ends with a value that has already been replaced. 
                // Therefore, the string builder with the result is complete and no further action is required.
                return resultStringBuilder.ToString();
            }
        }

        // Append the last part to the result.
        int charsUntilStringEnd = str.Length - startSearchFromIndex;
        resultStringBuilder.Append(str, startSearchFromIndex, charsUntilStringEnd);
        return resultStringBuilder.ToString();
    }
}
