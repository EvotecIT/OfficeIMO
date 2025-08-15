using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents the results of a find and replace operation within a document.
    /// </summary>
    public class WordFind {
        /// <summary>Gets or sets the number of matches found.</summary>
        public int Found = 0;
        /// <summary>Gets or sets the number of replacements performed.</summary>
        public int Replacements = 0;

        /// <summary>Paragraphs containing matches.</summary>
        public List<WordParagraph> Paragraphs = new List<WordParagraph>();
        /// <summary>Table cells containing matches.</summary>
        public List<WordParagraph> Tables = new List<WordParagraph>();
        /// <summary>Header paragraphs containing matches when using the default header.</summary>
        public List<WordParagraph> HeaderDefault = new List<WordParagraph>();
        /// <summary>Header paragraphs containing matches when using even page headers.</summary>
        public List<WordParagraph> HeaderEven = new List<WordParagraph>();
        /// <summary>Header paragraphs containing matches when using the first page header.</summary>
        public List<WordParagraph> HeaderFirst = new List<WordParagraph>();
        /// <summary>Footer paragraphs containing matches when using the default footer.</summary>
        public List<WordParagraph> FooterDefault = new List<WordParagraph>();
        /// <summary>Footer paragraphs containing matches when using even page footers.</summary>
        public List<WordParagraph> FooterEven = new List<WordParagraph>();
        /// <summary>Footer paragraphs containing matches when using the first page footer.</summary>
        public List<WordParagraph> FooterFirst = new List<WordParagraph>();

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

            foreach (var paragraph in paragraphs) {
                var matches = regex.Matches(paragraph.Text);
                if (matches.Count > 0) {
                    Found += matches.Count;
                    destination.Add(paragraph);
                }
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
        List<string> list = new List<string>();
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
