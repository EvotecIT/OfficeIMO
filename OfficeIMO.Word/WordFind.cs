using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Word {
    public class WordFind {
        public int Found = 0;
        public int Replacements = 0;

        public List<WordParagraph> Paragraphs = new List<WordParagraph>();
        public List<WordParagraph> Tables = new List<WordParagraph>();
        public List<WordParagraph> HeaderDefault = new List<WordParagraph>();
        public List<WordParagraph> HeaderEven = new List<WordParagraph>();
        public List<WordParagraph> HeaderFirst = new List<WordParagraph>();
        public List<WordParagraph> FooterDefault = new List<WordParagraph>();
        public List<WordParagraph> FooterEven = new List<WordParagraph>();
        public List<WordParagraph> FooterFirst = new List<WordParagraph>();


        //public List<WordParagraph> Paragraphs {
        //    get;
        //    set;
        //}

        //public List<WordParagraph> Tables {
        //    get;
        //    set;
        //}

        //public List<WordParagraph> HeaderDefault {
        //    get;
        //    set;
        //}

        //public List<WordParagraph> HeaderEven {
        //    get;
        //    set;
        //}

        //public List<WordParagraph> HeaderFirst {
        //    get;
        //    set;
        //}

        //public List<WordParagraph> FooterDefault {
        //    get;
        //    set;
        //}

        //public List<WordParagraph> FooterEven {
        //    get;
        //    set;
        //}

        //public List<WordParagraph> FooterFirst {
        //    get;
        //    set;
        //}
    }
}



public static class StringExtensions {
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
