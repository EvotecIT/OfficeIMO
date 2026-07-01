using System.Collections.Generic;

namespace OfficeIMO.Word {
    /// <summary>
    /// Controls optional normalization and feature families used by <see cref="WordDocumentComparer.CompareStructure(string, string, WordComparisonOptions?)"/>.
    /// </summary>
    public sealed class WordComparisonOptions {
        /// <summary>
        /// Gets a new options instance with the default comparison behavior.
        /// </summary>
        public static WordComparisonOptions Default => new();

        /// <summary>
        /// Ignores differences caused only by whitespace runs in text-bearing comparison scopes.
        /// </summary>
        public bool IgnoreWhitespace { get; set; }

        /// <summary>
        /// Ignores character casing in text-bearing comparison scopes.
        /// </summary>
        public bool IgnoreCase { get; set; }

        /// <summary>
        /// Compares run formatting changes such as bold, italic, fonts, and other run properties.
        /// </summary>
        public bool CompareRunFormatting { get; set; } = true;

        /// <summary>
        /// Resolves document defaults, based-on style chains, paragraph styles, character styles, and direct properties before comparing formatting.
        /// </summary>
        public bool CompareEffectiveFormatting { get; set; } = true;

        /// <summary>
        /// Compares paragraph style identifiers as paragraph findings.
        /// </summary>
        public bool CompareParagraphStyleIds { get; set; } = true;

        /// <summary>
        /// Includes run style identifiers when comparing run formatting.
        /// </summary>
        public bool CompareRunStyleIds { get; set; } = true;

        /// <summary>
        /// Limits output to the specified finding scopes. When null or empty, all enabled scopes are included.
        /// </summary>
        public ISet<WordComparisonScope>? IncludedScopes { get; set; }

        /// <summary>
        /// Removes the specified finding scopes from the output after all enabled analyzers run.
        /// </summary>
        public ISet<WordComparisonScope>? ExcludedScopes { get; set; }

        /// <summary>
        /// Compares Word fields as feature-specific findings.
        /// </summary>
        public bool CompareFields { get; set; } = true;

        /// <summary>
        /// Compares structured document tags/content controls as feature-specific findings.
        /// </summary>
        public bool CompareContentControls { get; set; } = true;

        /// <summary>
        /// Compares bookmark names, range text, and locations as feature-specific findings.
        /// </summary>
        public bool CompareBookmarks { get; set; } = true;

        /// <summary>
        /// Compares internal and external hyperlinks as feature-specific findings.
        /// </summary>
        public bool CompareHyperlinks { get; set; } = true;

        /// <summary>
        /// Compares numbered and bulleted list items as feature-specific findings.
        /// </summary>
        public bool CompareLists { get; set; } = true;

        /// <summary>
        /// Compares comments, replies, target text, and resolved metadata as feature-specific findings.
        /// </summary>
        public bool CompareComments { get; set; } = true;

        /// <summary>
        /// Compares comment author and initials metadata.
        /// </summary>
        public bool CompareCommentAuthors { get; set; } = true;

        /// <summary>
        /// Compares comment body text.
        /// </summary>
        public bool CompareCommentText { get; set; } = true;

        /// <summary>
        /// Compares comment resolved/unresolved metadata.
        /// </summary>
        public bool CompareCommentResolvedState { get; set; } = true;

        /// <summary>
        /// Compares comment target text and target container flags.
        /// </summary>
        public bool CompareCommentTargets { get; set; } = true;

        /// <summary>
        /// Compares comment replies as separate review findings.
        /// </summary>
        public bool CompareCommentReplies { get; set; } = true;

        /// <summary>
        /// Compares tracked revisions and tracked formatting metadata as feature-specific findings.
        /// </summary>
        public bool CompareRevisions { get; set; } = true;

        /// <summary>
        /// Compares revision author metadata.
        /// </summary>
        public bool CompareRevisionAuthors { get; set; } = true;

        /// <summary>
        /// Compares affected and nearby revision text.
        /// </summary>
        public bool CompareRevisionText { get; set; } = true;

        /// <summary>
        /// Compares revision part and container-location metadata.
        /// </summary>
        public bool CompareRevisionLocations { get; set; } = true;

        /// <summary>
        /// Compares embedded and externally linked images.
        /// </summary>
        public bool CompareImages { get; set; } = true;

        /// <summary>
        /// Compares document block order when the same comparable blocks appear in a different order.
        /// </summary>
        public bool CompareBlockOrder { get; set; } = true;

        /// <summary>
        /// Compares generated identifiers such as bookmark ids, relationship ids, comment paragraph ids, and revision ids.
        /// </summary>
        public bool CompareGeneratedIds { get; set; } = true;

        /// <summary>
        /// Compares volatile metadata such as generated comment and revision timestamps.
        /// </summary>
        public bool CompareVolatileMetadata { get; set; } = true;
    }
}
