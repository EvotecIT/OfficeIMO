using System;

namespace OfficeIMO.Word {
    public partial class WordHelpers {
        /// <summary>
        /// Provides the next available structured document tag identifier for the specified document.
        /// The allocator guarantees positive, unique identifiers within the document scope.
        /// </summary>
        /// <param name="document">Target document that owns the structured document tags.</param>
        /// <returns>A positive integer identifier that has not been used in the document.</returns>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="document"/> is <c>null</c>.</exception>
        public static int GetNextSdtId(WordDocument document) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            return document.GenerateSdtId();
        }
    }
}
