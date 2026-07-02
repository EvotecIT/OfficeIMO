using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Shared;
using OfficeIMO.Word.Fluent;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides functionality for creating, loading and manipulating Word documents.
    /// </summary>
    public partial class WordDocument : IDisposable {

        internal int BookmarkId {
            get {
                return Bookmarks.Select(bookmark => bookmark.Id).DefaultIfEmpty(0).Max() + 1;
            }
        }

        /// <summary>
        /// Gets the table of contents defined in the document.
        /// </summary>
        public WordTableOfContent? TableOfContent {
            get {
                var sdtBlocks = _document.Body?.ChildElements.OfType<SdtBlock>() ?? Enumerable.Empty<SdtBlock>();
                return WordSection.ConvertStdBlockToTableOfContent(this, sdtBlocks);
            }
        }

        /// <summary>
        /// Gets the cover page if one is defined in the document.
        /// </summary>
        public WordCoverPage? CoverPage {
            get {
                var sdtBlocks = _document.Body?.ChildElements.OfType<SdtBlock>() ?? Enumerable.Empty<SdtBlock>();
                return WordSection.ConvertStdBlockToCoverPage(this, sdtBlocks);
            }
        }
    }
}
