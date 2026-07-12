using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core;
using OfficeIMO.Shared;
using OfficeIMO.Word.Fluent;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Collection of all comments inserted into the document.
        /// </summary>
        /// <returns>List of <see cref="WordComment"/> objects for each comment.</returns>
        public List<WordComment> Comments {
            get { return WordComment.GetAllComments(this); }
        }

        /// <summary>
        /// Removes comment with the specified id.
        /// </summary>
        /// <param name="commentId">Id of the comment to remove.</param>
        public void RemoveComment(string commentId) {
            var comment = this.Comments.FirstOrDefault(c => c.Id == commentId);
            comment?.Delete();
        }

        /// <summary>
        /// Removes the specified comment from the document.
        /// </summary>
        /// <param name="comment">Comment instance to remove.</param>
        public void RemoveComment(WordComment comment) {
            comment?.Delete();
        }

        /// <summary>
        /// Removes all comments from the document.
        /// </summary>
        public void RemoveAllComments() {
            foreach (var comment in this.Comments.ToList()) {
                comment.Delete();
            }
        }

        /// <summary>
        /// Gets the value of a document variable or <c>null</c> if the variable does not exist.
        /// </summary>
        /// <param name="name">Variable name.</param>
        public string? GetDocumentVariable(string name) {
            return DocumentVariables.TryGetValue(name, out var value) ? value : null;
        }

        /// <summary>
        /// Sets the value of a document variable. Creates it if it does not exist.
        /// </summary>
        /// <param name="name">Variable name.</param>
        /// <param name="value">Variable value.</param>
        public void SetDocumentVariable(string name, string value) {
            DocumentVariables[name] = value;
        }

        /// <summary>
        /// Removes the document variable with the specified name if present.
        /// </summary>
        /// <param name="name">Variable name.</param>
        public void RemoveDocumentVariable(string name) {
            DocumentVariables.Remove(name);
        }

        /// <summary>
        /// Removes the document variable at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of the variable to remove.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when index is out of range.</exception>
        public void RemoveDocumentVariableAt(int index) {
            if (index < 0 || index >= DocumentVariables.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            string key = DocumentVariables.Keys.ElementAt(index);
            DocumentVariables.Remove(key);
        }

        /// <summary>
        /// Determines whether the document contains any document variables.
        /// </summary>
        public bool HasDocumentVariables => DocumentVariables.Count > 0;

        /// <summary>
        /// Returns a read-only view of all document variables.
        /// </summary>
        public IReadOnlyDictionary<string, string> GetDocumentVariables() {
            return new Dictionary<string, string>(DocumentVariables);
        }

        /// <summary>
        /// Enable or disable tracking of comment changes.
        /// </summary>
        public bool TrackComments {
            get => this.Settings.TrackComments;
            set => this.Settings.TrackComments = value;
        }

        /// <summary>
        /// Enable or disable tracking of all revisions, moves and formatting changes.
        /// </summary>
        public bool TrackChanges {
            get => this.Settings.TrackRevisions;
            set {
                this.Settings.TrackRevisions = value;
                this.Settings.TrackFormatting = value;
                this.Settings.TrackMoves = value;
            }
        }

        /// <summary>
        /// Collection of sections contained in the document.
        /// </summary>
        public List<WordSection> Sections = new List<WordSection>();

        /// <summary>
        /// Path to the file backing this document.
        /// </summary>
        public string FilePath { get; set; } = null!;

        /// <summary>
        /// Original stream where this document was created / loaded from.
        /// </summary>
        internal Stream OriginalStream { get; set; } = null!;

        /// <summary>
        /// Provides access to document settings.
        /// </summary>
        public WordSettings Settings = null!;

        /// <summary>
        /// Manages application related properties.
        /// </summary>
        public ApplicationProperties ApplicationProperties = null!;

        /// <summary>
        /// Provides access to built-in document properties.
        /// </summary>
        public BuiltinDocumentProperties BuiltinDocumentProperties = null!;

        private WordCoverPageProperties? _coverPageProperties;

        /// <summary>
        /// Provides access to the cover page properties custom XML part used by built-in templates.
        /// </summary>
        public WordCoverPageProperties CoverPageProperties => _coverPageProperties ??= new WordCoverPageProperties(this);

        /// <summary>
        /// Collection of custom document properties.
        /// </summary>
        public readonly Dictionary<string, WordCustomProperty> CustomDocumentProperties = new Dictionary<string, WordCustomProperty>();
        /// <summary>
        /// Collection of document variables accessible via <see cref="WordFieldType.DocVariable"/> fields.
        /// </summary>
        public Dictionary<string, string> DocumentVariables { get; } = new Dictionary<string, string>();

        /// <summary>
        /// Collection of bibliographic sources used in the document.
        /// </summary>
        public Dictionary<string, WordBibliographySource> BibliographySources { get; } = new Dictionary<string, WordBibliographySource>();

        /// <summary>
        /// Provides basic statistics for the document.
        /// </summary>
        public WordDocumentStatistics Statistics { get; internal set; } = null!;

        /// <summary>
        /// Gets the persistence policy selected when the document was created or loaded.
        /// </summary>
        public DocumentPersistenceMode PersistenceMode => _persistenceMode;

        /// <summary>
        /// When <c>true</c> the table of contents is flagged to update before saving.
        /// </summary>
        public bool AutoUpdateToc { get; set; }


        // we expose them to help with integration
        /// <summary>
        /// Underlying Open XML word processing document.
        /// </summary>
        public WordprocessingDocument _wordprocessingDocument = null!;

        /// <summary>
        /// Root document element.
        /// </summary>
        public Document _document = null!;
        //public WordCustomProperties _customDocumentProperties;


        /// <summary>
        /// FileOpenAccess of the document
        /// </summary>
        public FileAccess FileOpenAccess => _wordprocessingDocument.FileOpenAccess;
    }
}
