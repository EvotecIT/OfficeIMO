using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing.Internal;
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

        /// <summary>
        /// Provides access to the document background settings.
        /// </summary>
        public WordBackground Background { get; set; } = null!;

        /// <summary>
        /// Indicates whether the document passes Open XML validation.
        /// </summary>
        public bool DocumentIsValid {
            get {
                if (DocumentValidationErrors.Count > 0) {
                    return false;
                }

                return true;
            }
        }

        /// <summary>
        /// Gets the list of validation errors for the document.
        /// </summary>
        public List<ValidationErrorInfo> DocumentValidationErrors {
            get {
                return ValidateDocument();
            }
        }

        /// <summary>
        /// Validates the document using the specified file format version.
        /// </summary>
        /// <param name="fileFormatVersions">File format version to validate against.</param>
        /// <returns>List of validation errors.</returns>
        public List<ValidationErrorInfo> ValidateDocument(FileFormatVersions fileFormatVersions = FileFormatVersions.Microsoft365) {
            List<ValidationErrorInfo> listErrors = new List<ValidationErrorInfo>();
            OpenXmlValidator validator = new OpenXmlValidator(fileFormatVersions);
            foreach (ValidationErrorInfo error in validator.Validate(this._wordprocessingDocument)) {
                listErrors.Add(error);
            }
            return listErrors;
        }

        /// <summary>
        /// Creates a fluent wrapper for this document.
        /// </summary>
        /// <returns>A new <see cref="WordFluentDocument"/> instance.</returns>
        public WordFluentDocument AsFluent() {
            return new WordFluentDocument(this);
        }

        /// <summary>
        /// Gets or sets compatibility settings for the document.
        /// </summary>
        public WordCompatibilitySettings CompatibilitySettings { get; set; } = null!;

        internal void NotifyTableOfContentUpdateQueued() {
            _tableOfContentUpdateQueued = true;
        }

        internal void ResetTableOfContentUpdateQueue() {
            _tableOfContentUpdateQueued = false;
        }

        /// <summary>
        /// Ensures heading edits keep the table-of-contents refresh state aligned with the document settings.
        /// </summary>
        internal void HeadingModified() {
            var updateOnOpen = Settings.UpdateFieldsOnOpen;

            if (_tableOfContentUpdateQueued) {
                if (updateOnOpen) {
                    // Updates are already queued and Word will refresh them on open.
                    return;
                }

                // Word will not refresh fields anymore, so drop the stale queued state before requeueing.
                ResetTableOfContentUpdateQueue();
            }

            if (updateOnOpen) {
                // Keep the queue flag in sync with Word's behaviour when UpdateFieldsOnOpen is set by the user.
                _tableOfContentUpdateQueued = true;
                return;
            }

            var tableOfContent = TableOfContent;
            if (tableOfContent == null) {
                return;
            }

            // Re-enable the automatic refresh by marking the table-of-contents fields dirty again.
            tableOfContent.QueueUpdateOnOpen(force: true);
        }

        private void PreSaving() {
            MoveSectionProperties();
            // Keep tblGrid consistent for online viewers without changing authoring semantics.
            NormalizeTablesForOnline();
            SaveNumbering();
            if (AutoUpdateToc && TableOfContent != null) {
                TableOfContent.Update();
            }
            new WordCustomProperties(this, true);
            var settingsPart = _wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart;
            bool hasVariables = settingsPart?.Settings?.GetFirstChild<DocumentVariables>() != null;
            if (hasVariables || DocumentVariables.Count > 0) {
                new WordDocumentVariables(this, true);
            }
            new WordBibliography(this, true);
        }
    }
}
