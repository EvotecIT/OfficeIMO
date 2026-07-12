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

        /// <summary>
        /// This moves section within body from top to bottom to allow footers/headers to move
        /// Needs more work, but this is what Word does all the time
        /// </summary>
        private void MoveSectionProperties() {
            var body = BodyRoot;
            var sectionProperties = body.Elements<SectionProperties>().LastOrDefault();
            if (sectionProperties != null) {
                body.RemoveChild(sectionProperties);
                body.Append(sectionProperties);
            }
        }

        /// <summary>
        /// Releases resources associated with this <see cref="WordDocument"/> instance.
        /// </summary>
        public void Dispose() {
            if (this._disposed) {
                return;
            }

            Exception? persistenceFailure = null;
            var wordProcessingDocument = this._wordprocessingDocument;
            if (wordProcessingDocument != null) {
                if (wordProcessingDocument.AutoSave && wordProcessingDocument.FileOpenAccess != FileAccess.Read) {
                    try {
                        Save();
                    } catch (Exception ex) {
                        persistenceFailure = ex;
                    }
                }

                try {
                    wordProcessingDocument.Dispose();
                } catch (Exception ex) {
                    persistenceFailure ??= ex;
                }

                this._wordprocessingDocument = null!;
            }

            var ownedPackageStream = _ownedPackageStream;
            if (ownedPackageStream != null) {
                try {
                    ownedPackageStream.Dispose();
                } catch (ObjectDisposedException) {
                    // Disposing an already disposed owned stream is harmless.
                } catch (Exception ex) {
                    persistenceFailure ??= ex;
                }

                _ownedPackageStream = null;
            }

            if (this.OriginalStream != null) {
                // Original stream is owned by the caller and should remain open.
            }

            this._disposed = true;
            GC.SuppressFinalize(this);
            if (persistenceFailure != null) {
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(persistenceFailure).Throw();
            }
        }

        /// <summary>
        /// Releases resources associated with the document asynchronously.
        /// </summary>
        public async Task DisposeAsync() {
            if (this._disposed) {
                return;
            }

            Exception? persistenceFailure = null;
            var wordProcessingDocument = this._wordprocessingDocument;
            if (wordProcessingDocument != null) {
                if (wordProcessingDocument.AutoSave && wordProcessingDocument.FileOpenAccess != FileAccess.Read) {
                    try {
                        if (string.IsNullOrEmpty(FilePath) && OriginalStream != null) {
                            await SaveAsync(OriginalStream).ConfigureAwait(false);
                        } else {
                            await SaveAsync().ConfigureAwait(false);
                        }
                    } catch (Exception ex) {
                        persistenceFailure = ex;
                    }
                }

                try {
                    wordProcessingDocument.Dispose();
                } catch (Exception ex) {
                    persistenceFailure ??= ex;
                }

                this._wordprocessingDocument = null!;
            }

            var ownedPackageStream = _ownedPackageStream;
            if (ownedPackageStream != null) {
                try {
                    ownedPackageStream.Dispose();
                } catch (ObjectDisposedException) {
                    // Disposing an already disposed owned stream is harmless.
                } catch (Exception ex) {
                    persistenceFailure ??= ex;
                }

                _ownedPackageStream = null;
            }

            if (this.OriginalStream != null) {
                // Original stream is owned by the caller and should remain open.
            }

            this._disposed = true;
            GC.SuppressFinalize(this);
            if (persistenceFailure != null) {
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(persistenceFailure).Throw();
            }
        }

        private static void InitialiseStyleDefinitions(WordprocessingDocument wordDocument, bool readOnly, bool overrideStyles) {
            // In read-only mode we don't touch styles.
            if (readOnly) return;

            // Guard against malformed packages missing a main document part.
            var mainPart = wordDocument.MainDocumentPart;
            if (mainPart == null) {
                // Nothing we can do; leave silently to avoid NREs when loading odd files.
                return;
            }

            var styleDefinitionsPart = mainPart
                .GetPartsOfType<StyleDefinitionsPart>()
                .FirstOrDefault();
            if (styleDefinitionsPart != null) {
                // Safe-guard missing Styles root element.
                styleDefinitionsPart.Styles ??= new Styles();
                AddStyleDefinitions(styleDefinitionsPart, overrideStyles);
            } else {
                // Create Styles part if it doesn't exist yet.
                var styleDefinitionsPart1 = mainPart.AddNewPart<StyleDefinitionsPart>("rId1");
                GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);
            }
        }

        internal WordSection _currentSection => this.Sections.Last();


        private static void EnsureCustomStyleNames(WordprocessingDocument wordDocument) {
            var stylePart = wordDocument.MainDocumentPart?.StyleDefinitionsPart;
            var styles = stylePart?.Styles;
            if (styles == null) return;
            var map = WordParagraphStyle.CustomStyles
                .Where(s => !string.IsNullOrEmpty(s.StyleId?.Value))
                .GroupBy(s => s.StyleId!.Value!, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.Last(), StringComparer.OrdinalIgnoreCase);
            bool changed = false;
            foreach (var s in styles.OfType<DocumentFormat.OpenXml.Wordprocessing.Style>()) {
                var id = s.StyleId?.Value;
                if (id != null && map.TryGetValue(id, out var def)) {
                    var newName = def.StyleName?.Val ?? def.StyleId?.Value;
                    if (!string.IsNullOrEmpty(newName)) {
                        if (s.StyleName == null || !string.Equals(s.StyleName.Val, newName, StringComparison.Ordinal)) {
                            s.StyleName = new DocumentFormat.OpenXml.Wordprocessing.StyleName { Val = newName };
                            changed = true;
                        }
                    }
                }
            }
            if (changed) styles.Save();
        }

    }
}
