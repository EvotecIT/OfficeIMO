using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    // Provides functionality for creating, loading and manipulating Word documents.
    public partial class WordDocument : IDisposable {

        /// <summary>Whether this document has released its package and owned resources.</summary>
        public bool IsDisposed => _disposed;

        /// <summary>
        /// This moves section within body from top to bottom to allow footers/headers to move
        /// Needs more work, but this is what Word does all the time
        /// </summary>
        private void MoveSectionProperties() {
            var body = BodyRoot;
            if (body.LastChild is SectionProperties) {
                return;
            }

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
                if (_persistenceMode == DocumentPersistenceMode.SaveOnDispose && wordProcessingDocument.FileOpenAccess != FileAccess.Read) {
                    try {
                        Save();
                    } catch (Exception ex) {
                        persistenceFailure = ex;
                    }
                }

                try {
                    RetainColorContextForInspection();
                } catch (Exception ex) {
                    persistenceFailure ??= ex;
                }

                try {
                    wordProcessingDocument.Dispose();
                } catch (Exception ex) {
                    persistenceFailure ??= ex;
                }

                this._wordprocessingDocument = null!;
            }

            var ownedPackageStream = _ownedPackageStream;
            var legacyValidationLivePackageStream = _legacyValidationLivePackageStream;
            if (legacyValidationLivePackageStream != null &&
                !ReferenceEquals(legacyValidationLivePackageStream, ownedPackageStream)) {
                try {
                    legacyValidationLivePackageStream.Dispose();
                } catch (ObjectDisposedException) {
                    // Disposing an already disposed detached live package stream is harmless.
                } catch (Exception ex) {
                    persistenceFailure ??= ex;
                }
                _legacyValidationLivePackageStream = null;
            }
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
            _legacyValidationEncodedPackageBytes = null;

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
        public async ValueTask DisposeAsync() {
            if (this._disposed) {
                return;
            }

            Exception? persistenceFailure = null;
            var wordProcessingDocument = this._wordprocessingDocument;
            if (wordProcessingDocument != null) {
                if (_persistenceMode == DocumentPersistenceMode.SaveOnDispose && wordProcessingDocument.FileOpenAccess != FileAccess.Read) {
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
                    RetainColorContextForInspection();
                } catch (Exception ex) {
                    persistenceFailure ??= ex;
                }

                try {
                    wordProcessingDocument.Dispose();
                } catch (Exception ex) {
                    persistenceFailure ??= ex;
                }

                this._wordprocessingDocument = null!;
            }

            var ownedPackageStream = _ownedPackageStream;
            var legacyValidationLivePackageStream = _legacyValidationLivePackageStream;
            if (legacyValidationLivePackageStream != null &&
                !ReferenceEquals(legacyValidationLivePackageStream, ownedPackageStream)) {
                try {
                    legacyValidationLivePackageStream.Dispose();
                } catch (ObjectDisposedException) {
                    // Disposing an already disposed detached live package stream is harmless.
                } catch (Exception ex) {
                    persistenceFailure ??= ex;
                }
                _legacyValidationLivePackageStream = null;
            }
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
            _legacyValidationEncodedPackageBytes = null;

            if (this.OriginalStream != null) {
                // Original stream is owned by the caller and should remain open.
            }

            this._disposed = true;
            GC.SuppressFinalize(this);
            if (persistenceFailure != null) {
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(persistenceFailure).Throw();
            }
        }

        private static void InitialiseStyleDefinitions(
            WordprocessingDocument wordDocument,
            bool readOnly,
            bool overrideStyles) {
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
                var styleDefinitionsPart1 = mainPart.AddNewPart<StyleDefinitionsPart>();
                GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);
            }
        }

        internal void EnsureStyleDefinitionsInitialized() {
            if (FileOpenAccess == FileAccess.Read) return;
            var mainPart = _wordprocessingDocument.MainDocumentPart;
            if (mainPart == null) return;
            var part = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
            part.Styles ??= new Styles();
        }

        internal WordSection _currentSection => this.Sections.Last();



    }
}
