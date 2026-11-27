using OfficeIMO.Word.LegacyDoc;
using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OfficeIMO.Word.LegacyDoc.Docx;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Threading;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Loads a legacy binary `.doc` document and projects supported content into a normal OfficeIMO Word document.
        /// The resulting document saves through the normal Open XML path.
        /// </summary>
        internal static WordDocument LoadLegacyDoc(string path, LegacyDocImportOptions? options = null) {
            LegacyDocDocument document = LegacyDocDocument.Load(path, options);
            return ProjectLoadedLegacyDocDocument(document, path);
        }

        /// <summary>
        /// Loads a legacy binary `.doc` document and returns both the projected OfficeIMO document and import report.
        /// </summary>
        internal static LegacyDocLoadResult LoadLegacyDocWithReport(string path, LegacyDocImportOptions? options = null) {
            LegacyDocDocument document = LegacyDocDocument.Load(path, options);
            return CreateLegacyDocLoadResult(document, path);
        }

        /// <summary>
        /// Loads a legacy binary `.doc` stream and projects supported content into a normal OfficeIMO Word document.
        /// The resulting document saves through the normal Open XML path.
        /// </summary>
        internal static WordDocument LoadLegacyDoc(Stream stream, LegacyDocImportOptions? options = null) {
            LegacyDocDocument document = LegacyDocDocument.Load(stream, options);
            return ProjectLoadedLegacyDocDocument(document, sourcePath: null);
        }

        /// <summary>
        /// Loads a legacy binary `.doc` stream and returns both the projected OfficeIMO document and import report.
        /// </summary>
        internal static LegacyDocLoadResult LoadLegacyDocWithReport(Stream stream, LegacyDocImportOptions? options = null) {
            LegacyDocDocument document = LegacyDocDocument.Load(stream, options);
            return CreateLegacyDocLoadResult(document, sourcePath: null);
        }

        private static WordDocument LoadLegacyDocFromNormalFlow(
            byte[] bytes,
            string? sourcePath,
            bool saveOnDispose,
            bool readOnly,
            LegacyDocImportOptions? importOptions = null,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            if (saveOnDispose) {
                throw new NotSupportedException("SaveOnDispose is not supported when loading legacy binary .doc files. Load the document, then save explicitly to a .docx path.");
            }

            LegacyDocDocument document = LegacyDocDocument.Load(bytes, importOptions ?? new LegacyDocImportOptions(), cancellationToken);
            LegacyDocImportDiagnostic[] errors = document.Diagnostics
                .Where(diagnostic => diagnostic.Severity == LegacyDocDiagnosticSeverity.Error)
                .ToArray();
            if (errors.Length > 0) {
                throw new InvalidDataException("Legacy DOC import failed: " + FormatLegacyDocDiagnostics(errors));
            }

            return ProjectLoadedLegacyDocDocument(document, sourcePath, attachSourcePathForSave: sourcePath != null, readOnly, cancellationToken);
        }

        private static LegacyDocLoadResult CreateLegacyDocLoadResult(LegacyDocDocument legacyDocument, string? sourcePath) {
            try {
                return new LegacyDocLoadResult(ProjectLoadedLegacyDocDocument(legacyDocument, sourcePath, attachSourcePathForSave: false), legacyDocument);
            } catch (InvalidDataException exception) {
                return new LegacyDocLoadResult(document: null, legacyDocument, exception);
            }
        }

        private static WordDocument ProjectLoadedLegacyDocDocument(LegacyDocDocument legacyDocument, string? sourcePath) =>
            ProjectLoadedLegacyDocDocument(legacyDocument, sourcePath, attachSourcePathForSave: sourcePath != null);

        /// <summary>
        /// Writes the DOCX package for the decoded model through the owned Open XML layer and loads it as a normal
        /// document. Binary DOC projects its own formatting onto the engine's default package, without authoring defaults.
        /// </summary>
        private static WordDocument ProjectLoadedLegacyDocDocument(LegacyDocDocument legacyDocument, string? sourcePath, bool attachSourcePathForSave,
            bool readOnly = false, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            LegacyDocImportDiagnostic[] errors = legacyDocument.Diagnostics
                .Where(diagnostic => diagnostic.Severity == LegacyDocDiagnosticSeverity.Error)
                .ToArray();
            if (errors.Length > 0) {
                throw new InvalidDataException("Legacy DOC import failed: " + FormatLegacyDocDiagnostics(errors));
            }

            byte[] package = LegacyDocDocxWriter.Write(legacyDocument, DefaultDocumentPackage.Value, cancellationToken);
            // Custom properties are written through the property API, which needs an editable document.
            bool writeCustomProperties = legacyDocument.DocumentProperties.CustomProperties.Count > 0;
            WordDocument document = LoadProjectedLegacyDocPackage(package, readOnly && !writeCustomProperties);
            try {
                if (writeCustomProperties) {
                    ApplyLegacyDocCustomProperties(document, legacyDocument.DocumentProperties);
                    if (readOnly) {
                        using var packageStream = new MemoryStream();
                        // An internal package transition that only enforces read-only access; the caller has not asked
                        // for a lossy export, so the public loss gate must not block it.
                        document.Save(packageStream, WordFileFormat.Docx, new WordSaveOptions { AllowLoss = true });
                        document.Dispose();
                        document = LoadProjectedLegacyDocPackage(packageStream.ToArray(), readOnly: true);
                    }
                }

                cancellationToken.ThrowIfCancellationRequested();
                // The document was loaded from the projected package, so its bookmark, comment and drawing identifiers
                // are read from that package when first needed.
                document.MarkLoadedFromLegacyDoc(sourcePath, legacyDocument, attachSourcePathForSave);
                return document;
            } catch {
                document.Dispose();
                throw;
            }
        }

        private static WordDocument LoadProjectedLegacyDocPackage(byte[] package, bool readOnly) {
            var options = new WordLoadOptions {
                AccessMode = readOnly ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite,
                PersistenceMode = DocumentPersistenceMode.Explicit
            };
            WordDocument document = TryLoadWithoutSdk(package, options, filePath: null, originalStream: null)
                ?? Load(new MemoryStream(package, writable: false), options);
            // Like a new document made from the default package, the projection carries no package signatures and
            // no relationships from other producers.
            document._packageContentTypesAdvertiseXmlSignatures = false;
            document._normalizeOpenOfficeRelationshipsOnSave = false;
            return document;
        }

        private static void ApplyLegacyDocCustomProperties(WordDocument document, LegacyDocDocumentProperties properties) {
            foreach (KeyValuePair<string, LegacyDocDocumentPropertyValue> property in properties.CustomProperties) {
                if (TryCreateWordCustomProperty(property.Value, out WordCustomProperty? wordProperty)) {
                    document.CustomProperties[property.Key] = wordProperty!;
                }
            }
        }

        private static bool TryCreateWordCustomProperty(LegacyDocDocumentPropertyValue property, out WordCustomProperty? wordProperty) {
            switch (property.Kind) {
                case LegacyDocDocumentPropertyValueKind.Text:
                    wordProperty = new WordCustomProperty(System.Convert.ToString(property.Value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty);
                    return true;
                case LegacyDocDocumentPropertyValueKind.Boolean:
                    wordProperty = new WordCustomProperty(System.Convert.ToBoolean(property.Value, System.Globalization.CultureInfo.InvariantCulture));
                    return true;
                case LegacyDocDocumentPropertyValueKind.DateTime:
                    wordProperty = new WordCustomProperty(System.Convert.ToDateTime(property.Value, System.Globalization.CultureInfo.InvariantCulture));
                    return true;
                case LegacyDocDocumentPropertyValueKind.Integer:
                    if (property.Value is long integer64) {
                        wordProperty = new WordCustomProperty(integer64);
                    } else {
                        wordProperty = new WordCustomProperty(System.Convert.ToInt32(property.Value, System.Globalization.CultureInfo.InvariantCulture));
                    }

                    return true;
                case LegacyDocDocumentPropertyValueKind.Number:
                    wordProperty = new WordCustomProperty(OfficeIMO.Core.Internal.OfficeRoundTripText.ToDouble(property.Value, System.Globalization.CultureInfo.InvariantCulture));
                    return true;
                case LegacyDocDocumentPropertyValueKind.Binary:
                    if (property.Value is byte[] bytes) {
                        wordProperty = new WordCustomProperty(bytes);
                        return true;
                    }

                    wordProperty = null;
                    return false;
                default:
                    wordProperty = null;
                    return false;
            }
        }

        private static string FormatLegacyDocDiagnostics(IEnumerable<LegacyDocImportDiagnostic> diagnostics) {
            const int maxDiagnostics = 6;
            LegacyDocImportDiagnostic[] selected = diagnostics.Take(maxDiagnostics + 1).ToArray();
            string message = string.Join("; ", selected.Take(maxDiagnostics).Select(diagnostic => diagnostic.ToString()));
            if (selected.Length > maxDiagnostics) {
                message += $"; and {selected.Length - maxDiagnostics} more";
            }

            return message;
        }
    }
}
