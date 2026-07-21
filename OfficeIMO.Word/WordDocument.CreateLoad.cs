using OfficeIMO.Drawing.Internal;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
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

        private static string GetUniqueFilePath(string filePath) {
            if (File.Exists(filePath)) {
                string folderPath = Path.GetDirectoryName(filePath)!;
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string fileExtension = Path.GetExtension(filePath);
                int number = 1;

                Match regex = Regex.Match(fileName, @"^(.+) \((\d+)\)$");

                if (regex.Success) {
                    fileName = regex.Groups[1].Value;
                    number = int.Parse(regex.Groups[2].Value);
                }

                do {
                    number++;
                    string newFileName = $"{fileName} ({number}){fileExtension}";
                    filePath = Path.Combine(folderPath, newFileName);
                } while (File.Exists(filePath));
            }

            return filePath;
        }

        private static OpenSettings CreateOpenSettings(OpenSettings? openSettings) {
            if (openSettings is null) {
                return new OpenSettings { AutoSave = false };
            }

            return new OpenSettings {
                AutoSave = false,
                CompatibilityLevel = openSettings.CompatibilityLevel,
                MarkupCompatibilityProcessSettings = openSettings.MarkupCompatibilityProcessSettings,
                MaxCharactersInPart = openSettings.MaxCharactersInPart,
            };
        }

        private static byte[] ReadSourceBytes(Stream stream, OfficePackageSecurityOptions? securityOptions) =>
            securityOptions == null
                ? OfficeStreamReader.ReadAllBytes(stream)
                : OfficePackageSecurityInspector.ReadBounded(stream, securityOptions);

        private static async Task<byte[]> ReadSourceBytesAsync(Stream stream,
            OfficePackageSecurityOptions? securityOptions, CancellationToken cancellationToken) =>
            securityOptions == null
                ? await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken).ConfigureAwait(false)
                : await OfficePackageSecurityInspector.ReadBoundedAsync(stream, securityOptions, cancellationToken)
                    .ConfigureAwait(false);

        private static void ValidateSourcePackage(byte[] sourceBytes,
            OfficePackageSecurityOptions? securityOptions) {
            if (securityOptions != null) {
                OfficePackageSecurityInspector.Validate(sourceBytes, securityOptions);
            }
        }

        /// <summary>
        /// Create a new WordDocument
        /// </summary>
        /// <param name="options">Creation options. SaveOnDispose is invalid without an associated destination.</param>
        public static WordDocument Create(WordCreateOptions? options = null) {
            WordCreateOptions resolved = options ?? new WordCreateOptions();
            if (resolved.PersistenceMode == DocumentPersistenceMode.SaveOnDispose) {
                throw new ArgumentException("SaveOnDispose requires an associated file path or writable stream.", nameof(options));
            }
            return CreateInternal(filePath: null, stream: null, resolved.DocumentType, resolved.PersistenceMode);
        }

        /// <summary>Creates a Word document associated with a path that is written on explicit save.</summary>
        /// <param name="filePath">Destination associated with the document.</param>
        /// <param name="options">Creation and persistence options.</param>
        public static WordDocument Create(string filePath, WordCreateOptions? options = null) {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            WordCreateOptions resolved = options ?? new WordCreateOptions();
            var documentType = GetDocumentType(filePath);
            var word = CreateInternal(filePath, null, documentType, resolved.PersistenceMode);
            return word;
        }

        private static WordprocessingDocumentType GetDocumentType(string? filePath) {
            if (string.IsNullOrEmpty(filePath)) {
                return WordprocessingDocumentType.Document;
            }

            var extension = Path.GetExtension(filePath);
            return extension.ToLowerInvariant() switch {
                ".docm" => WordprocessingDocumentType.MacroEnabledDocument,
                ".dotx" => WordprocessingDocumentType.Template,
                ".dotm" => WordprocessingDocumentType.MacroEnabledTemplate,
                _ => WordprocessingDocumentType.Document
            };
        }

        private static void AlignDocumentTypeWithFilePath(WordprocessingDocument document, string filePath) {
            var documentType = GetDocumentType(filePath);
            if (document.DocumentType != documentType) {
                document.ChangeDocumentType(documentType);
            }

            if (!IsMacroEnabledDocumentType(documentType)) {
                RemoveVbaProjectPart(document);
            }
        }

        private static bool IsMacroEnabledDocumentType(WordprocessingDocumentType documentType) {
            return documentType == WordprocessingDocumentType.MacroEnabledDocument ||
                   documentType == WordprocessingDocumentType.MacroEnabledTemplate;
        }

        private static void RemoveVbaProjectPart(WordprocessingDocument document) {
            var mainPart = document.MainDocumentPart;
            if (mainPart?.VbaProjectPart != null) {
                mainPart.DeletePart(mainPart.VbaProjectPart);
            }
        }

        private static WordDocument CreateInternal(string? filePath, Stream? stream, WordprocessingDocumentType documentType, DocumentPersistenceMode persistenceMode) {
            WordDocument word = new WordDocument();
            if (stream != null) {
                word.OriginalStream = stream;
            }

            var packageStream = new MemoryStream();
            WordprocessingDocument wordDocument = WordprocessingDocument.Create(packageStream, documentType, autoSave: false);

            wordDocument.AddMainDocumentPart();
            var mainPart = wordDocument.MainDocumentPart!;
            mainPart.Document = new Document() {
                MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14" }
            };
            mainPart.Document.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            mainPart.Document.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            mainPart.Document.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            mainPart.Document.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            mainPart.Document.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            mainPart.Document.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            mainPart.Document.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            mainPart.Document.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            mainPart.Document.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            mainPart.Document.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            mainPart.Document.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            mainPart.Document.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            mainPart.Document.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            mainPart.Document.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            mainPart.Document.AddNamespaceDeclaration("oel", "http://schemas.microsoft.com/office/2019/extlst");
            mainPart.Document.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            mainPart.Document.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            mainPart.Document.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            mainPart.Document.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            mainPart.Document.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            mainPart.Document.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            mainPart.Document.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            mainPart.Document.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            mainPart.Document.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            mainPart.Document.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            mainPart.Document.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            mainPart.Document.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            mainPart.Document.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            mainPart.Document.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            mainPart.Document.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            mainPart.Document.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            mainPart.Document.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            mainPart.Document.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            mainPart.Document.Body = new DocumentFormat.OpenXml.Wordprocessing.Body();

            word.FilePath = filePath;
            word._ownedPackageStream = packageStream;
            word._wordprocessingDocument = wordDocument;
            word._document = mainPart.Document;
            word._persistenceMode = persistenceMode;
            word.InitializeSdtIdState();

            StyleDefinitionsPart styleDefinitionsPart1 = mainPart.AddNewPart<StyleDefinitionsPart>("rId1");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            WebSettingsPart webSettingsPart1 = mainPart.AddNewPart<WebSettingsPart>("rId3");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            DocumentSettingsPart documentSettingsPart1 = mainPart.AddNewPart<DocumentSettingsPart>("rId2");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            EndnotesPart endnotesPart1 = mainPart.AddNewPart<EndnotesPart>("rId4");
            GenerateEndNotesPart1Content(endnotesPart1);

            FootnotesPart footnotesPart1 = mainPart.AddNewPart<FootnotesPart>("rId5");
            GenerateFootNotesPart1Content(footnotesPart1);

            FontTablePart fontTablePart1 = mainPart.AddNewPart<FontTablePart>("rId6");
            GenerateFontTablePart1Content(fontTablePart1);

            ThemePart themePart1 = mainPart.AddNewPart<ThemePart>("rId7");
            GenerateThemePart1Content(themePart1);

            new WordSettings(word);
            new WordCompatibilitySettings(word);
            new ApplicationProperties(word);
            new BuiltinDocumentProperties(word);
            new WordSection(word, null!);
            new WordBackground(word);
            new WordDocumentStatistics(word);

            WordListStyles.InitializeAbstractNumberId(word._wordprocessingDocument);

            return word;
        }

        /// <summary>
        /// Create a new <see cref="WordDocument"/> writing directly to the provided stream.
        /// </summary>
        /// <param name="stream">Destination stream.</param>
        /// <param name="options">Creation and persistence options.</param>
        /// <returns>Instance of <see cref="WordDocument"/>.</returns>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="stream"/> is null.</exception>
        public static WordDocument Create(Stream stream, WordCreateOptions? options = null) {
            OfficeDocumentLifecycle.EnsureAssociatedDestination(stream, nameof(stream));
            WordCreateOptions resolved = options ?? new WordCreateOptions();
            var word = CreateInternal(null, stream, resolved.DocumentType, resolved.PersistenceMode);
            return word;
        }

        /// <summary>
        /// PreSaving function to be called before saving the document
        /// </summary>
        private void LoadDocument() {
            Sections.Clear();
            InitializeSdtIdState();
            // add settings if not existing
            new WordSettings(this);
            new ApplicationProperties(this);
            new BuiltinDocumentProperties(this);
            new WordCustomProperties(this);
            new WordDocumentVariables(this);
            new WordBibliography(this);
            new WordBackground(this);
            new WordDocumentStatistics(this);
            new WordCompatibilitySettings(this);
            //CustomDocumentProperties customDocumentProperties = new CustomDocumentProperties(this);
            // add a section that's assigned to top of the document
            var wordSection = new WordSection(this, null!, null!);

            var list = BodyRoot.ChildElements.ToList(); //.OfType<Paragraph>().ToList();
            foreach (var element in list) {
                if (element is Paragraph) {
                    Paragraph paragraph = (Paragraph)element;
                    if (paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.SectionProperties != null) {
                        wordSection = new WordSection(this, paragraph.ParagraphProperties.SectionProperties, paragraph);
                    }
                } else if (element is Table) {
                    // WordTable wordTable = new WordTable(this, wordSection, (Table)element);
                } else if (element is SectionProperties) {
                    // we don't do anything as we already created it above - i think
                } else if (element is SdtBlock) {
                    // we don't do anything as we load stuff with get on demand
                } else if (element is OpenXmlUnknownElement) {
                    // this happens when adding dirty element - mainly during TOC Update() function
                } else if (element is BookmarkEnd) {

                } else {
                    //throw new NotImplementedException("This isn't implemented yet");
                }
            }

            RearrangeSectionsAfterLoad();
        }

        /// <summary>
        /// Rearrange sections after loading the document
        /// </summary>
        private void RearrangeSectionsAfterLoad() {
            if (Sections.Count > 0) {
                //var firstElement = Sections[0];
                var firstElementHeader = Sections[0].Header;
                var firstElementFooter = Sections[0].Footer;
                var firstElementSection = Sections[0]._sectionProperties;

                for (int i = 0; i < Sections.Count; i++) {
                    var element = Sections[i];
                    //var tempFooter = element.Footer;
                    //var tempHeader = element.Header;
                    //var tempSectionProp = element._sectionProperties;

                    if (i + 1 < Sections.Count) {
                        Sections[i].Footer = Sections[i + 1].Footer;
                        Sections[i].Header = Sections[i + 1].Header;
                        Sections[i]._sectionProperties = Sections[i + 1]._sectionProperties;

                        Sections[i + 1].Footer = element.Footer;
                        Sections[i + 1].Header = element.Header;
                        Sections[i + 1]._sectionProperties = element._sectionProperties;
                    } else {
                        Sections[i].Footer = firstElementFooter;
                        Sections[i].Header = firstElementHeader;
                        Sections[i]._sectionProperties = firstElementSection;
                    }
                }
            }
        }

        /// <summary>
        /// Load WordDocument from filePath
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="options">Access, persistence, style, and low-level package options.</param>
        /// <returns></returns>
        /// <exception cref="FileNotFoundException"></exception>
        public static WordDocument Load(string filePath, WordLoadOptions? options = null) {
            if (filePath is null) {
                throw new ArgumentNullException(nameof(filePath));
            }
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            WordLoadOptions resolved = options ?? new WordLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "document");
            bool readOnly = resolved.AccessMode == DocumentAccessMode.ReadOnly;
            bool saveOnDispose = resolved.PersistenceMode == DocumentPersistenceMode.SaveOnDispose;
            var word = new WordDocument { _persistenceMode = resolved.PersistenceMode };

            var effectiveOpenSettings = CreateOpenSettings(resolved.OpenSettings);

            // Read the source file into memory with a shared read handle to avoid test collisions.
            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) {
                byte[] sourceBytes = ReadSourceBytes(fileStream, resolved.PackageSecurity);
                ValidateSourcePackage(sourceBytes, resolved.PackageSecurity);

                if (WordDocumentLoadRouting.IsLegacyDoc(sourceBytes, filePath)) {
                    return LoadLegacyDocFromNormalFlow(sourceBytes, filePath, saveOnDispose, readOnly);
                }

                var memoryStream = new MemoryStream(sourceBytes.Length);
                memoryStream.Write(sourceBytes, 0, sourceBytes.Length);
                memoryStream.Position = 0;

                var wordDocument = WordprocessingDocument.Open(memoryStream, !readOnly, effectiveOpenSettings);

                bool applyOverrideStyles = resolved.OverrideStyles && !readOnly;
                InitialiseStyleDefinitions(wordDocument, readOnly, applyOverrideStyles);

                word.FilePath = filePath;
                word._ownedPackageStream = memoryStream;
                word._wordprocessingDocument = wordDocument;
                word._document = wordDocument.MainDocumentPart?.Document ?? throw new InvalidOperationException("Document is missing.");
                word._openXmlOriginalPackageBytes = OfficeCompatibilitySourceCarrier.ContainsPackageCarrier(sourceBytes)
                    ? (byte[])sourceBytes.Clone()
                    : null;
                word.LoadDocument();
                if (applyOverrideStyles) {
                    // Ensure overrides are applied after any document initialization that may touch styles
                    InitialiseStyleDefinitions(wordDocument, readOnly, applyOverrideStyles);
                    EnsureCustomStyleNames(wordDocument);
                }
                WordChart.InitializeAxisIdSeed(wordDocument);
                WordChart.InitializeDocPrIdSeed(wordDocument);

                // initialize abstract number id for lists to make sure those are unique
                WordListStyles.InitializeAbstractNumberId(word._wordprocessingDocument);
                return word;
            }
        }

        /// <summary>
        /// Loads a password-encrypted Office Open XML Word document.
        /// </summary>
        /// <param name="filePath">Path to the encrypted document.</param>
        /// <param name="password">Password used to decrypt the document package.</param>
        /// <param name="options">Access and load options. SaveOnDispose is not supported for encrypted sources.</param>
        /// <returns>Loaded <see cref="WordDocument"/> instance.</returns>
        public static WordDocument LoadEncrypted(string filePath, string password, WordLoadOptions? options = null) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (password == null) throw new ArgumentNullException(nameof(password));
            WordLoadOptions resolved = options ?? new WordLoadOptions();
            EnsureEncryptedLoadUsesExplicitPersistence(resolved);
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            byte[] encryptedBytes;
            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete)) {
                encryptedBytes = ReadSourceBytes(fileStream, resolved.PackageSecurity);
            }
            ValidateSourcePackage(encryptedBytes, resolved.PackageSecurity);
            byte[] packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            using var decryptedSource = new MemoryStream(packageBytes, writable: false);
            return Load(decryptedSource, resolved);
        }

        /// <summary>
        /// Loads a password-encrypted Office Open XML Word document from a stream.
        /// </summary>
        /// <param name="stream">Readable stream containing the encrypted document.</param>
        /// <param name="password">Password used to decrypt the document package.</param>
        /// <param name="options">Access and load options. SaveOnDispose is not supported for encrypted sources.</param>
        /// <returns>Loaded <see cref="WordDocument"/> instance.</returns>
        public static WordDocument LoadEncrypted(Stream stream, string password, WordLoadOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            WordLoadOptions resolved = options ?? new WordLoadOptions();
            EnsureEncryptedLoadUsesExplicitPersistence(resolved);

            byte[] encryptedBytes = ReadSourceBytes(stream, resolved.PackageSecurity);
            ValidateSourcePackage(encryptedBytes, resolved.PackageSecurity);
            byte[] packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            using var decryptedSource = new MemoryStream(packageBytes, writable: false);
            return Load(decryptedSource, resolved);
        }

        private static void EnsureEncryptedLoadUsesExplicitPersistence(WordLoadOptions options) {
            if (options.PersistenceMode != DocumentPersistenceMode.Explicit) {
                throw new NotSupportedException("SaveOnDispose is not supported for encrypted Word sources. Use SaveEncrypted to persist encrypted changes.");
            }
        }

        /// <summary>
        /// Asynchronously loads a <see cref="WordDocument"/> from the given file.
        /// </summary>
        /// <param name="filePath">Path to the file.</param>
        /// <param name="options">Access, persistence, style, and low-level package options.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Loaded <see cref="WordDocument"/> instance.</returns>
        /// <exception cref="FileNotFoundException">Thrown when the file does not exist.</exception>
        public static async Task<WordDocument> LoadAsync(string filePath, WordLoadOptions? options = null, CancellationToken cancellationToken = default) {
            if (filePath is null) {
                throw new ArgumentNullException(nameof(filePath));
            }
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            WordLoadOptions resolved = options ?? new WordLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "document");
            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete, 4096, FileOptions.Asynchronous);
            byte[] sourceBytes = await ReadSourceBytesAsync(fileStream, resolved.PackageSecurity,
                cancellationToken).ConfigureAwait(false);
            ValidateSourcePackage(sourceBytes, resolved.PackageSecurity);
            bool readOnly = resolved.AccessMode == DocumentAccessMode.ReadOnly;
            bool saveOnDispose = resolved.PersistenceMode == DocumentPersistenceMode.SaveOnDispose;
            var effectiveOpenSettings = CreateOpenSettings(resolved.OpenSettings);

            if (WordDocumentLoadRouting.IsLegacyDoc(sourceBytes, filePath)) {
                return LoadLegacyDocFromNormalFlow(sourceBytes, filePath, saveOnDispose, readOnly);
            }

            var memoryStream = new MemoryStream(sourceBytes.Length);
            memoryStream.Write(sourceBytes, 0, sourceBytes.Length);
            memoryStream.Position = 0;

            var wordDocument = WordprocessingDocument.Open(memoryStream, !readOnly, effectiveOpenSettings);

            var word = new WordDocument {
                FilePath = filePath,
                _ownedPackageStream = memoryStream,
                _wordprocessingDocument = wordDocument,
                _document = wordDocument.MainDocumentPart?.Document ?? throw new InvalidOperationException("Document is missing."),
                _openXmlOriginalPackageBytes = OfficeCompatibilitySourceCarrier.ContainsPackageCarrier(sourceBytes)
                    ? (byte[])sourceBytes.Clone()
                    : null,
                _persistenceMode = resolved.PersistenceMode
            };

            bool applyOverrideStyles = resolved.OverrideStyles && !readOnly;
            InitialiseStyleDefinitions(wordDocument, readOnly, applyOverrideStyles);
            word.LoadDocument();
            if (applyOverrideStyles) {
                InitialiseStyleDefinitions(wordDocument, readOnly, applyOverrideStyles);
                EnsureCustomStyleNames(wordDocument);
            }
            WordChart.InitializeAxisIdSeed(wordDocument);
            WordChart.InitializeDocPrIdSeed(wordDocument);
            WordListStyles.InitializeAbstractNumberId(word._wordprocessingDocument);
            return word;
        }

        /// <summary>Asynchronously loads a Word document from a readable stream.</summary>
        /// <param name="stream">Stream containing DOC or DOCX content.</param>
        /// <param name="options">Access, persistence, style, and low-level package options.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>The loaded document. The caller retains ownership of <paramref name="stream"/>.</returns>
        public static async Task<WordDocument> LoadAsync(
            Stream stream,
            WordLoadOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            WordLoadOptions resolved = options ?? new WordLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "document");
            bool readOnly = resolved.AccessMode == DocumentAccessMode.ReadOnly;
            bool copyBackToSource = resolved.PersistenceMode == DocumentPersistenceMode.SaveOnDispose && !readOnly;
            OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(stream, resolved.PersistenceMode, nameof(stream));

            byte[] sourceBytes = await ReadSourceBytesAsync(stream, resolved.PackageSecurity,
                cancellationToken).ConfigureAwait(false);
            using var bufferedStream = new MemoryStream(sourceBytes.Length);
            bufferedStream.Write(sourceBytes, 0, sourceBytes.Length);
            bufferedStream.Position = 0;
            WordDocument document = Load(bufferedStream, resolved);
            if (document.SourceFormat != WordFileFormat.Doc) {
                document.OriginalStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(stream, resolved.AccessMode)!;
            }

            return document;
        }

        /// <summary>
        /// Load WordDocument from stream
        /// </summary>
        /// <param name="stream">Readable source. Editable writable seekable sources become the associated destination; other sources remain detached.</param>
        /// <param name="options">Access, persistence, style, and low-level package options.</param>
        /// <returns></returns>
        public static WordDocument Load(Stream stream, WordLoadOptions? options = null) {
            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }
            if (!stream.CanRead) {
                throw new ArgumentException("Stream must be readable.", nameof(stream));
            }

            WordLoadOptions resolved = options ?? new WordLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "document");
            bool readOnly = resolved.AccessMode == DocumentAccessMode.ReadOnly;
            bool saveOnDispose = resolved.PersistenceMode == DocumentPersistenceMode.SaveOnDispose;
            OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(stream, resolved.PersistenceMode, nameof(stream));
            var effectiveOpenSettings = CreateOpenSettings(resolved.OpenSettings);
            byte[] sourceBytes = ReadSourceBytes(stream, resolved.PackageSecurity);
            ValidateSourcePackage(sourceBytes, resolved.PackageSecurity);

            if (WordDocumentLoadRouting.IsLegacyDoc(sourceBytes, filePath: null)) {
                return LoadLegacyDocFromNormalFlow(sourceBytes, sourcePath: null, saveOnDispose, readOnly);
            }

            var packageStream = new MemoryStream(sourceBytes.Length);
            packageStream.Write(sourceBytes, 0, sourceBytes.Length);
            packageStream.Position = 0;
            try {
                var document = new WordDocument() {
                    OriginalStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(stream, resolved.AccessMode)!,
                    _ownedPackageStream = packageStream,
                    _openXmlOriginalPackageBytes = OfficeCompatibilitySourceCarrier.ContainsPackageCarrier(sourceBytes)
                        ? (byte[])sourceBytes.Clone()
                        : null,
                    _persistenceMode = resolved.PersistenceMode
                };

                var wordDocument = WordprocessingDocument.Open(packageStream, !readOnly, effectiveOpenSettings);
                bool applyOverrideStyles = resolved.OverrideStyles && !readOnly;
                InitialiseStyleDefinitions(wordDocument, readOnly, applyOverrideStyles);

                document._wordprocessingDocument = wordDocument;
                document._document = wordDocument.MainDocumentPart?.Document ?? throw new InvalidOperationException("Document is missing.");
                document.LoadDocument();
                if (applyOverrideStyles) {
                    InitialiseStyleDefinitions(wordDocument, readOnly, applyOverrideStyles);
                    EnsureCustomStyleNames(wordDocument);
                }

                WordChart.InitializeAxisIdSeed(wordDocument);
                WordChart.InitializeDocPrIdSeed(wordDocument);
                WordListStyles.InitializeAbstractNumberId(document._wordprocessingDocument);
                return document;
            } catch {
                packageStream.Dispose();
                throw;
            }
        }

    }
}
