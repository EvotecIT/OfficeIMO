using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
#if NET5_0_OR_GREATER
using System.Runtime.Versioning;
#endif
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        private const string LegacyDocComValidationEnv = "OFFICEIMO_RUN_LEGACY_DOC_COM_VALIDATION";
        private const int WdDoNotSaveChanges = 0;
        private const int WdFormatDocument = 0;
        private const int WdStyleTypeParagraph = 1;
        private const int WdStyleHeading1 = -2;
        private const int WdStyleNormal = -1;
        private const int WdAlignParagraphLeft = 0;
        private const int WdAlignParagraphCenter = 1;
        private const int WdAlignParagraphRight = 2;
        private const int WdUnderlineNone = 0;
        private const int WdUnderlineSingle = 1;
        private const int WdColorAutomatic = -16777216;
        private const int WdColorRed = 255;
        private static readonly TimeSpan WordComOpenTimeout = TimeSpan.FromMinutes(2);

        [LegacyDocComFact]
        [Trait("Category", "MicrosoftOfficeInteroperability")]
        public void LegacyDoc_ComGeneratedDocument_ImportsAndNativeSaveOpensInDesktopWordWhenRequested() {
            Assert.True(IsWindowsPlatform(), "Legacy DOC COM validation requires Windows.");
            Assert.True(IsWordComAvailable(), "Legacy DOC COM validation requires Microsoft Word COM automation.");

            string directory = Path.Combine(_directoryWithFiles, "LegacyDocCom", GetCurrentTargetFrameworkLabel());
            Directory.CreateDirectory(directory);
            string sourceDocPath = Path.Combine(directory, "word-com-source.doc");
            string nativeDocPath = Path.Combine(directory, "officeimo-native.doc");
            string convertedDocxPath = Path.Combine(directory, "officeimo-converted.docx");

            CreateLegacyDocViaWordCom(sourceDocPath);
            AssertDocumentsOpenViaWordComWhenAvailable(new[] { sourceDocPath }, "The generated legacy DOC source did not open through desktop Word.");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(sourceDocPath);
            result.EnsureNoImportErrors();
            Assert.True(result.HasDocument);
            Assert.True(result.Document.SourceFormat == WordFileFormat.Doc);
            Assert.Contains(result.Document.Paragraphs, paragraph => paragraph.Text == "First Word COM paragraph");
            Assert.Contains(result.Document.Paragraphs, paragraph => paragraph.Text == "Centered Word COM paragraph" && paragraph.ParagraphAlignment == JustificationValues.Center);
            Assert.Contains(result.Document.Paragraphs, paragraph => paragraph.Text == "Right Word COM paragraph" && paragraph.ParagraphAlignment == JustificationValues.Right);
            Assert.Contains(result.Document.Paragraphs, paragraph => paragraph.Text == "bold Word COM paragraph" && paragraph.Bold);
            Assert.Contains(result.Document.Paragraphs, paragraph => paragraph.Text == "underlined red Word COM paragraph" && paragraph.Underline == UnderlineValues.Single);

            result.Document.Save(nativeDocPath);
            AssertDocumentsOpenViaWordComWhenAvailable(new[] { nativeDocPath }, "The OfficeIMO native DOC output did not open through desktop Word.");

            WordDocument.Convert(sourceDocPath, convertedDocxPath, new WordDocumentConversionOptions {
                LossPolicy = WordConversionLossPolicy.Allow
            });
            AssertDocumentsOpenViaWordComWhenAvailable(new[] { convertedDocxPath }, "The OfficeIMO converted DOCX output did not open through desktop Word.");
        }

        [LegacyDocComFact]
        public void LegacyDoc_NativeFootnoteSaveOpensInDesktopWordWhenRequested() {
            Assert.True(IsWindowsPlatform(), "Legacy DOC COM validation requires Windows.");
            Assert.True(IsWordComAvailable(), "Legacy DOC COM validation requires Microsoft Word COM automation.");

            string directory = Path.Combine(_directoryWithFiles, "LegacyDocCom", GetCurrentTargetFrameworkLabel());
            Directory.CreateDirectory(directory);
            string nativeDocPath = Path.Combine(directory, "officeimo-native-footnote.doc");
            DeleteIfExists(nativeDocPath);

            using (WordDocument document = WordDocument.Create()) {
                WordParagraph paragraph = document.AddParagraph("Body with desktop Word footnote");
                paragraph.AddFootNote("Desktop Word footnote");
                document.Save(nativeDocPath);
            }

            AssertDocumentsOpenViaWordComWhenAvailable(new[] { nativeDocPath }, "The OfficeIMO native legacy DOC footnote output did not open through desktop Word.");
        }

        [LegacyDocComFact]
        public void LegacyDoc_NativeEndnoteSaveOpensInDesktopWordWhenRequested() {
            Assert.True(IsWindowsPlatform(), "Legacy DOC COM validation requires Windows.");
            Assert.True(IsWordComAvailable(), "Legacy DOC COM validation requires Microsoft Word COM automation.");

            string directory = Path.Combine(_directoryWithFiles, "LegacyDocCom", GetCurrentTargetFrameworkLabel());
            Directory.CreateDirectory(directory);
            string nativeDocPath = Path.Combine(directory, "officeimo-native-endnote.doc");
            DeleteIfExists(nativeDocPath);

            using (WordDocument document = WordDocument.Create()) {
                WordParagraph paragraph = document.AddParagraph("Body with desktop Word endnote");
                paragraph.AddEndNote("Desktop Word endnote");
                document.Save(nativeDocPath);
            }

            AssertDocumentsOpenViaWordComWhenAvailable(new[] { nativeDocPath }, "The OfficeIMO native legacy DOC endnote output did not open through desktop Word.");
        }

        [LegacyDocComFact]
        [Trait("Category", "MicrosoftOfficeInteroperability")]
        public void LegacyDoc_ConvertGeneratedDocxToDocOpensInDesktopWordWhenRequested() {
            Assert.True(IsWindowsPlatform(), "Legacy DOC COM validation requires Windows.");
            Assert.True(IsWordComAvailable(), "Legacy DOC COM validation requires Microsoft Word COM automation.");

            string directory = Path.Combine(_directoryWithFiles, "LegacyDocCom", GetCurrentTargetFrameworkLabel());
            Directory.CreateDirectory(directory);
            string sourceDocxPath = Path.Combine(directory, "officeimo-source.docx");
            string directDocPath = Path.Combine(directory, "officeimo-source-direct.doc");
            string convertedDocPath = Path.Combine(directory, "officeimo-source-converted.doc");

            using (WordDocument document = WordDocument.Create()) {
                document.AddParagraph("OfficeIMO generated conversion paragraph");
                document.Save(sourceDocxPath);
                document.Save(directDocPath);
            }

            WordDocument.Convert(sourceDocxPath, convertedDocPath);

            AssertDocumentsOpenViaWordComWhenAvailable(new[] { directDocPath, convertedDocPath }, "One or more OfficeIMO generated native DOC outputs did not open through desktop Word.");
        }

        [LegacyDocComFact]
        public void LegacyDoc_ComGeneratedCustomParagraphStyle_ImportsStylesheetStyleWhenRequested() {
            Assert.True(IsWindowsPlatform(), "Legacy DOC COM validation requires Windows.");
            Assert.True(IsWordComAvailable(), "Legacy DOC COM validation requires Microsoft Word COM automation.");

            string directory = Path.Combine(_directoryWithFiles, "LegacyDocCom", GetCurrentTargetFrameworkLabel());
            Directory.CreateDirectory(directory);
            string sourceDocPath = Path.Combine(directory, "word-com-custom-style-source.doc");

            CreateLegacyDocWithCustomParagraphStyleViaWordCom(sourceDocPath);
            AssertDocumentsOpenViaWordComWhenAvailable(new[] { sourceDocPath }, "The generated custom-style legacy DOC source did not open through desktop Word.");

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(sourceDocPath);
            result.EnsureNoImportErrors();
            WordParagraph paragraph = result.Document.Paragraphs
                .First(item => item.Text == "Custom style Word COM paragraph");

            Assert.Equal(WordParagraphStyles.Custom, paragraph.Style);
            Assert.Equal("LegacyDocOfficeIMOCustomBody", paragraph.StyleId);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToBytes()));
            WordParagraph convertedParagraph = converted.Paragraphs
                .First(item => item.Text == "Custom style Word COM paragraph");
            Assert.Equal(WordParagraphStyles.Custom, convertedParagraph.Style);
            Assert.Equal("LegacyDocOfficeIMOCustomBody", convertedParagraph.StyleId);
            Assert.Contains(
                converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OfType<Style>(),
                style => style.StyleId?.Value == "LegacyDocOfficeIMOCustomBody" && style.StyleName?.Val?.Value == "OfficeIMO Custom Body");
            Style customStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OfType<Style>()
                .First(style => style.StyleId?.Value == "LegacyDocOfficeIMOCustomBody");
            Assert.Equal("Heading1", customStyle.BasedOn?.Val?.Value);
            StyleParagraphProperties paragraphProperties = Assert.IsType<StyleParagraphProperties>(customStyle.GetFirstChild<StyleParagraphProperties>());
            Assert.Equal(JustificationValues.Center, paragraphProperties.GetFirstChild<Justification>()?.Val?.Value);
            Assert.Equal("240", paragraphProperties.GetFirstChild<SpacingBetweenLines>()?.After?.Value);
            StyleRunProperties runProperties = Assert.IsType<StyleRunProperties>(customStyle.GetFirstChild<StyleRunProperties>());
            Assert.NotNull(runProperties.GetFirstChild<Bold>());
            Assert.Equal("28", runProperties.GetFirstChild<FontSize>()?.Val?.Value);
            Assert.Equal("FF0000", runProperties.GetFirstChild<Color>()?.Val?.Value);
            RunFonts runFonts = Assert.IsType<RunFonts>(runProperties.GetFirstChild<RunFonts>());
            Assert.Equal("Courier New", runFonts.Ascii?.Value);
            Assert.Equal("Courier New", runFonts.HighAnsi?.Value);
        }

        [LegacyDocComFact]
        [Trait("Category", "MicrosoftOfficeInteroperability")]
        public void LegacyDoc_CorpusFixtures_OpenInDesktopWordWhenRequested() {
            Assert.True(IsWindowsPlatform(), "Legacy DOC COM validation requires Windows.");
            Assert.True(IsWordComAvailable(), "Legacy DOC COM validation requires Microsoft Word COM automation.");

            string corpusDirectory = Path.Combine(GetWordTestsProjectRoot(), "Documents", "LegacyDocCorpus");
            Assert.True(Directory.Exists(corpusDirectory), $"Legacy DOC corpus directory '{corpusDirectory}' was not found.");

            string[] documentPaths = Directory.GetFiles(corpusDirectory, "*.doc", SearchOption.AllDirectories)
                .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            Assert.NotEmpty(documentPaths);

            AssertDocumentsOpenViaWordComWhenAvailable(documentPaths, "One or more legacy DOC corpus fixtures failed to open in desktop Word.");
        }

        private static bool IsLegacyDocComValidationRequested() {
            string? value = Environment.GetEnvironmentVariable(LegacyDocComValidationEnv);
            return string.Equals(value, "1", StringComparison.Ordinal)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

#if NET5_0_OR_GREATER
        [SupportedOSPlatformGuard("windows")]
#endif
        private static bool IsWindowsPlatform() =>
            RuntimeInformation.IsOSPlatform(OSPlatform.Windows);

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static bool IsWordComAvailable() =>
            Type.GetTypeFromProgID("Word.Application") != null;

        private static string GetCurrentTargetFrameworkLabel() {
#if NET472
            return "net472";
#elif NET10_0_OR_GREATER
            return "net10";
#elif NET8_0_OR_GREATER
            return "net8";
#else
            string frameworkName = typeof(Word).Assembly
                .GetCustomAttribute<System.Runtime.Versioning.TargetFrameworkAttribute>()?
                .FrameworkName ?? "unknown";

            var builder = new System.Text.StringBuilder(frameworkName.Length);
            foreach (char character in frameworkName) {
                builder.Append(char.IsLetterOrDigit(character) ? character : '_');
            }

            return builder.ToString().Trim('_');
#endif
        }

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static void CreateLegacyDocViaWordCom(string path) =>
            CreateLegacyDocViaWordCom(path, CreateLegacyDocViaWordComOnStaThread);

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static void CreateLegacyDocWithCustomParagraphStyleViaWordCom(string path) =>
            CreateLegacyDocViaWordCom(path, CreateLegacyDocWithCustomParagraphStyleViaWordComOnStaThread);

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static void CreateLegacyDocViaWordCom(string path, Action<string> createDocument) {
            var failures = new List<string>();
            var thread = new Thread(() => {
                try {
                    createDocument(path);
                } catch (Exception ex) when (ex is COMException or InvalidOperationException or MissingMethodException or TargetInvocationException) {
                    failures.Add(DescribeWordComFailure(ex));
                }
            });

            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            if (!thread.Join(WordComOpenTimeout)) {
                failures.Add($"Word COM legacy DOC generation timed out after {WordComOpenTimeout.TotalSeconds:0} seconds.");
            }

            Assert.True(failures.Count == 0, "Failed to generate the legacy DOC document through desktop Word." + Environment.NewLine + string.Join(Environment.NewLine, failures));
        }

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static void AssertDocumentsOpenViaWordComWhenAvailable(IEnumerable<string> paths, string failureMessage) {
            if (!IsWordComAvailable()) {
                return;
            }

            List<string> failures = new();
            var thread = new Thread(() => OpenDocumentsViaWordCom(paths.ToList(), failures));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            if (!thread.Join(WordComOpenTimeout)) {
                failures.Add($"Word COM smoke test timed out after {WordComOpenTimeout.TotalSeconds:0} seconds.");
            }

            Assert.True(failures.Count == 0, failureMessage + Environment.NewLine + string.Join(Environment.NewLine, failures));
        }

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static void CreateLegacyDocViaWordComOnStaThread(string path) {
            object? word = null;
            object? documents = null;
            object? document = null;
            object? selection = null;

            try {
                word = CreateWordComApplication();
                documents = GetComProperty(word, "Documents");
                document = InvokeCom(documents!, "Add");
                selection = GetComProperty(word, "Selection");

                AddWordComParagraph(selection!, "First Word COM paragraph");
                AddWordComParagraph(selection!, "Heading Word COM paragraph", selectionItem => {
                    SetComProperty(selectionItem, "Style", WdStyleHeading1);
                });
                AddWordComParagraph(selection!, "Centered Word COM paragraph", selectionItem => {
                    object? paragraphFormat = GetComProperty(selectionItem, "ParagraphFormat");
                    SetComProperty(paragraphFormat!, "Alignment", WdAlignParagraphCenter);
                    ReleaseComObject(paragraphFormat);
                });
                AddWordComParagraph(selection!, "Right Word COM paragraph", selectionItem => {
                    object? paragraphFormat = GetComProperty(selectionItem, "ParagraphFormat");
                    SetComProperty(paragraphFormat!, "Alignment", WdAlignParagraphRight);
                    ReleaseComObject(paragraphFormat);
                });
                AddWordComParagraph(selection!, "bold Word COM paragraph", selectionItem => {
                    object? font = GetComProperty(selectionItem, "Font");
                    SetComProperty(font!, "Bold", 1);
                    ReleaseComObject(font);
                });
                AddWordComParagraph(selection!, "underlined red Word COM paragraph", selectionItem => {
                    object? font = GetComProperty(selectionItem, "Font");
                    SetComProperty(font!, "Underline", WdUnderlineSingle);
                    SetComProperty(font!, "Color", WdColorRed);
                    SetComProperty(font!, "Name", "Courier New");
                    SetComProperty(font!, "Size", 14);
                    ReleaseComObject(font);
                });

                if (File.Exists(path)) {
                    File.Delete(path);
                }

                InvokeCom(document!, "SaveAs2", path, WdFormatDocument);
            } finally {
                CloseWordComDocument(document);
                QuitWordComApplication(word);
                ReleaseComObject(selection);
                ReleaseComObject(document);
                ReleaseComObject(documents);
                ReleaseComObject(word);
            }
        }

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static void CreateLegacyDocWithCustomParagraphStyleViaWordComOnStaThread(string path) {
            object? word = null;
            object? documents = null;
            object? document = null;
            object? selection = null;
            object? styles = null;
            object? customStyle = null;

            try {
                word = CreateWordComApplication();
                documents = GetComProperty(word, "Documents");
                document = InvokeCom(documents!, "Add");
                selection = GetComProperty(word, "Selection");
                styles = GetComProperty(document!, "Styles");
                customStyle = InvokeCom(styles!, "Add", "OfficeIMO Custom Body", WdStyleTypeParagraph);
                SetComProperty(customStyle!, "BaseStyle", WdStyleHeading1);
                object? paragraphFormat = null;
                object? font = null;
                try {
                    paragraphFormat = GetComProperty(customStyle!, "ParagraphFormat");
                    SetComProperty(paragraphFormat!, "Alignment", WdAlignParagraphCenter);
                    SetComProperty(paragraphFormat!, "SpaceAfter", 12);
                    font = GetComProperty(customStyle!, "Font");
                    SetComProperty(font!, "Bold", 1);
                    SetComProperty(font!, "Color", WdColorRed);
                    SetComProperty(font!, "Name", "Courier New");
                    SetComProperty(font!, "Size", 14);
                } finally {
                    ReleaseComObject(font);
                    ReleaseComObject(paragraphFormat);
                }

                AddWordComParagraph(selection!, "Custom style Word COM paragraph", selectionItem => {
                    SetComProperty(selectionItem, "Style", "OfficeIMO Custom Body");
                });
                AddWordComParagraph(selection!, "Normal after custom style");

                if (File.Exists(path)) {
                    File.Delete(path);
                }

                InvokeCom(document!, "SaveAs2", path, WdFormatDocument);
            } finally {
                CloseWordComDocument(document);
                QuitWordComApplication(word);
                ReleaseComObject(customStyle);
                ReleaseComObject(styles);
                ReleaseComObject(selection);
                ReleaseComObject(document);
                ReleaseComObject(documents);
                ReleaseComObject(word);
            }
        }

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static void OpenDocumentsViaWordCom(IReadOnlyList<string> paths, List<string> failures) {
            object? word = null;
            object? documents = null;

            try {
                word = CreateWordComApplication();
                documents = GetComProperty(word, "Documents");

                foreach (string path in paths) {
                    object? document = null;
                    try {
                        document = InvokeCom(documents!, "Open", path, false, true, false);
                    } catch (Exception ex) when (ex is COMException or InvalidOperationException or MissingMethodException or TargetInvocationException) {
                        failures.Add($"{Path.GetFileName(path)}: {DescribeWordComFailure(ex)}");
                    } finally {
                        try {
                            CloseWordComDocument(document);
                        } catch (Exception ex) when (ex is COMException or MissingMethodException or TargetInvocationException) {
                            failures.Add($"{Path.GetFileName(path)} close: {DescribeWordComFailure(ex)}");
                        }

                        ReleaseComObject(document);
                    }
                }
            } catch (Exception ex) when (ex is COMException or InvalidOperationException or MissingMethodException or TargetInvocationException) {
                failures.Add(DescribeWordComFailure(ex));
            } finally {
                QuitWordComApplication(word);
                ReleaseComObject(documents);
                ReleaseComObject(word);
            }
        }

        private static object CreateWordComApplication() {
            var wordType = Type.GetTypeFromProgID("Word.Application")
                ?? throw new InvalidOperationException("Word COM automation is not available.");
            object word = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Failed to create Word COM automation instance.");

            SetComProperty(word, "DisplayAlerts", 0);
            SetComProperty(word, "Visible", false);
            return word;
        }

        private static void AddWordComParagraph(object selection, string text, Action<object>? configure = null) {
            ResetWordComSelectionFormat(selection);
            configure?.Invoke(selection);
            InvokeCom(selection, "TypeText", text);
            InvokeCom(selection, "TypeParagraph");
            ResetWordComSelectionFormat(selection);
        }

        private static void ResetWordComSelectionFormat(object selection) {
            object? font = null;
            object? paragraphFormat = null;
            try {
                SetComProperty(selection, "Style", WdStyleNormal);
                font = GetComProperty(selection, "Font");
                SetComProperty(font!, "Bold", 0);
                SetComProperty(font!, "Italic", 0);
                SetComProperty(font!, "Underline", WdUnderlineNone);
                SetComProperty(font!, "Size", 11);
                SetComProperty(font!, "Name", "Calibri");
                SetComProperty(font!, "Color", WdColorAutomatic);
                paragraphFormat = GetComProperty(selection, "ParagraphFormat");
                SetComProperty(paragraphFormat!, "Alignment", WdAlignParagraphLeft);
                SetComProperty(paragraphFormat!, "SpaceBefore", 0);
                SetComProperty(paragraphFormat!, "SpaceAfter", 0);
                SetComProperty(paragraphFormat!, "LineSpacing", 12);
                SetComProperty(paragraphFormat!, "LeftIndent", 0);
                SetComProperty(paragraphFormat!, "RightIndent", 0);
                SetComProperty(paragraphFormat!, "FirstLineIndent", 0);
            } finally {
                ReleaseComObject(paragraphFormat);
                ReleaseComObject(font);
            }
        }

        private static object? GetComProperty(object comObject, string name, params object[] arguments) =>
            comObject.GetType().InvokeMember(name, BindingFlags.GetProperty, null, comObject, arguments.Length == 0 ? null : arguments);

        private static void SetComProperty(object comObject, string name, object value) =>
            comObject.GetType().InvokeMember(name, BindingFlags.SetProperty, null, comObject, new[] { value });

        private static object? InvokeCom(object comObject, string name, params object[] arguments) =>
            comObject.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, comObject, arguments.Length == 0 ? null : arguments);

        private static void CloseWordComDocument(object? document) {
            if (document == null) {
                return;
            }

            InvokeCom(document, "Close", WdDoNotSaveChanges);
        }

        private static void QuitWordComApplication(object? word) {
            if (word == null) {
                return;
            }

            InvokeCom(word, "Quit", WdDoNotSaveChanges);
        }

        private static void ReleaseComObject(object? comObject) {
            if (comObject != null && Marshal.IsComObject(comObject)) {
                Marshal.FinalReleaseComObject(comObject);
            }
        }

        private static string DescribeWordComFailure(Exception ex) {
            Exception actual = ex is TargetInvocationException { InnerException: not null } target
                ? target.InnerException!
                : ex;
            return actual.GetType().Name + ": " + actual.Message;
        }
    }
}
