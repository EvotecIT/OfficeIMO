using System;
using System.IO;

namespace OfficeIMO.Examples {
    internal static class Program {
        private static void Setup(string path) {
            if (!Directory.Exists(path)) {
                Directory.CreateDirectory(path);
            }
        }

        static void Main(string[] args) {
            string baseFolder = Path.TrimEndingDirectorySeparator(AppContext.BaseDirectory);
            Directory.SetCurrentDirectory(baseFolder);
            string templatesPath = Path.Combine(baseFolder, "Templates");
            string folderPath = Path.Combine(baseFolder, "Documents");
            Setup(folderPath);

            // // Visio - Core Examples
            // Visio.BasicVisioDocument.Example_BasicVisio(folderPath, false);
            // Visio.ConnectRectangles.Example_ConnectRectangles(folderPath, false);
            // Visio.ConnectionPoints.Example_ConnectionPoints(folderPath, false);
            // Visio.ComprehensiveColoredShapes.Example_ComprehensiveColoredShapes(folderPath, false);
            // Visio.ReadVisioDocument.Example_ReadVisio(folderPath, false);

            // // Excel/BasicExcelFunctionality
            // Excel.BasicExcelFunctionality.BasicExcel_Example1(folderPath, false);
            // Excel.BasicExcelFunctionality.BasicExcel_Example2(folderPath, false);
            // Excel.BasicExcelFunctionality.BasicExcel_Example3(false);
            // // Excel/BasicExcelFunctionalityAsync
            // Excel.BasicExcelFunctionalityAsync.Example_ExcelAsync(folderPath).GetAwaiter().GetResult();
            // // Excel/AutoFit
            // Excel.AutoFit.Example(folderPath, false);
            // // Excel/AddTable
            // Excel.AddTable.Example(folderPath, false);
            // // Excel/AddTableMissingCells
            // Excel.AddTableMissingCells.Example(folderPath, false);
            // // Excel/AutoFilter
            // Excel.AutoFilter.Example(folderPath, false);
            // // Excel/Styles & Colors
            // Excel.StylesColors.Example(folderPath, false);
            // // Excel/Freeze
            // Excel.Freeze.Example(folderPath, false);
            // // Excel/ConditionalFormatting
            // Excel.ConditionalFormatting.Example(folderPath, false);
            // // Excel/ConcurrentWrites
            // Excel.ConcurrentWrites.Example(folderPath, false);
            // // Excel/ExcelConcurrentAccessAsync
            // Excel.ExcelConcurrentAccessAsync.Example_ExcelAsyncConcurrent(folderPath).GetAwaiter().GetResult();
            // // Excel/CellValues
            // Excel.CellValues.Example(folderPath, false);
            // // Excel/CellValuesParallel
            // Excel.CellValuesParallel.Example(folderPath, false);
            // // Excel/ValidateDocument
            // Excel.ValidateDocument.Example(folderPath, false);
            // // Excel/Fluent
            // Excel.FluentWorkbook.Example_FluentWorkbook(folderPath, false);
            // Excel.FluentWorkbook.Example_RangeBuilder(folderPath, false);
            // Excel.FluentWorkbook.Example_FluentWorkbook_AutoFilter(folderPath, false);
            // Excel.TestDateTime.Example_TestDateTime(folderPath, false);
            // // Excel/Read with custom converters
            // Excel.ReadWithConverters.Example(folderPath, false);
            // // Excel/Read presets and static helpers
            // Excel.ReadPresetsAndHelpers.Example(folderPath, false);
            // // Excel/Read for PowerShell consumption (emits JSON rows)
            // Excel.ReadForPowerShell.Example(folderPath, false);
            // // Excel/PowerShell-style round trip: write → read → modify → write → JSON
            // Excel.PowerShellRoundTrip.Example(folderPath, false);
            // // Excel/Headers + Footers + Properties
            // Excel.HeadersFootersAndProperties.Example(folderPath, false);
            // Excel.DomainDetectiveReport.Example(folderPath, false);
            // // Excel: New Excelish Sheets demo (side-by-side comparison)
            // Excel.DomainDetectiveReportSheets.Example(folderPath, false);
            // // Excel: Classic baseline Sheets demo (explicit/standard techniques)
            // Excel.DomainDetectiveReportSheetsClassic.Example(folderPath, false);
            // // Excel: Anchors and back-to-top demo
            // Excel.AnchorsAndBackToTop.Example(folderPath, false);
            // // Excel: Left-to-right multiple tables on same sheet
            // Excel.SheetComposerMultiTables.Example_LeftToRight(folderPath, false);
            // Excel.WrapText.Example(folderPath, false);
            // Excel.RowsFromObjects.Example(folderPath, false);
            // Excel.RowsFromObjectsPriorityProperties.Example(folderPath, false);
            // Excel.DomainDetectiveReportSheets.Example(folderPath, false);
            // // // Markdown: Anchors + Theme Toggle
            // Markdown.Markdown03_Anchors_Theme.Example_AnchorsAndTheme(folderPath, false);
            // // Markdown: TOC Layouts & Themes
            // Markdown.Markdown04_TocLayoutsAndThemes.Example_Toc_PanelTop(folderPath, false);
            // Markdown.Markdown04_TocLayoutsAndThemes.Example_Toc_SidebarLeft(folderPath, false);
            // Markdown.Markdown04_TocLayoutsAndThemes.Example_Toc_SidebarRight_ScrollSpy(folderPath, false);
            // Markdown.Markdown04_TocLayoutsAndThemes.Example_Toc_ScrollSpy_Long_IndigoTheme(folderPath, false);
            // // Markdown: Built-in HTML style gallery
            // Markdown.Markdown05_ThemesGallery.Example_Themes(folderPath, false);
            // // Word ⇄ Markdown ⇄ HTML End-to-End
            // Word.EndToEnd.Word_EndToEnd.Example(folderPath, false);
            // // Markdown/DomainDetective report (mirrors the Excel structure)
            // Markdown.DomainDetectiveReportMarkdown.Example(folderPath, false);
            // // PowerPoint
            // PowerPoint.BasicPowerPointDocument.Example_BasicPowerPoint(folderPath, false);
            // PowerPoint.AdvancedPowerPoint.Example_AdvancedPowerPoint(folderPath, false);
            // PowerPoint.FluentPowerPoint.Example_FluentPowerPoint(folderPath, false);
            // PowerPoint.ShapesPowerPoint.Example_PowerPointShapes(folderPath, false);
            // PowerPoint.SlidesManagementPowerPoint.Example_SlidesManagement(folderPath, false);
            // PowerPoint.TablesPowerPoint.Example_PowerPointTables(folderPath, false);
            // PowerPoint.TextFormattingPowerPoint.Example_TextFormattingPowerPoint(folderPath, false);
            // PowerPoint.ThemeAndLayoutPowerPoint.Example_PowerPointThemeAndLayout(folderPath, false);
            // PowerPoint.UpdatePicturePowerPoint.Example_PowerPointUpdatePicture(folderPath, false);
            // PowerPoint.ValidateDocument.Example(folderPath, false);
            // PowerPoint.TestLazyInit.Example_TestLazyInit(folderPath, false);
            // // Html/Html (consolidated set)
            // Html.Html.Example_Html01_LoadAndRoundTripBasics(folderPath, false);
            // Html.Html.Example_Html02_SaveAsHtmlFromWord(folderPath, false);
            // Html.Html.Example_Html03_TextFormatting(folderPath, false);
            // Html.Html.Example_Html04_ListsAndNumbering(folderPath, false);
            // Html.Html.Example_Html05_TablesComplex(folderPath, false);
            // Html.Html.Example_Html06_ImagesAllModes(folderPath, false);
            // Html.Html.Example_Html07_LinksAndAnchors(folderPath, false);
            // Html.Html.Example_Html08_SemanticsAndCitations(folderPath, false);
            // Html.Html.Example_Html09_CodePreWhitespace(folderPath, false);
            // Html.Html.Example_Html10_OptionsAndAsync(folderPath, false).GetAwaiter().GetResult();
            // Html.Html.Example_Html00_AllInOne(folderPath, false);
            // PDF: Zero-dependency writer + reader examples
            Pdf.BasicPdf.Example_Pdf_HelloWorld(folderPath, true);
            Pdf.ReadPdf.Example_Pdf_ReadPlainText(folderPath, true);
            Pdf.ReadSpans.Example_Pdf_TextSpans(folderPath, true);
            Pdf.WriterHeadersFooters.Example_Pdf_PageNumbers(folderPath, true);
            Pdf.WriterListsTables.Example_Pdf_BulletsAndTable(folderPath, true);
            Pdf.ReadDocumentText.Example_Pdf_ReadDocumentText(folderPath, true);
            return;

            // Markdown/Markdown
            Markdown.Markdown.Example_MarkdownInterface(folderPath, false);
            Markdown.Markdown.Example_MarkdownLists(folderPath, false);
            Markdown.Markdown.Example_MarkdownRoundTrip(folderPath, false);
            Markdown.Markdown.Example_MarkdownFootNotes(folderPath, false);
            Markdown.Markdown.Example_MarkdownHeadingsBoldLinks(folderPath, false);
            // Markdown/Builder & TOC
            Markdown.Markdown01_Builder_Basics.Example_Builder_Readme(folderPath, false);
            Markdown.Markdown01_Builder_Basics.Example_Scaffold_Readme(folderPath, false);
            Markdown.Markdown02_DataToTableAndLists.Example_TablesAndLists(folderPath, false);
            Markdown.Markdown02_DataToTableAndLists.Example_Toc(folderPath, false);
            Markdown.Markdown02_DataToTableAndLists.Example_Table_FromAny_WithOptions(folderPath, false);
            Markdown.Markdown02_DataToTableAndLists.Example_Table_FromSequence_WithSelectors(folderPath, false);
            Markdown.Markdown02_DataToTableAndLists.Example_HeaderTransform_CustomAcronyms(folderPath, false);
            Markdown.Markdown02_DataToTableAndLists.Example_Table_AutoAligners(folderPath, false);
            Markdown.Markdown02_DataToTableAndLists.Example_TocForSection(folderPath, false);
            // Word/AdvancedDocument
            Word.AdvancedDocument.Example_AdvancedWord(folderPath, false);
            Word.AdvancedDocument.Example_AdvancedWord2(folderPath, false);
            // Word/Background
            Word.Background.Example_BackgroundImageAdvanced(folderPath, false);
            Word.Background.Example_BackgroundImageSimple(folderPath, false);
            // Word/BasicDocument
            Word.BasicDocument.Example_BasicDocument(folderPath, false);
            Word.BasicDocument.Example_BasicDocumentSaveAs1(folderPath, false);
            Word.BasicDocument.Example_BasicDocumentSaveAs2(folderPath, false);
            Word.BasicDocument.Example_BasicDocumentSaveAs3(folderPath, false);
            Word.BasicDocument.Example_BasicDocumentWithoutUsing(folderPath, false);
            Word.BasicDocument.Example_BasicEmptyWord(folderPath, false);
            Word.BasicDocument.Example_BasicLoadHamlet(templatesPath, folderPath, false);
            Word.BasicDocument.Example_BasicWord(folderPath, false);
            Word.BasicDocument.Example_BasicWord2(folderPath, false);
            Word.BasicDocument.Example_BasicWordAsync(folderPath).GetAwaiter().GetResult();
            Word.BasicDocument.Example_BasicWordWithBreaks(folderPath, false);
            Word.BasicDocument.Example_BasicWordWithDefaultFontChange(folderPath, false);
            Word.BasicDocument.Example_BasicWordWithDefaultStyleChange(folderPath, false);
            Word.BasicDocument.Example_BasicWordWithLineSpacing(folderPath, false);
            Word.BasicDocument.Example_BasicWordWithMargins(folderPath, false);
            Word.BasicDocument.Example_BasicWordWithMarginsAndImage(folderPath, false);
            Word.BasicDocument.Example_BasicWordWithMarginsInCentimeters(folderPath, false);
            Word.BasicDocument.Example_BasicWordWithNewLines(folderPath, false);
            Word.BasicDocument.Example_BasicWordWithPolishChars(folderPath, false);
            Word.BasicDocument.Example_BasicWordWithSomeParagraphs(folderPath, false);
            Word.BasicDocument.Example_BasicWordWithTabs(folderPath, false);
            // Word/Bookmarks
            Word.Bookmarks.Example_BasicWordWithBookmarks(folderPath, false);
            // Word/BordersAndMargins
            Word.BordersAndMargins.Example_BasicPageBorders1(folderPath, false);
            Word.BordersAndMargins.Example_BasicPageBorders2(folderPath, false);
            Word.BordersAndMargins.Example_BasicWordMarginsSizes(folderPath, false);
            // Word/Charts
            Word.Charts.Example_AddingMultipleCharts(folderPath, false);
            Word.Charts.Example_Area3DChart(folderPath, false);
            Word.Charts.Example_AreaChart(folderPath, false);
            Word.Charts.Example_Bar3DChart(folderPath, false);
            Word.Charts.Example_BarChart(folderPath, false);
            Word.Charts.Example_ComboChart(folderPath, false);
            Word.Charts.Example_Line3DChart(folderPath, false);
            Word.Charts.Example_LineChart(folderPath, false);
            Word.Charts.Example_Pie3DChart(folderPath, false);
            Word.Charts.Example_PieChart(folderPath, false);
            Word.Charts.Example_RadarChart(folderPath, false);
            Word.Charts.Example_ScatterChart(folderPath, false);
            // Word/CheckBoxes
            Word.CheckBoxes.Example_BasicCheckBox(folderPath, false);
            // Word/CitationsExamples
            Word.CitationsExamples.Example_AdvancedCitations(folderPath, false);
            Word.CitationsExamples.Example_BasicCitations(folderPath, false);
            // Word/ComboBoxes
            Word.ComboBoxes.Example_BasicComboBox(folderPath, false);
            // Word/Comments
            Word.Comments.Example_PlayingWithComments(folderPath, false);
            Word.Comments.Example_RemoveCommentsAndTrack(folderPath, false);
            Word.Comments.Example_ThreadedComments(folderPath, false);
            // Word/CompareDocuments
            Word.CompareDocuments.Example_BasicComparison(folderPath, false);
            // Word/ContentControls
            Word.ContentControls.Example_AddContentControl(folderPath, false);
            Word.ContentControls.Example_AdvancedContentControls(folderPath, false);
            Word.ContentControls.Example_ContentControlsInTable(folderPath, false);
            Word.ContentControls.Example_MultipleContentControls(folderPath, false);
            // Word/CleanupDocuments
            Word.CleanupDocuments.CleanupDocuments_Sample01(false);
            Word.CleanupDocuments.CleanupDocuments_Sample02(folderPath, false);
            Word.CleanupDocuments.CleanupDocuments_Sample03(folderPath, false);
            Word.CleanupDocuments.CleanupDocuments_Sample04(folderPath, false);
            // Word/CoverPages
            Word.CoverPages.Example_AddingCoverPage(folderPath, false);
            Word.CoverPages.Example_AddingCoverPage2(folderPath, false);
            // Word/CrossReferences
            Word.CrossReferences.Example_BasicCrossReferences(folderPath, false);
            // Word/CustomAndBuiltinProperties
            Word.CustomAndBuiltinProperties.Example_BasicCustomProperties(folderPath, false);
            Word.CustomAndBuiltinProperties.Example_BasicDocumentProperties(folderPath, false);
            Word.CustomAndBuiltinProperties.Example_Load(false);
            Word.CustomAndBuiltinProperties.Example_LoadDocumentWithProperties(false);
            Word.CustomAndBuiltinProperties.Example_ReadWord(false);
            Word.CustomAndBuiltinProperties.Example_ValidateDocument(folderPath);
            Word.CustomAndBuiltinProperties.Example_ValidateDocument_BeforeSave();
            // Word/DatePickers
            Word.DatePickers.Example_BasicDatePicker(folderPath, false);
            Word.DatePickers.Example_AdvancedDatePicker(folderPath, false);
            // Word/DocumentVariablesExamples
            Word.DocumentVariablesExamples.Example_AdvancedDocumentVariables(folderPath, false);
            Word.DocumentVariablesExamples.Example_BasicDocumentVariables(folderPath, false);
            // Word/DropDownLists
            Word.DropDownLists.Example_BasicDropDownList(folderPath, false);
            Word.DropDownLists.Example_AdvancedDropDownList(folderPath, false);
            // Word/Embed
            Word.Embed.Example_EmbedFileExcel(folderPath, templatesPath, false);
            Word.Embed.Example_EmbedFileHTML(folderPath, templatesPath, false);
            Word.Embed.Example_EmbedFileMultiple(folderPath, templatesPath, false);
            Word.Embed.Example_EmbedFileRTF(folderPath, templatesPath, false);
            Word.Embed.Example_EmbedFileRTFandHTML(folderPath, templatesPath, false);
            Word.Embed.Example_EmbedFileRTFandHTMLandTOC(folderPath, templatesPath, false);
            Word.Embed.Example_EmbedFragmentAfter(folderPath, false);
            Word.Embed.Example_EmbedHTMLFragment(folderPath, false);
            // Word/Equations
            Word.Equations.Example_AddEquation(folderPath, false);
            Word.Equations.Example_AddEquationExponent(folderPath, false);
            Word.Equations.Example_AddEquationIntegral(folderPath, false);
            // Word/Fields
            Word.Fields.Example_CustomFormattedDateField(folderPath, false);
            Word.Fields.Example_CustomFormattedHeaderDate(folderPath, false);
            Word.Fields.Example_CustomFormattedTimeField(folderPath, false);
            Word.Fields.Example_DocumentWithFields(folderPath, false);
            Word.Fields.Example_DocumentWithFields02(folderPath, false);
            Word.Fields.Example_FieldBuilderNested(folderPath, false);
            //OfficeIMO.Examples.Word.Fields.Example_FieldBuilderSimple(folderPath, false);
            Word.Fields.Example_FieldFormatAdvanced(folderPath, false);
            Word.Fields.Example_FieldFormatRoman(folderPath, false);
            Word.Fields.Example_FieldWithMultipleSwitches(folderPath, false);
            // Word/FindAndReplace
            Word.FindAndReplace.Example_FindAndReplace01(folderPath, false);
            Word.FindAndReplace.Example_FindAndReplace02(folderPath, false);
            Word.FindAndReplace.Example_ReplaceTextWithHtmlFragment(folderPath, false);
            // Word/Fluent
            Word.FluentDocument.Example_FluentDocument(folderPath, false);
            Word.FluentDocument.Example_FluentHeadersAndFooters(folderPath, false);
            Word.FluentDocument.Example_FluentListBuilder(folderPath, false);
            Word.FluentDocument.Example_FluentParagraphFormatting(folderPath, false);
            Word.FluentDocument.Example_FluentReadHelpers(folderPath, false);
            Word.FluentDocument.Example_FluentSectionLayout(folderPath, false);
            Word.FluentDocument.Example_FluentTableBuilder(folderPath, false);
            Word.FluentDocument.Example_FluentTextBuilder(folderPath, false);
            // Word/Fonts
            Word.Fonts.Example_EmbeddedAndBuiltinFonts(templatesPath, folderPath, false);
            Word.Fonts.Example_EmbeddedFontStyle(templatesPath, folderPath, false);
            Word.Fonts.Example_EmbedFont(templatesPath, folderPath, false);
            Word.Fonts.Example_EmbedFontWithStyle(templatesPath, folderPath, false);
            // Word/FootNotes
            Word.FootNotes.Example_DocumentWithFootNotes(folderPath, false);
            Word.FootNotes.Example_DocumentWithFootNotesEmpty(folderPath, false);
            // Word/HeadersAndFooters
            Word.HeadersAndFooters.Example_BasicWordWithHeaderAndFooter(folderPath, false);
            Word.HeadersAndFooters.Example_BasicWordWithHeaderAndFooter0(folderPath, false);
            Word.HeadersAndFooters.Example_BasicWordWithHeaderAndFooter1(folderPath, false);
            Word.HeadersAndFooters.Example_BasicWordWithHeaderAndFooterWithoutSections(folderPath, false);
            Word.HeadersAndFooters.Sections1(folderPath, false);
            // Word/HyperLinks
            Word.HyperLinks.EasyExample(folderPath, false);
            Word.HyperLinks.Example_AddingFields(folderPath, false);
            Word.HyperLinks.Example_BasicWordWithHyperLinks(folderPath, false);
            Word.HyperLinks.Example_BasicWordWithHyperLinksInTables(folderPath, false);
            Word.HyperLinks.Example_FormattedHyperLinks(folderPath, false);
            Word.HyperLinks.Example_FormattedHyperLinksAdvanced(folderPath, false);
            Word.HyperLinks.Example_FormattedHyperLinksListReuse(folderPath, false);
            // Word/Images
            Word.Images.Example_AddingFixedImages(folderPath, false);
            Word.Images.Example_AddingImages(folderPath, false);
            Word.Images.Example_AddingImagesHeadersFooters(folderPath, false);
            Word.Images.Example_AddingImagesInline(folderPath, false);
            Word.Images.Example_AddingImagesMultipleTypes(folderPath, false);
            Word.Images.Example_AddingImagesSample4(folderPath, false);
            Word.Images.Example_AddingImagesSampleToTable(folderPath, false);
            Word.Images.Example_ImageCroppingAdvanced(folderPath, false);
            Word.Images.Example_ImageCroppingBasic(folderPath, false);
            Word.Images.Example_ImageNewFeatures(folderPath, false);
            Word.Images.Example_ImageTransparencyAdvanced(folderPath, false);
            Word.Images.Example_ImageTransparencySimple(folderPath, false);
            Word.Images.Example_ReadWordWithImages();
            Word.Images.Example_ReadWordWithImagesAndDiffWraps();
            // Word/Lists
            Word.Lists.Example_BasicLists(folderPath, false);
            Word.Lists.Example_BasicLists10(folderPath, false);
            Word.Lists.Example_BasicLists11(folderPath, false);
            Word.Lists.Example_BasicLists12(folderPath, false);
            Word.Lists.Example_BasicLists2(folderPath, false);
            Word.Lists.Example_BasicLists2Load(folderPath, false);
            Word.Lists.Example_BasicLists3(folderPath, false);
            Word.Lists.Example_BasicLists4(folderPath, false);
            Word.Lists.Example_BasicLists6(folderPath, false);
            Word.Lists.Example_BasicLists7(folderPath, false);
            Word.Lists.Example_BasicLists8(folderPath, false);
            Word.Lists.Example_BasicLists9(folderPath, false);
            Word.Lists.Example_BasicListsWithChangedStyling(folderPath, false);
            Word.Lists.Example_CloneList(folderPath, false);
            Word.Lists.Example_CustomBulletColor(folderPath, false);
            Word.Lists.Example_CustomList1(folderPath, false);
            Word.Lists.Example_DetectListStyles(folderPath, false);
            Word.Lists.Example_ListStartNumber(folderPath, false);
            Word.Lists.Example_PictureBulletList(folderPath, false);
            Word.Lists.Example_PictureBulletListAdvanced(folderPath, false);
            // Word/LoadDocuments
            Word.LoadDocuments.LoadWordDocument_Sample1(false);
            Word.LoadDocuments.LoadWordDocument_Sample2(false);
            Word.LoadDocuments.LoadWordDocument_Sample3(false);
            // Word/Macros
            Word.Macros.Example_AddMacroToExistingDocx(templatesPath, folderPath, false);
            Word.Macros.Example_CreateDocmWithMacro(templatesPath, folderPath, false);
            Word.Macros.Example_ExtractAndRemoveMacro(templatesPath, folderPath, false);
            Word.Macros.Example_ListAndRemoveMacro(templatesPath, folderPath, false);
            Word.Macros.Example_ListMacros(templatesPath, folderPath, false);
            // Word/MailMerge
            Word.MailMerge.Example_MailMergeAdvanced(folderPath, false);
            Word.MailMerge.Example_MailMergeSimple(folderPath, false);
            // Word/PageBreaks
            Word.PageBreaks.Example_PageBreaks(folderPath, false);
            Word.PageBreaks.Example_PageBreaks1(folderPath, false);
            // Word/PageNumbers
            Word.PageNumbers.Example_PageNumbers1(folderPath, false);
            Word.PageNumbers.Example_PageNumbers2(folderPath, false);
            Word.PageNumbers.Example_PageNumbers3(folderPath, false);
            Word.PageNumbers.Example_PageNumbers4(folderPath, false);
            Word.PageNumbers.Example_PageNumbers5(folderPath, false);
            Word.PageNumbers.Example_PageNumbers6(folderPath, false);
            Word.PageNumbers.Example_PageNumbers7(folderPath, false);
            Word.PageNumbers.Example_PageNumbers8(folderPath, false);
            // Word/PageSettings
            Word.PageSettings.Example_BasicSettings(folderPath, false);
            Word.PageSettings.Example_PageOrientation(folderPath, false);
            // Word/Paragraphs
            Word.Paragraphs.Example_BasicParagraphs(folderPath, false);
            Word.Paragraphs.Example_AddFormattedText(folderPath, false);
            Word.Paragraphs.Example_BasicTabStops(folderPath, false);
            Word.Paragraphs.Example_InsertParagraphAt(folderPath, false);
            Word.Paragraphs.Example_BasicParagraphStyles(folderPath, false);
            Word.Paragraphs.Example_InlineRunHelper(folderPath, false);
            Word.Paragraphs.Example_MultipleParagraphsViaDifferentWays(folderPath, false);
            Word.Paragraphs.Example_RunCharacterStylesSimple(folderPath, false);
            Word.Paragraphs.Example_RunCharacterStylesAdvanced(folderPath, false);
            Word.Paragraphs.Example_RegisterCustomParagraphStyle(folderPath, false);
            Word.Paragraphs.Example_MultipleCustomParagraphStyles(folderPath, false);
            Word.Paragraphs.Example_OverrideBuiltInParagraphStyle(folderPath, false);
            Word.Paragraphs.Example_Word_Fluent_Paragraph_TextAndFormatting(folderPath, false);
            // Word/Pdf
            Word.Pdf.Example_HeaderFooterImages(folderPath, false);
            Word.Pdf.Example_PdfInterface(folderPath, false);
            Word.Pdf.Example_SaveAsPdf(folderPath, false);
            Word.Pdf.Example_SaveAsPdfInMemory(folderPath, false);
            Word.Pdf.Example_SaveAsPdfRelative(folderPath, false);
            Word.Pdf.Example_SaveAsPdfWithHyperlinks(folderPath, false);
            Word.Pdf.Example_SaveAsPdfWithMetadata(folderPath, false);
            Word.Pdf.Example_SaveAsPdfWithLicense(folderPath, false);
            Word.Pdf.Example_SaveLists(folderPath, false);
            Word.Pdf.Example_TableStyles(folderPath, false);
            Word.Pdf.Example_PdfCustomFonts(folderPath, false);
            // Word/PictureControls
            Word.PictureControls.Example_BasicPictureControl(folderPath, false);
            // Word/Protection
            Word.Protect.Example_FinalDocument(folderPath, false);
            Word.Protect.Example_ReadOnlyEnforced(folderPath, false);
            Word.Protect.Example_ReadOnlyRecommended(folderPath, false);
            // Word/RepeatingSections
            Word.RepeatingSections.Example_BasicRepeatingSection(folderPath, false);
            // Word/Revisions
            Word.Revisions.Example_ConvertRevisionsToMarkup(folderPath, false);
            Word.Revisions.Example_TrackChangesToggle(folderPath, false);
            Word.Revisions.Example_TrackedChanges(folderPath, false);
            // Word/SaveToStream
            Word.SaveToStream.Example_CreateInProvidedStream(folderPath, false);
            Word.SaveToStream.Example_CreateInProvidedStreamAdvanced(folderPath, false);
            Word.SaveToStream.Example_SaveAsByteArray(folderPath, false);
            Word.SaveToStream.Example_SaveAsMemoryStream(folderPath, false);
            Word.SaveToStream.Example_SaveAsStream(folderPath, false);
            Word.SaveToStream.Example_SaveToOriginalStream(folderPath, false);
            Word.SaveToStream.Example_StreamDocumentProperties(folderPath, false);
            // Word/Sections
            Word.Sections.Example_BasicSections(folderPath, false);
            Word.Sections.Example_BasicSections2(folderPath, false);
            Word.Sections.Example_BasicSections3WithColumns(folderPath, false);
            Word.Sections.Example_BasicWordWithSections(folderPath, false);
            Word.Sections.Example_SectionsWithHeaders(folderPath, false);
            Word.Sections.Example_SectionsWithHeadersDefault(folderPath, false);
            Word.Sections.Example_SectionsWithParagraphs(folderPath, false);
            Word.Sections.Example_CloneSection(folderPath, false);
            // Word/Shapes
            Word.Shapes.Example_AddBasicShape(folderPath, false);
            Word.Shapes.Example_AddEllipseAndPolygon(folderPath, false);
            Word.Shapes.Example_AddLine(folderPath, false);
            Word.Shapes.Example_AddMultipleShapes(folderPath, false);
            Word.Shapes.Example_LoadShapes(folderPath, false);
            Word.Shapes.Example_RemoveShape(folderPath, false);
            // Word/SmartArt
            Word.SmartArt.Example_AddAdvancedSmartArt(folderPath, false);
            Word.SmartArt.Example_AddBasicSmartArt(folderPath, false);
            // Additional SmartArt examples from FixSmartArtShapes branch
            Word.SmartArt.Example_AddCustomSmartArt1(folderPath, false);
            Word.SmartArt.Example_AddCustomSmartArt2(folderPath, false);
            // SmartArt edit flows
            Word.SmartArt.Example_EditCustomSmartArt1(folderPath, false);
            Word.SmartArt.Example_EditCustomSmartArt2(folderPath, false);
            Word.SmartArt.Example_FlexibleBasicSmartArt_FullFlow(folderPath, false);
            Word.SmartArt.Example_FlexibleCycleSmartArt_FullFlow(folderPath, false);
            // Word/Tables
            Word.Tables.Example_AllTables(folderPath, false);
            Word.Tables.Example_BasicTables1(folderPath, false);
            Word.Tables.Example_BasicTables10_StylesModificationWithCentimeters(folderPath, false);
            Word.Tables.Example_BasicTables6(folderPath, false);
            Word.Tables.Example_BasicTables8(folderPath, false);
            Word.Tables.Example_BasicTables8_StylesModification(folderPath, false);
            Word.Tables.Example_BasicTablesLoad1(folderPath, false);
            Word.Tables.Example_BasicTablesLoad2(templatesPath, folderPath, false);
            Word.Tables.Example_BasicTablesLoad3(templatesPath, false);
            Word.Tables.Example_CloneTable(folderPath, false);
            Word.Tables.Example_ConditionalFormattingAdvanced(folderPath, false);
            Word.Tables.Example_ConditionalFormattingValues(folderPath, false);
            Word.Tables.Example_DifferentTableSizes(folderPath, false);
            Word.Tables.Example_InsertTableAfterAdvanced(folderPath, false);
            Word.Tables.Example_InsertTableAfterSimple(folderPath, false);
            Word.Tables.Example_InsertTableAfterWithXml(folderPath, false);
            Word.Tables.Example_NestedTables(folderPath, false);
            Word.Tables.Example_SplitHorizontally(folderPath, false);
            Word.Tables.Example_SplitVertically(folderPath, false);
            Word.Tables.Example_TableBorders(folderPath, false);
            Word.Tables.Example_TableCellOptions(folderPath, false);
            Word.Tables.Example_Tables(folderPath, false);
            Word.Tables.Example_Tables1CopyRow(folderPath, false);
            Word.Tables.Example_TablesAddedAfterParagraph(folderPath, false);
            Word.Tables.Example_TablesWidthAndAlignment(folderPath, false);
            Word.Tables.Example_UnifiedTableBorders(folderPath, false);
            // Word/TOC
            Word.TOC.Example_BasicTOC1(folderPath, false);
            Word.TOC.Example_BasicTOC2(folderPath, false);
            Word.TOC.Example_RemoveRegenerateTOC(folderPath, false);
            // Word/UpdateFieldsSample
            Word.UpdateFieldsSample.Example_UpdateFields(folderPath, false);
            // Word/Watermark
            Word.Watermark.Watermark_Remove(folderPath, false);
            Word.Watermark.Watermark_Sample1(folderPath, false);
            Word.Watermark.Watermark_Sample2(folderPath, false);
            Word.Watermark.Watermark_Sample3(folderPath, false);
            Word.Watermark.Watermark_SampleImage1(folderPath, false);
            // Word/WordTextBox
            Word.WordTextBox.Example_AddingTextbox(folderPath, false);
            Word.WordTextBox.Example_AddingTextbox2(folderPath, false);
            Word.WordTextBox.Example_AddingTextbox3(folderPath, false);
            Word.WordTextBox.Example_AddingTextbox4(folderPath, false);
            Word.WordTextBox.Example_AddingTextbox5(folderPath, false);
            Word.WordTextBox.Example_AddingTextboxCentimeters(folderPath, false);
            Word.WordTextBox.Example_TextBoxAutoFitOptions(folderPath, false);
            // Word/XmlSerialization
            Word.XmlSerialization.Example_XmlSerializationAdvanced(folderPath, false);
            Word.XmlSerialization.Example_XmlSerializationBasic(folderPath, false);
        }
    }
}
