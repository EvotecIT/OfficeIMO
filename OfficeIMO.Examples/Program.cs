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
            string templatesPath = Path.Combine(Directory.GetCurrentDirectory(), "Templates");
            string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "Documents");
            Setup(folderPath);

            // Excel/BasicExcelFunctionality
            OfficeIMO.Examples.Excel.BasicExcelFunctionality.BasicExcel_Example1(folderPath, false);
            OfficeIMO.Examples.Excel.BasicExcelFunctionality.BasicExcel_Example2(folderPath, false);
            OfficeIMO.Examples.Excel.BasicExcelFunctionality.BasicExcel_Example3(false);
            // Excel/BasicExcelFunctionalityAsync
            OfficeIMO.Examples.Excel.BasicExcelFunctionalityAsync.Example_ExcelAsync(folderPath).GetAwaiter().GetResult();
            // Html/Html
            OfficeIMO.Examples.Html.Html.Example_HtmlHeadings(folderPath, false);
            OfficeIMO.Examples.Html.Html.Example_HtmlImages(folderPath, false);
            OfficeIMO.Examples.Html.Html.Example_HtmlFigureWithCaption(folderPath, false);
            OfficeIMO.Examples.Html.Html.Example_HtmlInterface(folderPath, false);
            OfficeIMO.Examples.Html.Html.Example_HtmlLists(folderPath, false);
            OfficeIMO.Examples.Html.Html.Example_HtmlRoundTrip(folderPath, false);
            OfficeIMO.Examples.Html.Html.Example_HtmlTables(folderPath, false);
            // Markdown/Markdown
            OfficeIMO.Examples.Markdown.Markdown.Example_MarkdownInterface(folderPath, false);
            OfficeIMO.Examples.Markdown.Markdown.Example_MarkdownLists(folderPath, false);
            OfficeIMO.Examples.Markdown.Markdown.Example_MarkdownRoundTrip(folderPath, false);
            OfficeIMO.Examples.Markdown.Markdown.Example_MarkdownFootNotes(folderPath, false);
            OfficeIMO.Examples.Markdown.Markdown.Example_MarkdownHeadingsBoldLinks(folderPath, false);
            // Word/AdvancedDocument
            OfficeIMO.Examples.Word.AdvancedDocument.Example_AdvancedWord(folderPath, false);
            OfficeIMO.Examples.Word.AdvancedDocument.Example_AdvancedWord2(folderPath, false);
            // Word/Background
            OfficeIMO.Examples.Word.Background.Example_BackgroundImageAdvanced(folderPath, false);
            OfficeIMO.Examples.Word.Background.Example_BackgroundImageSimple(folderPath, false);
            // Word/BasicDocument
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicDocument(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicDocumentSaveAs1(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicDocumentSaveAs2(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicDocumentSaveAs3(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicDocumentWithoutUsing(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicEmptyWord(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicLoadHamlet(templatesPath, folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicWord(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicWord2(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicWordAsync(folderPath).GetAwaiter().GetResult();
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicWordWithBreaks(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicWordWithDefaultFontChange(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicWordWithDefaultStyleChange(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicWordWithLineSpacing(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicWordWithMargins(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicWordWithMarginsAndImage(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicWordWithMarginsInCentimeters(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicWordWithNewLines(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicWordWithPolishChars(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicWordWithSomeParagraphs(folderPath, false);
            OfficeIMO.Examples.Word.BasicDocument.Example_BasicWordWithTabs(folderPath, false);
            // Word/Bookmarks
            OfficeIMO.Examples.Word.Bookmarks.Example_BasicWordWithBookmarks(folderPath, false);
            // Word/BordersAndMargins
            OfficeIMO.Examples.Word.BordersAndMargins.Example_BasicPageBorders1(folderPath, false);
            OfficeIMO.Examples.Word.BordersAndMargins.Example_BasicPageBorders2(folderPath, false);
            OfficeIMO.Examples.Word.BordersAndMargins.Example_BasicWordMarginsSizes(folderPath, false);
            // Word/Charts
            OfficeIMO.Examples.Word.Charts.Example_AddingMultipleCharts(folderPath, false);
            OfficeIMO.Examples.Word.Charts.Example_Area3DChart(folderPath, false);
            OfficeIMO.Examples.Word.Charts.Example_AreaChart(folderPath, false);
            OfficeIMO.Examples.Word.Charts.Example_Bar3DChart(folderPath, false);
            OfficeIMO.Examples.Word.Charts.Example_BarChart(folderPath, false);
            OfficeIMO.Examples.Word.Charts.Example_ComboChart(folderPath, false);
            OfficeIMO.Examples.Word.Charts.Example_Line3DChart(folderPath, false);
            OfficeIMO.Examples.Word.Charts.Example_LineChart(folderPath, false);
            OfficeIMO.Examples.Word.Charts.Example_Pie3DChart(folderPath, false);
            OfficeIMO.Examples.Word.Charts.Example_PieChart(folderPath, false);
            OfficeIMO.Examples.Word.Charts.Example_RadarChart(folderPath, false);
            OfficeIMO.Examples.Word.Charts.Example_ScatterChart(folderPath, false);
            // Word/CheckBoxes
            OfficeIMO.Examples.Word.CheckBoxes.Example_BasicCheckBox(folderPath, false);
            // Word/CitationsExamples
            OfficeIMO.Examples.Word.CitationsExamples.Example_AdvancedCitations(folderPath, false);
            OfficeIMO.Examples.Word.CitationsExamples.Example_BasicCitations(folderPath, false);
            // Word/ComboBoxes
            OfficeIMO.Examples.Word.ComboBoxes.Example_BasicComboBox(folderPath, false);
            // Word/Comments
            OfficeIMO.Examples.Word.Comments.Example_PlayingWithComments(folderPath, false);
            OfficeIMO.Examples.Word.Comments.Example_RemoveCommentsAndTrack(folderPath, false);
            OfficeIMO.Examples.Word.Comments.Example_ThreadedComments(folderPath, false);
            // Word/CompareDocuments
            OfficeIMO.Examples.Word.CompareDocuments.Example_BasicComparison(folderPath, false);
            // Word/ContentControls
            OfficeIMO.Examples.Word.ContentControls.Example_AddContentControl(folderPath, false);
            OfficeIMO.Examples.Word.ContentControls.Example_AdvancedContentControls(folderPath, false);
            OfficeIMO.Examples.Word.ContentControls.Example_ContentControlsInTable(folderPath, false);
            OfficeIMO.Examples.Word.ContentControls.Example_MultipleContentControls(folderPath, false);
            // Word/CoverPages
            OfficeIMO.Examples.Word.CoverPages.Example_AddingCoverPage(folderPath, false);
            OfficeIMO.Examples.Word.CoverPages.Example_AddingCoverPage2(folderPath, false);
            // Word/CrossReferences
            OfficeIMO.Examples.Word.CrossReferences.Example_BasicCrossReferences(folderPath, false);
            // Word/CustomAndBuiltinProperties
            OfficeIMO.Examples.Word.CustomAndBuiltinProperties.Example_BasicCustomProperties(folderPath, false);
            OfficeIMO.Examples.Word.CustomAndBuiltinProperties.Example_BasicDocumentProperties(folderPath, false);
            OfficeIMO.Examples.Word.CustomAndBuiltinProperties.Example_Load(false);
            OfficeIMO.Examples.Word.CustomAndBuiltinProperties.Example_LoadDocumentWithProperties(false);
            OfficeIMO.Examples.Word.CustomAndBuiltinProperties.Example_ReadWord(false);
            OfficeIMO.Examples.Word.CustomAndBuiltinProperties.Example_ValidateDocument(folderPath);
            OfficeIMO.Examples.Word.CustomAndBuiltinProperties.Example_ValidateDocument_BeforeSave();
            // Word/DatePickers
            OfficeIMO.Examples.Word.DatePickers.Example_AdvancedDatePicker(folderPath, false);
            OfficeIMO.Examples.Word.DatePickers.Example_BasicDatePicker(folderPath, false);
            // Word/DocumentVariablesExamples
            OfficeIMO.Examples.Word.DocumentVariablesExamples.Example_AdvancedDocumentVariables(folderPath, false);
            OfficeIMO.Examples.Word.DocumentVariablesExamples.Example_BasicDocumentVariables(folderPath, false);
            // Word/DropDownLists
            OfficeIMO.Examples.Word.DropDownLists.Example_AdvancedDropDownList(folderPath, false);
            OfficeIMO.Examples.Word.DropDownLists.Example_BasicDropDownList(folderPath, false);
            // Word/Embed
            OfficeIMO.Examples.Word.Embed.Example_EmbedFileExcel(folderPath, templatesPath, false);
            OfficeIMO.Examples.Word.Embed.Example_EmbedFileHTML(folderPath, templatesPath, false);
            OfficeIMO.Examples.Word.Embed.Example_EmbedFileMultiple(folderPath, templatesPath, false);
            OfficeIMO.Examples.Word.Embed.Example_EmbedFileRTF(folderPath, templatesPath, false);
            OfficeIMO.Examples.Word.Embed.Example_EmbedFileRTFandHTML(folderPath, templatesPath, false);
            OfficeIMO.Examples.Word.Embed.Example_EmbedFileRTFandHTMLandTOC(folderPath, templatesPath, false);
            OfficeIMO.Examples.Word.Embed.Example_EmbedFragmentAfter(folderPath, false);
            OfficeIMO.Examples.Word.Embed.Example_EmbedHTMLFragment(folderPath, false);
            // Word/Equations
            OfficeIMO.Examples.Word.Equations.Example_AddEquation(folderPath, false);
            OfficeIMO.Examples.Word.Equations.Example_AddEquationExponent(folderPath, false);
            OfficeIMO.Examples.Word.Equations.Example_AddEquationIntegral(folderPath, false);
            // Word/Fields
            OfficeIMO.Examples.Word.Fields.Example_CustomFormattedDateField(folderPath, false);
            OfficeIMO.Examples.Word.Fields.Example_CustomFormattedHeaderDate(folderPath, false);
            OfficeIMO.Examples.Word.Fields.Example_CustomFormattedTimeField(folderPath, false);
            OfficeIMO.Examples.Word.Fields.Example_DocumentWithFields(folderPath, false);
            OfficeIMO.Examples.Word.Fields.Example_DocumentWithFields02(folderPath, false);
            OfficeIMO.Examples.Word.Fields.Example_FieldBuilderNested(folderPath, false);
            OfficeIMO.Examples.Word.Fields.Example_FieldBuilderSimple(folderPath, false);
            OfficeIMO.Examples.Word.Fields.Example_FieldFormatAdvanced(folderPath, false);
            OfficeIMO.Examples.Word.Fields.Example_FieldFormatRoman(folderPath, false);
            OfficeIMO.Examples.Word.Fields.Example_FieldWithMultipleSwitches(folderPath, false);
            // Word/FindAndReplace
            OfficeIMO.Examples.Word.FindAndReplace.Example_FindAndReplace01(folderPath, false);
            OfficeIMO.Examples.Word.FindAndReplace.Example_FindAndReplace02(folderPath, false);
            OfficeIMO.Examples.Word.FindAndReplace.Example_ReplaceTextWithHtmlFragment(folderPath, false);
            // Word/Fonts
            OfficeIMO.Examples.Word.Fonts.Example_EmbeddedAndBuiltinFonts(templatesPath, folderPath, false);
            OfficeIMO.Examples.Word.Fonts.Example_EmbeddedFontStyle(templatesPath, folderPath, false);
            OfficeIMO.Examples.Word.Fonts.Example_EmbedFont(templatesPath, folderPath, false);
            OfficeIMO.Examples.Word.Fonts.Example_EmbedFontWithStyle(templatesPath, folderPath, false);
            // Word/FootNotes
            OfficeIMO.Examples.Word.FootNotes.Example_DocumentWithFootNotes(folderPath, false);
            OfficeIMO.Examples.Word.FootNotes.Example_DocumentWithFootNotesEmpty(folderPath, false);
            // Word/HeadersAndFooters
            OfficeIMO.Examples.Word.HeadersAndFooters.Example_BasicWordWithHeaderAndFooter(folderPath, false);
            OfficeIMO.Examples.Word.HeadersAndFooters.Example_BasicWordWithHeaderAndFooter0(folderPath, false);
            OfficeIMO.Examples.Word.HeadersAndFooters.Example_BasicWordWithHeaderAndFooter1(folderPath, false);
            OfficeIMO.Examples.Word.HeadersAndFooters.Example_BasicWordWithHeaderAndFooterWithoutSections(folderPath, false);
            OfficeIMO.Examples.Word.HeadersAndFooters.Sections1(folderPath, false);
            // Word/HyperLinks
            OfficeIMO.Examples.Word.HyperLinks.EasyExample(folderPath, false);
            OfficeIMO.Examples.Word.HyperLinks.Example_AddingFields(folderPath, false);
            OfficeIMO.Examples.Word.HyperLinks.Example_BasicWordWithHyperLinks(folderPath, false);
            OfficeIMO.Examples.Word.HyperLinks.Example_BasicWordWithHyperLinksInTables(folderPath, false);
            OfficeIMO.Examples.Word.HyperLinks.Example_FormattedHyperLinks(folderPath, false);
            OfficeIMO.Examples.Word.HyperLinks.Example_FormattedHyperLinksAdvanced(folderPath, false);
            OfficeIMO.Examples.Word.HyperLinks.Example_FormattedHyperLinksListReuse(folderPath, false);
            // Word/Images
            OfficeIMO.Examples.Word.Images.Example_AddingFixedImages(folderPath, false);
            OfficeIMO.Examples.Word.Images.Example_AddingImages(folderPath, false);
            OfficeIMO.Examples.Word.Images.Example_AddingImagesHeadersFooters(folderPath, false);
            OfficeIMO.Examples.Word.Images.Example_AddingImagesInline(folderPath, false);
            OfficeIMO.Examples.Word.Images.Example_AddingImagesMultipleTypes(folderPath, false);
            OfficeIMO.Examples.Word.Images.Example_AddingImagesSample4(folderPath, false);
            OfficeIMO.Examples.Word.Images.Example_AddingImagesSampleToTable(folderPath, false);
            OfficeIMO.Examples.Word.Images.Example_ImageCroppingAdvanced(folderPath, false);
            OfficeIMO.Examples.Word.Images.Example_ImageCroppingBasic(folderPath, false);
            OfficeIMO.Examples.Word.Images.Example_ImageNewFeatures(folderPath, false);
            OfficeIMO.Examples.Word.Images.Example_ImageTransparencyAdvanced(folderPath, false);
            OfficeIMO.Examples.Word.Images.Example_ImageTransparencySimple(folderPath, false);
            OfficeIMO.Examples.Word.Images.Example_ReadWordWithImages();
            OfficeIMO.Examples.Word.Images.Example_ReadWordWithImagesAndDiffWraps();
            // Word/Lists
            OfficeIMO.Examples.Word.Lists.Example_BasicLists(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_BasicLists10(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_BasicLists11(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_BasicLists12(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_BasicLists2(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_BasicLists2Load(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_BasicLists3(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_BasicLists4(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_BasicLists6(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_BasicLists7(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_BasicLists8(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_BasicLists9(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_BasicListsWithChangedStyling(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_CloneList(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_CustomBulletColor(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_CustomList1(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_DetectListStyles(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_ListStartNumber(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_PictureBulletList(folderPath, false);
            OfficeIMO.Examples.Word.Lists.Example_PictureBulletListAdvanced(folderPath, false);
            // Word/LoadDocuments
            OfficeIMO.Examples.Word.LoadDocuments.LoadWordDocument_Sample1(false);
            OfficeIMO.Examples.Word.LoadDocuments.LoadWordDocument_Sample2(false);
            OfficeIMO.Examples.Word.LoadDocuments.LoadWordDocument_Sample3(false);
            // Word/Macros
            OfficeIMO.Examples.Word.Macros.Example_AddMacroToExistingDocx(templatesPath, folderPath, false);
            OfficeIMO.Examples.Word.Macros.Example_CreateDocmWithMacro(templatesPath, folderPath, false);
            OfficeIMO.Examples.Word.Macros.Example_ExtractAndRemoveMacro(templatesPath, folderPath, false);
            OfficeIMO.Examples.Word.Macros.Example_ListAndRemoveMacro(templatesPath, folderPath, false);
            OfficeIMO.Examples.Word.Macros.Example_ListMacros(templatesPath, folderPath, false);
            // Word/MailMerge
            OfficeIMO.Examples.Word.MailMerge.Example_MailMergeAdvanced(folderPath, false);
            OfficeIMO.Examples.Word.MailMerge.Example_MailMergeSimple(folderPath, false);
            // Word/PageBreaks
            OfficeIMO.Examples.Word.PageBreaks.Example_PageBreaks(folderPath, false);
            OfficeIMO.Examples.Word.PageBreaks.Example_PageBreaks1(folderPath, false);
            // Word/PageNumbers
            OfficeIMO.Examples.Word.PageNumbers.Example_PageNumbers1(folderPath, false);
            OfficeIMO.Examples.Word.PageNumbers.Example_PageNumbers2(folderPath, false);
            OfficeIMO.Examples.Word.PageNumbers.Example_PageNumbers3(folderPath, false);
            OfficeIMO.Examples.Word.PageNumbers.Example_PageNumbers4(folderPath, false);
            OfficeIMO.Examples.Word.PageNumbers.Example_PageNumbers5(folderPath, false);
            OfficeIMO.Examples.Word.PageNumbers.Example_PageNumbers6(folderPath, false);
            OfficeIMO.Examples.Word.PageNumbers.Example_PageNumbers7(folderPath, false);
            OfficeIMO.Examples.Word.PageNumbers.Example_PageNumbers8(folderPath, false);
            // Word/Pdf
            OfficeIMO.Examples.Word.Pdf.Example_HeaderFooterImages(folderPath, false);
            OfficeIMO.Examples.Word.Pdf.Example_PdfInterface(folderPath, false);
            OfficeIMO.Examples.Word.Pdf.Example_SaveAsPdf(folderPath, false);
            OfficeIMO.Examples.Word.Pdf.Example_SaveAsPdfInMemory(folderPath, false);
            OfficeIMO.Examples.Word.Pdf.Example_SaveAsPdfRelative(folderPath, false);
            OfficeIMO.Examples.Word.Pdf.Example_SaveAsPdfWithHyperlinks(folderPath, false);
            OfficeIMO.Examples.Word.Pdf.Example_SaveAsPdfWithMetadata(folderPath, false);
            OfficeIMO.Examples.Word.Pdf.Example_SaveAsPdfWithLicense(folderPath, false);
            OfficeIMO.Examples.Word.Pdf.Example_SaveLists(folderPath, false);
            OfficeIMO.Examples.Word.Pdf.Example_TableStyles(folderPath, false);
            // Word/PictureControls
            OfficeIMO.Examples.Word.PictureControls.Example_BasicPictureControl(folderPath, false);
            // Word/RepeatingSections
            OfficeIMO.Examples.Word.RepeatingSections.Example_BasicRepeatingSection(folderPath, false);
            // Word/Revisions
            OfficeIMO.Examples.Word.Revisions.Example_ConvertRevisionsToMarkup(folderPath, false);
            OfficeIMO.Examples.Word.Revisions.Example_TrackChangesToggle(folderPath, false);
            OfficeIMO.Examples.Word.Revisions.Example_TrackedChanges(folderPath, false);
            // Word/SaveToStream
            OfficeIMO.Examples.Word.SaveToStream.Example_CreateInProvidedStream(folderPath, false);
            OfficeIMO.Examples.Word.SaveToStream.Example_CreateInProvidedStreamAdvanced(folderPath, false);
            OfficeIMO.Examples.Word.SaveToStream.Example_SaveAsByteArray(folderPath, false);
            OfficeIMO.Examples.Word.SaveToStream.Example_SaveAsMemoryStream(folderPath, false);
            OfficeIMO.Examples.Word.SaveToStream.Example_SaveAsStream(folderPath, false);
            OfficeIMO.Examples.Word.SaveToStream.Example_SaveToOriginalStream(folderPath, false);
            OfficeIMO.Examples.Word.SaveToStream.Example_StreamDocumentProperties(folderPath, false);
            // Word/Sections
            OfficeIMO.Examples.Word.Sections.Example_BasicSections(folderPath, false);
            OfficeIMO.Examples.Word.Sections.Example_BasicSections2(folderPath, false);
            OfficeIMO.Examples.Word.Sections.Example_BasicSections3WithColumns(folderPath, false);
            OfficeIMO.Examples.Word.Sections.Example_BasicWordWithSections(folderPath, false);
            OfficeIMO.Examples.Word.Sections.Example_SectionsWithHeaders(folderPath, false);
            OfficeIMO.Examples.Word.Sections.Example_SectionsWithHeadersDefault(folderPath, false);
            OfficeIMO.Examples.Word.Sections.Example_SectionsWithParagraphs(folderPath, false);
            // Word/Shapes
            OfficeIMO.Examples.Word.Shapes.Example_AddBasicShape(folderPath, false);
            OfficeIMO.Examples.Word.Shapes.Example_AddEllipseAndPolygon(folderPath, false);
            OfficeIMO.Examples.Word.Shapes.Example_AddLine(folderPath, false);
            OfficeIMO.Examples.Word.Shapes.Example_AddMultipleShapes(folderPath, false);
            OfficeIMO.Examples.Word.Shapes.Example_LoadShapes(folderPath, false);
            OfficeIMO.Examples.Word.Shapes.Example_RemoveShape(folderPath, false);
            // Word/SmartArt
            OfficeIMO.Examples.Word.SmartArt.Example_AddAdvancedSmartArt(folderPath, false);
            OfficeIMO.Examples.Word.SmartArt.Example_AddBasicSmartArt(folderPath, false);
            // Word/Tables
            OfficeIMO.Examples.Word.Tables.Example_AllTables(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_BasicTables1(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_BasicTables10_StylesModificationWithCentimeters(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_BasicTables6(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_BasicTables8(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_BasicTables8_StylesModification(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_BasicTablesLoad1(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_BasicTablesLoad2(templatesPath, folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_BasicTablesLoad3(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_CloneTable(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_ConditionalFormattingAdvanced(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_ConditionalFormattingValues(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_DifferentTableSizes(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_InsertTableAfterAdvanced(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_InsertTableAfterSimple(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_InsertTableAfterWithXml(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_NestedTables(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_SplitHorizontally(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_SplitVertically(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_TableBorders(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_TableCellOptions(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_Tables(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_Tables1CopyRow(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_TablesAddedAfterParagraph(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_TablesWidthAndAlignment(folderPath, false);
            OfficeIMO.Examples.Word.Tables.Example_UnifiedTableBorders(folderPath, false);
            // Word/TOC
            OfficeIMO.Examples.Word.TOC.Example_BasicTOC1(folderPath, false);
            OfficeIMO.Examples.Word.TOC.Example_BasicTOC2(folderPath, false);
            OfficeIMO.Examples.Word.TOC.Example_RemoveRegenerateTOC(folderPath, false);
            // Word/UpdateFieldsSample
            OfficeIMO.Examples.Word.UpdateFieldsSample.Example_UpdateFields(folderPath, false);
            // Word/Watermark
            OfficeIMO.Examples.Word.Watermark.Watermark_Remove(folderPath, false);
            OfficeIMO.Examples.Word.Watermark.Watermark_Sample1(folderPath, false);
            OfficeIMO.Examples.Word.Watermark.Watermark_Sample2(folderPath, false);
            OfficeIMO.Examples.Word.Watermark.Watermark_Sample3(folderPath, false);
            OfficeIMO.Examples.Word.Watermark.Watermark_SampleImage1(folderPath, false);
            // Word/WordTextBox
            OfficeIMO.Examples.Word.WordTextBox.Example_AddingTextbox(folderPath, false);
            OfficeIMO.Examples.Word.WordTextBox.Example_AddingTextbox2(folderPath, false);
            OfficeIMO.Examples.Word.WordTextBox.Example_AddingTextbox3(folderPath, false);
            OfficeIMO.Examples.Word.WordTextBox.Example_AddingTextbox4(folderPath, false);
            OfficeIMO.Examples.Word.WordTextBox.Example_AddingTextbox5(folderPath, false);
            OfficeIMO.Examples.Word.WordTextBox.Example_AddingTextboxCentimeters(folderPath, false);
            OfficeIMO.Examples.Word.WordTextBox.Example_TextBoxAutoFitOptions(folderPath, false);
            // Word/XmlSerialization
            OfficeIMO.Examples.Word.XmlSerialization.Example_XmlSerializationAdvanced(folderPath, false);
            OfficeIMO.Examples.Word.XmlSerialization.Example_XmlSerializationBasic(folderPath, false);
        }
    }
}