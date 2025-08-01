using System.IO;

using OfficeIMO.Examples.Excel;
using OfficeIMO.Examples.Word;

namespace OfficeIMO.Examples {
    internal static class Program {
        private static void Setup(string path) {
            if (!Directory.Exists(path)) {
                Directory.CreateDirectory(path);
            }
        }

        static void Main(string[] args) {
            string templatesPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");
            string folderPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Documents");
            Setup(folderPath);

            BasicDocument.Example_BasicEmptyWord(folderPath, false);
            BasicDocument.Example_BasicWord(folderPath, false);
            BasicDocument.Example_BasicWord2(folderPath, false);
            BasicDocument.Example_BasicWordWithBreaks(folderPath, false);
            BasicDocument.Example_BasicWordWithDefaultStyleChange(folderPath, false);
            BasicDocument.Example_BasicWordWithDefaultFontChange(folderPath, false);
            Fonts.Example_EmbedFont(templatesPath, folderPath, false);
            Fonts.Example_EmbeddedAndBuiltinFonts(templatesPath, folderPath, false);
            Fonts.Example_EmbeddedFontStyle(templatesPath, folderPath, false);
            Fonts.Example_EmbedFontWithStyle(templatesPath, folderPath, false);
            BasicDocument.Example_BasicLoadHamlet(templatesPath, folderPath, false);
            BasicDocument.Example_BasicWordWithPolishChars(folderPath, false);
            BasicDocument.Example_BasicWordWithNewLines(folderPath, false);
            BasicDocument.Example_BasicWordWithTabs(folderPath, false);
            BasicDocument.Example_BasicWordWithMargins(folderPath, false);
            BasicDocument.Example_BasicWordWithMarginsInCentimeters(folderPath, false);
            BasicDocument.Example_BasicWordWithMarginsAndImage(folderPath, false);
            BasicDocument.Example_BasicWordWithLineSpacing(folderPath, false);
            BasicDocument.Example_BasicWordWithSomeParagraphs(folderPath, false);
            BasicDocument.Example_BasicWordAsync(folderPath).GetAwaiter().GetResult();

            AdvancedDocument.Example_AdvancedWord(folderPath, false);
            AdvancedDocument.Example_AdvancedWord2(folderPath, false);

            ContentControls.Example_AddContentControl(folderPath, false);
            ContentControls.Example_MultipleContentControls(folderPath, false);
            ContentControls.Example_AdvancedContentControls(folderPath, false);
            ContentControls.Example_ContentControlsInTable(folderPath, false);
            CheckBoxes.Example_BasicCheckBox(folderPath, false);

            DatePickers.Example_BasicDatePicker(folderPath, false);
            DatePickers.Example_AdvancedDatePicker(folderPath, false);
            DropDownLists.Example_BasicDropDownList(folderPath, false);
            DropDownLists.Example_AdvancedDropDownList(folderPath, false);
            Paragraphs.Example_BasicParagraphs(folderPath, false);
            Paragraphs.Example_BasicParagraphStyles(folderPath, false);
            Paragraphs.Example_RegisterCustomParagraphStyle(folderPath, false);
            Paragraphs.Example_MultipleCustomParagraphStyles(folderPath, false);
            Paragraphs.Example_OverrideBuiltInParagraphStyle(folderPath, false);
            Paragraphs.Example_MultipleParagraphsViaDifferentWays(folderPath, false);
            Paragraphs.Example_BasicTabStops(folderPath, false);
            Paragraphs.Example_RunCharacterStylesSimple(folderPath, false);
            Paragraphs.Example_RunCharacterStylesAdvanced(folderPath, false);
            Paragraphs.Example_InsertParagraphAt(folderPath, false);

            BasicDocument.Example_BasicDocument(folderPath, false);
            BasicDocument.Example_BasicDocumentSaveAs1(folderPath, false);
            BasicDocument.Example_BasicDocumentSaveAs2(folderPath, false);
            BasicDocument.Example_BasicDocumentSaveAs3(folderPath, false);
            BasicDocument.Example_BasicDocumentWithoutUsing(folderPath, false);

            Lists.Example_BasicLists(folderPath, false);
            Lists.Example_BasicLists6(folderPath, false);
            Lists.Example_BasicLists2(folderPath, false);
            Lists.Example_BasicLists3(folderPath, false);
            Lists.Example_BasicLists9(folderPath, false);
            Lists.Example_BasicLists4(folderPath, false);
            Lists.Example_BasicLists2Load(folderPath, false);
            Lists.Example_BasicLists7(folderPath, false);
            Lists.Example_BasicLists8(folderPath, false);
            Lists.Example_BasicLists10(folderPath, false);
            Lists.Example_BasicLists11(folderPath, false);
            Lists.Example_BasicLists12(folderPath, false);
            Lists.Example_CustomList1(folderPath, false);
            Lists.Example_BasicListsWithChangedStyling(folderPath, false);
            Lists.Example_CloneList(folderPath, false);
            Lists.Example_ListStartNumber(folderPath, false);
            Lists.Example_PictureBulletList(folderPath, false);
            Lists.Example_PictureBulletListAdvanced(folderPath, false);

            Tables.Example_BasicTables1(folderPath, false);
            Tables.Example_BasicTablesLoad1(folderPath, false);
            Tables.Example_BasicTablesLoad2(templatesPath, folderPath, false);
            Tables.Example_BasicTablesLoad3(templatesPath, false);
            Tables.Example_TablesWidthAndAlignment(folderPath, false);
            Tables.Example_AllTables(folderPath, false);
            Tables.Example_Tables(folderPath, false);
            Tables.Example_TableBorders(folderPath, false);
            Tables.Example_NestedTables(folderPath, false);
            Tables.Example_TablesAddedAfterParagraph(folderPath, false);
            Tables.Example_InsertTableAfterSimple(folderPath, false);
            Tables.Example_InsertTableAfterAdvanced(folderPath, false);
            Tables.Example_BasicTables8(folderPath, false);
            Tables.Example_Tables1CopyRow(folderPath, false);
            Tables.Example_BasicTables8_StylesModification(folderPath, false);
            Tables.Example_UnifiedTableBorders(folderPath, false);
            Tables.Example_BasicTables10_StylesModificationWithCentimeters(folderPath, false);
            Tables.Example_DifferentTableSizes(folderPath, false);
            Tables.Example_CloneTable(folderPath, false);
            Tables.Example_SplitVertically(folderPath, false);
            Tables.Example_SplitHorizontally(folderPath, false);
            Tables.Example_ConditionalFormattingValues(folderPath, false);
            Tables.Example_ConditionalFormattingAdvanced(folderPath, false);
            PageSettings.Example_BasicSettings(folderPath, false);
            PageSettings.Example_PageOrientation(folderPath, false);

            PageNumbers.Example_PageNumbers1(folderPath, false);
            PageNumbers.Example_PageNumbers2(folderPath, false);
            PageNumbers.Example_PageNumbers3(folderPath, false);
            PageNumbers.Example_PageNumbers4(folderPath, false);
            PageNumbers.Example_PageNumbers5(folderPath, false);
            PageNumbers.Example_PageNumbers6(folderPath, false);
            PageNumbers.Example_PageNumbers7(folderPath, false);
            PageNumbers.Example_PageNumbers8(folderPath, false);

            Sections.Example_BasicSections(folderPath, false);
            Sections.Example_BasicSections2(folderPath, false);
            Sections.Example_BasicSections3WithColumns(folderPath, false);
            Sections.Example_SectionsWithParagraphs(folderPath, false);
            Sections.Example_SectionsWithHeadersDefault(folderPath, false);
            Sections.Example_SectionsWithHeaders(folderPath, false);
            Sections.Example_BasicWordWithSections(folderPath, false);

            CoverPages.Example_AddingCoverPage(folderPath, false);
            CoverPages.Example_AddingCoverPage2(folderPath, false);

            LoadDocuments.LoadWordDocument_Sample1(false);
            LoadDocuments.LoadWordDocument_Sample2(false);
            LoadDocuments.LoadWordDocument_Sample3(false);

            CustomAndBuiltinProperties.Example_BasicDocumentProperties(folderPath, false);
            CustomAndBuiltinProperties.Example_ReadWord(false);
            CustomAndBuiltinProperties.Example_BasicCustomProperties(folderPath, false);
            CustomAndBuiltinProperties.Example_ValidateDocument(folderPath);
            CustomAndBuiltinProperties.Example_ValidateDocument_BeforeSave();
            CustomAndBuiltinProperties.Example_LoadDocumentWithProperties(false);
            CustomAndBuiltinProperties.Example_Load(false);

            HyperLinks.EasyExample(folderPath, false);
            HyperLinks.Example_BasicWordWithHyperLinks(folderPath, false);
            HyperLinks.Example_FormattedHyperLinks(folderPath, false);
            HyperLinks.Example_FormattedHyperLinksAdvanced(folderPath, false);
            HyperLinks.Example_FormattedHyperLinksListReuse(folderPath, false);
            HyperLinks.Example_AddingFields(folderPath, false);
            HyperLinks.Example_BasicWordWithHyperLinksInTables(folderPath, false);

            HeadersAndFooters.Sections1(folderPath, false);
            HeadersAndFooters.Example_BasicWordWithHeaderAndFooter0(folderPath, false);
            HeadersAndFooters.Example_BasicWordWithHeaderAndFooter(folderPath, false);
            HeadersAndFooters.Example_BasicWordWithHeaderAndFooter1(folderPath, false);

            Charts.Example_AddingMultipleCharts(folderPath, false);
            Charts.Example_BarChart(folderPath, false);
            Charts.Example_PieChart(folderPath, false);
            Charts.Example_LineChart(folderPath, false);
            Charts.Example_AreaChart(folderPath, false);
            Charts.Example_ScatterChart(folderPath, false);
            Charts.Example_RadarChart(folderPath, false);
            Charts.Example_Bar3DChart(folderPath, false);
            Charts.Example_Pie3DChart(folderPath, false);
            Charts.Example_Line3DChart(folderPath, false);
            Charts.Example_Area3DChart(folderPath, false);

            Images.Example_AddingImages(folderPath, false);
            Images.Example_ReadWordWithImages();
            Images.Example_AddingImagesMultipleTypes(folderPath, false);
            Images.Example_ReadWordWithImagesAndDiffWraps();
            Images.Example_AddingFixedImages(folderPath, false);
            Images.Example_AddingImagesSampleToTable(folderPath, false);
            Images.Example_ImageTransparencySimple(folderPath, false);
            Images.Example_ImageTransparencyAdvanced(folderPath, false);
            Images.Example_ImageNewFeatures(folderPath, false);

            Background.Example_BackgroundImageSimple(folderPath, false);
            Background.Example_BackgroundImageAdvanced(folderPath, false);

            PageBreaks.Example_PageBreaks(folderPath, false);
            PageBreaks.Example_PageBreaks1(folderPath, false);

            HeadersAndFooters.Example_BasicWordWithHeaderAndFooterWithoutSections(folderPath, false);

            TOC.Example_BasicTOC1(folderPath, false);
            TOC.Example_BasicTOC2(folderPath, false);

            Comments.Example_PlayingWithComments(folderPath, false);
            Comments.Example_RemoveCommentsAndTrack(folderPath, false);
            Comments.Example_ThreadedComments(folderPath, false);

            BasicExcelFunctionality.BasicExcel_Example1(folderPath, false);
            BasicExcelFunctionality.BasicExcel_Example2(folderPath, false);
            BasicExcelFunctionality.BasicExcel_Example3(false);
            BasicExcelFunctionalityAsync.Example_ExcelAsync(folderPath).GetAwaiter().GetResult();

            BordersAndMargins.Example_BasicWordMarginsSizes(folderPath, false);
            BordersAndMargins.Example_BasicPageBorders1(folderPath, false);
            BordersAndMargins.Example_BasicPageBorders2(folderPath, false);

            Bookmarks.Example_BasicWordWithBookmarks(folderPath, false);
            Fields.Example_DocumentWithFields(folderPath, false);
            Fields.Example_DocumentWithFields02(folderPath, false);
            Fields.Example_CustomFormattedDateField(folderPath, false);
            Fields.Example_CustomFormattedTimeField(folderPath, false);
            Fields.Example_CustomFormattedHeaderDate(folderPath, false);
            Fields.Example_FieldFormatRoman(folderPath, false);
            Fields.Example_FieldFormatAdvanced(folderPath, false);
            Fields.Example_FieldWithMultipleSwitches(folderPath, false);

            CitationsExamples.Example_BasicCitations(folderPath, false);
            CitationsExamples.Example_AdvancedCitations(folderPath, false);
            CrossReferences.Example_BasicCrossReferences(folderPath, false);

            Watermark.Watermark_Sample2(folderPath, false);
            Watermark.Watermark_Sample1(folderPath, false);
            Watermark.Watermark_Sample3(folderPath, false);
            Watermark.Watermark_SampleImage1(folderPath, false);
            Watermark.Watermark_Remove(folderPath, false);

            Embed.Example_EmbedFileHTML(folderPath, templatesPath, false);
            Embed.Example_EmbedFileRTF(folderPath, templatesPath, false);
            Embed.Example_EmbedFileRTFandHTML(folderPath, templatesPath, false);
            Embed.Example_EmbedFileRTFandHTMLandTOC(folderPath, templatesPath, false);
            Embed.Example_EmbedFileMultiple(folderPath, templatesPath, false);
            Embed.Example_EmbedHTMLFragment(folderPath, false);
            Embed.Example_EmbedFragmentAfter(folderPath, false);

            CleanupDocuments.CleanupDocuments_Sample01(false);
            CleanupDocuments.CleanupDocuments_Sample02(folderPath, false);

            FindAndReplace.Example_FindAndReplace01(folderPath, false);
            FindAndReplace.Example_FindAndReplace02(folderPath, false);
            FindAndReplace.Example_ReplaceTextWithHtmlFragment(folderPath, false);

            FootNotes.Example_DocumentWithFootNotes(templatesPath, false);
            FootNotes.Example_DocumentWithFootNotesEmpty(folderPath, false);

            SaveToStream.Example_StreamDocumentProperties(folderPath, false);
            SaveToStream.Example_CreateInProvidedStream(folderPath, false);
            SaveToStream.Example_CreateInProvidedStreamAdvanced(folderPath, false);
            SaveToStream.Example_SaveToOriginalStream(folderPath, false);

            Protect.Example_FinalDocument(folderPath, false);
            Protect.Example_ReadOnlyEnforced(folderPath, false);
            Protect.Example_ReadOnlyRecommended(folderPath, false);

            WordTextBox.Example_AddingTextbox(folderPath, false);
            WordTextBox.Example_AddingTextbox2(folderPath, false);
            WordTextBox.Example_AddingTextbox4(folderPath, false);
            WordTextBox.Example_AddingTextbox5(folderPath, false);
            WordTextBox.Example_AddingTextbox3(folderPath, false);
            WordTextBox.Example_AddingTextboxCentimeters(folderPath, false);

            Embed.Example_EmbedFileExcel(folderPath, templatesPath, false);
            Shapes.Example_AddBasicShape(folderPath, false);
            Shapes.Example_AddLine(folderPath, false);
            Shapes.Example_AddEllipseAndPolygon(folderPath, false);
            Shapes.Example_AddMultipleShapes(folderPath, false);
            Shapes.Example_RemoveShape(folderPath, false);
            Shapes.Example_LoadShapes(folderPath, false);
            SmartArt.Example_AddBasicSmartArt(folderPath, false);
            SmartArt.Example_AddAdvancedSmartArt(folderPath, false);

            Revisions.Example_TrackedChanges(folderPath, false);
            MailMerge.Example_MailMergeSimple(folderPath, false);
            MailMerge.Example_MailMergeAdvanced(folderPath, false);

            Macros.Example_CreateDocmWithMacro(templatesPath, folderPath, false);
            Macros.Example_AddMacroToExistingDocx(templatesPath, folderPath, false);
            Macros.Example_ListMacros(templatesPath, folderPath, false);
            Macros.Example_ExtractAndRemoveMacro(templatesPath, folderPath, false);
            Macros.Example_ListAndRemoveMacro(templatesPath, folderPath, false);

            XmlSerialization.Example_XmlSerializationBasic(folderPath, false);
            XmlSerialization.Example_XmlSerializationAdvanced(folderPath, false);
        }
    }
}
