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
            BasicDocument.Example_BasicLoadHamlet(templatesPath, folderPath, false);
            BasicDocument.Example_BasicWordWithPolishChars(folderPath, false);
            BasicDocument.Example_BasicWordWithNewLines(folderPath, false);
            BasicDocument.Example_BasicWordWithTabs(folderPath, false);
            BasicDocument.Example_BasicWordWithMargins(folderPath, false);
            BasicDocument.Example_BasicWordWithMarginsInCentimeters(folderPath, false);
            BasicDocument.Example_BasicWordWithMarginsAndImage(folderPath, false);
            BasicDocument.Example_BasicWordWithLineSpacing(folderPath, false);

            AdvancedDocument.Example_AdvancedWord(folderPath, false);
            AdvancedDocument.Example_AdvancedWord2(folderPath, false);

            Paragraphs.Example_BasicParagraphs(folderPath, false);
            Paragraphs.Example_BasicParagraphStyles(folderPath, false);
            Paragraphs.Example_MultipleParagraphsViaDifferentWays(folderPath, false);
            Paragraphs.Example_BasicTabStops(folderPath, false);

            BasicDocument.Example_BasicDocument(folderPath, false);
            BasicDocument.Example_BasicDocumentSaveAs1(folderPath, false);
            BasicDocument.Example_BasicDocumentSaveAs2(folderPath, false);
            BasicDocument.Example_BasicDocumentSaveAs3(folderPath, false);
            BasicDocument.Example_BasicDocumentWithoutUsing(folderPath, false);

            Lists.Example_BasicLists(folderPath, false);
            Lists.Example_BasicLists6(folderPath, false);
            Lists.Example_BasicLists2(folderPath, false);
            Lists.Example_BasicLists3(folderPath, false);
            Lists.Example_BasicLists4(folderPath, false);
            Lists.Example_BasicLists2Load(folderPath, false);
            Lists.Example_BasicLists7(folderPath, false);
            Lists.Example_BasicLists8(folderPath, false);
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

            PageSettings.Example_BasicSettings(folderPath, false);
            PageSettings.Example_PageOrientation(folderPath, false);

            PageNumbers.Example_PageNumbers1(folderPath, false);

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
            HyperLinks.Example_AddingFields(folderPath, false);
            HyperLinks.Example_BasicWordWithHyperLinksInTables(folderPath, false);

            HeadersAndFooters.Sections1(folderPath, false);
            HeadersAndFooters.Example_BasicWordWithHeaderAndFooter0(folderPath, false);
            HeadersAndFooters.Example_BasicWordWithHeaderAndFooter(folderPath, false);
            HeadersAndFooters.Example_BasicWordWithHeaderAndFooter1(folderPath, false);

            Charts.Example_AddingMultipleCharts(folderPath, false);

            Images.Example_AddingImages(folderPath, false);
            Images.Example_ReadWordWithImages();
            Images.Example_AddingImagesMultipleTypes(folderPath, false);
            Images.Example_ReadWordWithImagesAndDiffWraps();
            Images.Example_AddingFixedImages(folderPath, false);
            Images.Example_AddingImagesSampleToTable(folderPath, false);

            PageBreaks.Example_PageBreaks(folderPath, false);
            PageBreaks.Example_PageBreaks1(folderPath, false);

            HeadersAndFooters.Example_BasicWordWithHeaderAndFooterWithoutSections(folderPath, false);

            TOC.Example_BasicTOC1(folderPath, false);
            TOC.Example_BasicTOC2(folderPath, false);

            Comments.Example_PlayingWithComments(folderPath, false);

            BasicExcelFunctionality.BasicExcel_Example1(folderPath, false);
            BasicExcelFunctionality.BasicExcel_Example2(folderPath, false);
            BasicExcelFunctionality.BasicExcel_Example3(false);

            BordersAndMargins.Example_BasicWordMarginsSizes(folderPath, false);
            BordersAndMargins.Example_BasicPageBorders1(folderPath, false);
            BordersAndMargins.Example_BasicPageBorders2(folderPath, false);

            Bookmarks.Example_BasicWordWithBookmarks(folderPath, false);
            Fields.Example_DocumentWithFields(folderPath, false);
            Fields.Example_DocumentWithFields02(folderPath, false);

            Watermark.Watermark_Sample2(folderPath, false);
            Watermark.Watermark_Sample1(folderPath, false);
            Watermark.Watermark_Sample3(folderPath, false);

            Embed.Example_EmbedFileHTML(folderPath, templatesPath, false);
            Embed.Example_EmbedFileRTF(folderPath, templatesPath, false);
            Embed.Example_EmbedFileRTFandHTML(folderPath, templatesPath, false);
            Embed.Example_EmbedFileRTFandHTMLandTOC(folderPath, templatesPath, false);
            Embed.Example_EmbedFileMultiple(folderPath, templatesPath, false);

            CleanupDocuments.CleanupDocuments_Sample01(false);
            CleanupDocuments.CleanupDocuments_Sample02(folderPath, false);

            FindAndReplace.Example_FindAndReplace01(folderPath, false);
            FindAndReplace.Example_FindAndReplace02(folderPath, false);

            FootNotes.Example_DocumentWithFootNotes(templatesPath, false);
            FootNotes.Example_DocumentWithFootNotesEmpty(folderPath, false);

            SaveToStream.Example_StreamDocumentProperties(folderPath, false);

            Protect.Example_ProtectFinalDocument(folderPath, false);
            Protect.Example_ProtectAlwaysReadOnly(folderPath, false);

            WordTextBox.Example_AddingTextbox(folderPath, false);
            WordTextBox.Example_AddingTextbox2(folderPath, false);
            WordTextBox.Example_AddingTextbox4(folderPath, false);
            WordTextBox.Example_AddingTextbox5(folderPath, false);
            WordTextBox.Example_AddingTextbox3(folderPath, false);
            WordTextBox.Example_AddingTextboxCentimeters(folderPath, false);
        }
    }
}
