using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a cover page within a Word document.
    /// </summary>
    public partial class WordCoverPage {

        private static SdtBlock CoverPageSideLine {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId();

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Cover Pages" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();
                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "001824A2", RsidRunAdditionDefault = "001824A2", ParagraphId = "28502A23", TextId = "2F688D08" };

                Table table1 = new Table();

                TableProperties tableProperties1 = new TableProperties();
                TablePositionProperties tablePositionProperties1 = new TablePositionProperties() { LeftFromText = 187, RightFromText = 187, HorizontalAnchor = HorizontalAnchorValues.Margin, TablePositionXAlignment = HorizontalAlignmentValues.Center, TablePositionY = 2881 };
                TableWidth tableWidth1 = new TableWidth() { Width = "4000", Type = TableWidthUnitValues.Pct };

                TableBorders tableBorders1 = new TableBorders();
                LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "4472C4", ThemeColor = ThemeColorValues.Accent1, Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                tableBorders1.Append(leftBorder1);

                TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
                TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 144, Type = TableWidthValues.Dxa };
                TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 115, Type = TableWidthValues.Dxa };

                tableCellMarginDefault1.Append(tableCellLeftMargin1);
                tableCellMarginDefault1.Append(tableCellRightMargin1);
                TableLook tableLook1 = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

                tableProperties1.Append(tablePositionProperties1);
                tableProperties1.Append(tableWidth1);
                tableProperties1.Append(tableBorders1);
                tableProperties1.Append(tableCellMarginDefault1);
                tableProperties1.Append(tableLook1);

                TableGrid tableGrid1 = new TableGrid();
                GridColumn gridColumn1 = new GridColumn() { Width = "7476" };

                tableGrid1.Append(gridColumn1);

                TableRow tableRow1 = new TableRow() { RsidTableRowAddition = "001824A2", ParagraphId = "47020744", TextId = "77777777" };

                SdtCell sdtCell1 = new SdtCell();

                SdtProperties sdtProperties2 = new SdtProperties();

                RunProperties runProperties1 = new RunProperties();
                Color color1 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
                FontSize fontSize1 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

                runProperties1.Append(color1);
                runProperties1.Append(fontSize1);
                runProperties1.Append(fontSizeComplexScript1);
                SdtAlias sdtAlias1 = new SdtAlias() { Val = "Company" };
                SdtId sdtId2 = new SdtId();

                SdtPlaceholder sdtPlaceholder1 = new SdtPlaceholder();
                DocPartReference docPartReference1 = new DocPartReference() { Val = "6104919ED59E4975BF9F41C4B89124BB" };

                sdtPlaceholder1.Append(docPartReference1);
                ShowingPlaceholder showingPlaceholder1 = new ShowingPlaceholder();
                DataBinding dataBinding1 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\'", XPath = "/ns0:Properties[1]/ns0:Company[1]", StoreItemId = "{6668398D-A668-4E3E-A5EB-62B293D839F1}" };
                SdtContentText sdtContentText1 = new SdtContentText();

                sdtProperties2.Append(runProperties1);
                sdtProperties2.Append(sdtAlias1);
                sdtProperties2.Append(sdtId2);
                sdtProperties2.Append(sdtPlaceholder1);
                sdtProperties2.Append(showingPlaceholder1);
                sdtProperties2.Append(dataBinding1);
                sdtProperties2.Append(sdtContentText1);

                SdtContentCell sdtContentCell1 = new SdtContentCell();

                TableCell tableCell1 = new TableCell();

                TableCellProperties tableCellProperties1 = new TableCellProperties();
                TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "7672", Type = TableWidthUnitValues.Dxa };

                TableCellMargin tableCellMargin1 = new TableCellMargin();
                TopMargin topMargin1 = new TopMargin() { Width = "216", Type = TableWidthUnitValues.Dxa };
                LeftMargin leftMargin1 = new LeftMargin() { Width = "115", Type = TableWidthUnitValues.Dxa };
                BottomMargin bottomMargin1 = new BottomMargin() { Width = "216", Type = TableWidthUnitValues.Dxa };
                RightMargin rightMargin1 = new RightMargin() { Width = "115", Type = TableWidthUnitValues.Dxa };

                tableCellMargin1.Append(topMargin1);
                tableCellMargin1.Append(leftMargin1);
                tableCellMargin1.Append(bottomMargin1);
                tableCellMargin1.Append(rightMargin1);

                tableCellProperties1.Append(tableCellWidth1);
                tableCellProperties1.Append(tableCellMargin1);

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "001824A2", RsidRunAdditionDefault = "001824A2", ParagraphId = "266D37CA", TextId = "77777777" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "NoSpacing" };

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                Color color2 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
                FontSize fontSize2 = new FontSize() { Val = "24" };

                paragraphMarkRunProperties1.Append(color2);
                paragraphMarkRunProperties1.Append(fontSize2);

                paragraphProperties1.Append(paragraphStyleId1);
                paragraphProperties1.Append(paragraphMarkRunProperties1);

                Run run1 = new Run();

                RunProperties runProperties2 = new RunProperties();
                Color color3 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
                FontSize fontSize3 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "24" };

                runProperties2.Append(color3);
                runProperties2.Append(fontSize3);
                runProperties2.Append(fontSizeComplexScript2);
                Text text1 = new Text();
                text1.Text = "[Company name]";

                run1.Append(runProperties2);
                run1.Append(text1);

                paragraph2.Append(paragraphProperties1);
                paragraph2.Append(run1);

                tableCell1.Append(tableCellProperties1);
                tableCell1.Append(paragraph2);

                sdtContentCell1.Append(tableCell1);

                sdtCell1.Append(sdtProperties2);
                sdtCell1.Append(sdtContentCell1);

                tableRow1.Append(sdtCell1);

                TableRow tableRow2 = new TableRow() { RsidTableRowAddition = "001824A2", ParagraphId = "55C7924A", TextId = "77777777" };

                TableCell tableCell2 = new TableCell();

                TableCellProperties tableCellProperties2 = new TableCellProperties();
                TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "7672", Type = TableWidthUnitValues.Dxa };

                tableCellProperties2.Append(tableCellWidth2);

                SdtBlock sdtBlock2 = new SdtBlock();

                SdtProperties sdtProperties3 = new SdtProperties();

                RunProperties runProperties3 = new RunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Color color4 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize4 = new FontSize() { Val = "88" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "88" };

                runProperties3.Append(runFonts1);
                runProperties3.Append(color4);
                runProperties3.Append(fontSize4);
                runProperties3.Append(fontSizeComplexScript3);
                SdtAlias sdtAlias2 = new SdtAlias() { Val = "Title" };
                SdtId sdtId3 = new SdtId();

                SdtPlaceholder sdtPlaceholder2 = new SdtPlaceholder();
                DocPartReference docPartReference2 = new DocPartReference() { Val = "51BABD9EF9164A1B940254E8ACAC92FD" };

                sdtPlaceholder2.Append(docPartReference2);
                ShowingPlaceholder showingPlaceholder2 = new ShowingPlaceholder();
                DataBinding dataBinding2 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' xmlns:ns1=\'http://purl.org/dc/elements/1.1/\'", XPath = "/ns0:coreProperties[1]/ns1:title[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText2 = new SdtContentText();

                sdtProperties3.Append(runProperties3);
                sdtProperties3.Append(sdtAlias2);
                sdtProperties3.Append(sdtId3);
                sdtProperties3.Append(sdtPlaceholder2);
                sdtProperties3.Append(showingPlaceholder2);
                sdtProperties3.Append(dataBinding2);
                sdtProperties3.Append(sdtContentText2);

                SdtContentBlock sdtContentBlock2 = new SdtContentBlock();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "001824A2", RsidRunAdditionDefault = "001824A2", ParagraphId = "33DD8F19", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "NoSpacing" };
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = "216", LineRule = LineSpacingRuleValues.Auto };

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                RunFonts runFonts2 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Color color5 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize5 = new FontSize() { Val = "88" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "88" };

                paragraphMarkRunProperties2.Append(runFonts2);
                paragraphMarkRunProperties2.Append(color5);
                paragraphMarkRunProperties2.Append(fontSize5);
                paragraphMarkRunProperties2.Append(fontSizeComplexScript4);

                paragraphProperties2.Append(paragraphStyleId2);
                paragraphProperties2.Append(spacingBetweenLines1);
                paragraphProperties2.Append(paragraphMarkRunProperties2);

                Run run2 = new Run();

                RunProperties runProperties4 = new RunProperties();
                RunFonts runFonts3 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Color color6 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize6 = new FontSize() { Val = "88" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "88" };

                runProperties4.Append(runFonts3);
                runProperties4.Append(color6);
                runProperties4.Append(fontSize6);
                runProperties4.Append(fontSizeComplexScript5);
                Text text2 = new Text();
                text2.Text = "[Document title]";

                run2.Append(runProperties4);
                run2.Append(text2);

                paragraph3.Append(paragraphProperties2);
                paragraph3.Append(run2);

                sdtContentBlock2.Append(paragraph3);

                sdtBlock2.Append(sdtProperties3);
                sdtBlock2.Append(sdtContentBlock2);

                tableCell2.Append(tableCellProperties2);
                tableCell2.Append(sdtBlock2);

                tableRow2.Append(tableCell2);

                TableRow tableRow3 = new TableRow() { RsidTableRowAddition = "001824A2", ParagraphId = "2043CDE1", TextId = "77777777" };

                SdtCell sdtCell2 = new SdtCell();

                SdtProperties sdtProperties4 = new SdtProperties();

                RunProperties runProperties5 = new RunProperties();
                Color color7 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
                FontSize fontSize7 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "24" };

                runProperties5.Append(color7);
                runProperties5.Append(fontSize7);
                runProperties5.Append(fontSizeComplexScript6);
                SdtAlias sdtAlias3 = new SdtAlias() { Val = "Subtitle" };
                SdtId sdtId4 = new SdtId();

                SdtPlaceholder sdtPlaceholder3 = new SdtPlaceholder();
                DocPartReference docPartReference3 = new DocPartReference() { Val = "C8965A10F1FB45A584134C6BE40D8476" };

                sdtPlaceholder3.Append(docPartReference3);
                ShowingPlaceholder showingPlaceholder3 = new ShowingPlaceholder();
                DataBinding dataBinding3 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' xmlns:ns1=\'http://purl.org/dc/elements/1.1/\'", XPath = "/ns0:coreProperties[1]/ns1:subject[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText3 = new SdtContentText();

                sdtProperties4.Append(runProperties5);
                sdtProperties4.Append(sdtAlias3);
                sdtProperties4.Append(sdtId4);
                sdtProperties4.Append(sdtPlaceholder3);
                sdtProperties4.Append(showingPlaceholder3);
                sdtProperties4.Append(dataBinding3);
                sdtProperties4.Append(sdtContentText3);

                SdtContentCell sdtContentCell2 = new SdtContentCell();

                TableCell tableCell3 = new TableCell();

                TableCellProperties tableCellProperties3 = new TableCellProperties();
                TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "7672", Type = TableWidthUnitValues.Dxa };

                TableCellMargin tableCellMargin2 = new TableCellMargin();
                TopMargin topMargin2 = new TopMargin() { Width = "216", Type = TableWidthUnitValues.Dxa };
                LeftMargin leftMargin2 = new LeftMargin() { Width = "115", Type = TableWidthUnitValues.Dxa };
                BottomMargin bottomMargin2 = new BottomMargin() { Width = "216", Type = TableWidthUnitValues.Dxa };
                RightMargin rightMargin2 = new RightMargin() { Width = "115", Type = TableWidthUnitValues.Dxa };

                tableCellMargin2.Append(topMargin2);
                tableCellMargin2.Append(leftMargin2);
                tableCellMargin2.Append(bottomMargin2);
                tableCellMargin2.Append(rightMargin2);

                tableCellProperties3.Append(tableCellWidth3);
                tableCellProperties3.Append(tableCellMargin2);

                Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "001824A2", RsidRunAdditionDefault = "001824A2", ParagraphId = "64440369", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "NoSpacing" };

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                Color color8 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
                FontSize fontSize8 = new FontSize() { Val = "24" };

                paragraphMarkRunProperties3.Append(color8);
                paragraphMarkRunProperties3.Append(fontSize8);

                paragraphProperties3.Append(paragraphStyleId3);
                paragraphProperties3.Append(paragraphMarkRunProperties3);

                Run run3 = new Run();

                RunProperties runProperties6 = new RunProperties();
                Color color9 = new Color() { Val = "2F5496", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
                FontSize fontSize9 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "24" };

                runProperties6.Append(color9);
                runProperties6.Append(fontSize9);
                runProperties6.Append(fontSizeComplexScript7);
                Text text3 = new Text();
                text3.Text = "[Document subtitle]";

                run3.Append(runProperties6);
                run3.Append(text3);

                paragraph4.Append(paragraphProperties3);
                paragraph4.Append(run3);

                tableCell3.Append(tableCellProperties3);
                tableCell3.Append(paragraph4);

                sdtContentCell2.Append(tableCell3);

                sdtCell2.Append(sdtProperties4);
                sdtCell2.Append(sdtContentCell2);

                tableRow3.Append(sdtCell2);

                table1.Append(tableProperties1);
                table1.Append(tableGrid1);
                table1.Append(tableRow1);
                table1.Append(tableRow2);
                table1.Append(tableRow3);

                Table table2 = new Table();

                TableProperties tableProperties2 = new TableProperties();
                TablePositionProperties tablePositionProperties2 = new TablePositionProperties() { LeftFromText = 187, RightFromText = 187, HorizontalAnchor = HorizontalAnchorValues.Margin, TablePositionXAlignment = HorizontalAlignmentValues.Center, TablePositionYAlignment = VerticalAlignmentValues.Bottom };
                TableWidth tableWidth2 = new TableWidth() { Width = "3857", Type = TableWidthUnitValues.Pct };
                TableLook tableLook2 = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

                tableProperties2.Append(tablePositionProperties2);
                tableProperties2.Append(tableWidth2);
                tableProperties2.Append(tableLook2);

                TableGrid tableGrid2 = new TableGrid();
                GridColumn gridColumn2 = new GridColumn() { Width = "7220" };

                tableGrid2.Append(gridColumn2);

                TableRow tableRow4 = new TableRow() { RsidTableRowAddition = "001824A2", ParagraphId = "4D7F3F6C", TextId = "77777777" };

                TableCell tableCell4 = new TableCell();

                TableCellProperties tableCellProperties4 = new TableCellProperties();
                TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "7221", Type = TableWidthUnitValues.Dxa };

                TableCellMargin tableCellMargin3 = new TableCellMargin();
                TopMargin topMargin3 = new TopMargin() { Width = "216", Type = TableWidthUnitValues.Dxa };
                LeftMargin leftMargin3 = new LeftMargin() { Width = "115", Type = TableWidthUnitValues.Dxa };
                BottomMargin bottomMargin3 = new BottomMargin() { Width = "216", Type = TableWidthUnitValues.Dxa };
                RightMargin rightMargin3 = new RightMargin() { Width = "115", Type = TableWidthUnitValues.Dxa };

                tableCellMargin3.Append(topMargin3);
                tableCellMargin3.Append(leftMargin3);
                tableCellMargin3.Append(bottomMargin3);
                tableCellMargin3.Append(rightMargin3);

                tableCellProperties4.Append(tableCellWidth4);
                tableCellProperties4.Append(tableCellMargin3);

                SdtBlock sdtBlock3 = new SdtBlock();

                SdtProperties sdtProperties5 = new SdtProperties();

                RunProperties runProperties7 = new RunProperties();
                Color color10 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize10 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "28" };

                runProperties7.Append(color10);
                runProperties7.Append(fontSize10);
                runProperties7.Append(fontSizeComplexScript8);
                SdtAlias sdtAlias4 = new SdtAlias() { Val = "Author" };
                SdtId sdtId5 = new SdtId();

                SdtPlaceholder sdtPlaceholder4 = new SdtPlaceholder();
                DocPartReference docPartReference4 = new DocPartReference() { Val = "07D56EEB93E5458EB4FF8045BB57C03E" };

                sdtPlaceholder4.Append(docPartReference4);
                DataBinding dataBinding4 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' xmlns:ns1=\'http://purl.org/dc/elements/1.1/\'", XPath = "/ns0:coreProperties[1]/ns1:creator[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText4 = new SdtContentText();

                sdtProperties5.Append(runProperties7);
                sdtProperties5.Append(sdtAlias4);
                sdtProperties5.Append(sdtId5);
                sdtProperties5.Append(sdtPlaceholder4);
                sdtProperties5.Append(dataBinding4);
                sdtProperties5.Append(sdtContentText4);

                SdtContentBlock sdtContentBlock3 = new SdtContentBlock();

                Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "001824A2", RsidRunAdditionDefault = "001824A2", ParagraphId = "643DF871", TextId = "37A4B08B" };

                ParagraphProperties paragraphProperties4 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "NoSpacing" };

                ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
                Color color11 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize11 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties4.Append(color11);
                paragraphMarkRunProperties4.Append(fontSize11);
                paragraphMarkRunProperties4.Append(fontSizeComplexScript9);

                paragraphProperties4.Append(paragraphStyleId4);
                paragraphProperties4.Append(paragraphMarkRunProperties4);

                Run run4 = new Run();

                RunProperties runProperties8 = new RunProperties();
                Color color12 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize12 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "28" };

                runProperties8.Append(color12);
                runProperties8.Append(fontSize12);
                runProperties8.Append(fontSizeComplexScript10);
                Text text4 = new Text();
                text4.Text = "Przemysław Kłys";

                run4.Append(runProperties8);
                run4.Append(text4);

                paragraph5.Append(paragraphProperties4);
                paragraph5.Append(run4);

                sdtContentBlock3.Append(paragraph5);

                sdtBlock3.Append(sdtProperties5);
                sdtBlock3.Append(sdtContentBlock3);

                SdtBlock sdtBlock4 = new SdtBlock();

                SdtProperties sdtProperties6 = new SdtProperties();

                RunProperties runProperties9 = new RunProperties();
                Color color13 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize13 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "28" };

                runProperties9.Append(color13);
                runProperties9.Append(fontSize13);
                runProperties9.Append(fontSizeComplexScript11);
                SdtAlias sdtAlias5 = new SdtAlias() { Val = "Date" };
                Tag tag1 = new Tag() { Val = "Date" };
                SdtId sdtId6 = new SdtId();

                SdtPlaceholder sdtPlaceholder5 = new SdtPlaceholder();
                DocPartReference docPartReference5 = new DocPartReference() { Val = "33E01DFD9623417BB5536301C15288E7" };

                sdtPlaceholder5.Append(docPartReference5);
                ShowingPlaceholder showingPlaceholder4 = new ShowingPlaceholder();
                DataBinding dataBinding5 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.microsoft.com/office/2006/coverPageProps\'", XPath = "/ns0:CoverPageProperties[1]/ns0:PublishDate[1]", StoreItemId = "{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}" };

                SdtContentDate sdtContentDate1 = new SdtContentDate();
                DateFormat dateFormat1 = new DateFormat() { Val = "M-d-yyyy" };
                LanguageId languageId1 = new LanguageId() { Val = "en-US" };
                SdtDateMappingType sdtDateMappingType1 = new SdtDateMappingType() { Val = DateFormatValues.DateTime };
                Calendar calendar1 = new Calendar() { Val = CalendarValues.Gregorian };

                sdtContentDate1.Append(dateFormat1);
                sdtContentDate1.Append(languageId1);
                sdtContentDate1.Append(sdtDateMappingType1);
                sdtContentDate1.Append(calendar1);

                sdtProperties6.Append(runProperties9);
                sdtProperties6.Append(sdtAlias5);
                sdtProperties6.Append(tag1);
                sdtProperties6.Append(sdtId6);
                sdtProperties6.Append(sdtPlaceholder5);
                sdtProperties6.Append(showingPlaceholder4);
                sdtProperties6.Append(dataBinding5);
                sdtProperties6.Append(sdtContentDate1);

                SdtContentBlock sdtContentBlock4 = new SdtContentBlock();

                Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "001824A2", RsidRunAdditionDefault = "001824A2", ParagraphId = "387100BC", TextId = "77777777" };

                ParagraphProperties paragraphProperties5 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId5 = new ParagraphStyleId() { Val = "NoSpacing" };

                ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
                Color color14 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize14 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties5.Append(color14);
                paragraphMarkRunProperties5.Append(fontSize14);
                paragraphMarkRunProperties5.Append(fontSizeComplexScript12);

                paragraphProperties5.Append(paragraphStyleId5);
                paragraphProperties5.Append(paragraphMarkRunProperties5);

                Run run5 = new Run();

                RunProperties runProperties10 = new RunProperties();
                Color color15 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize15 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "28" };

                runProperties10.Append(color15);
                runProperties10.Append(fontSize15);
                runProperties10.Append(fontSizeComplexScript13);
                Text text5 = new Text();
                text5.Text = "[Date]";

                run5.Append(runProperties10);
                run5.Append(text5);

                paragraph6.Append(paragraphProperties5);
                paragraph6.Append(run5);

                sdtContentBlock4.Append(paragraph6);

                sdtBlock4.Append(sdtProperties6);
                sdtBlock4.Append(sdtContentBlock4);

                Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "001824A2", RsidRunAdditionDefault = "001824A2", ParagraphId = "62360589", TextId = "77777777" };

                ParagraphProperties paragraphProperties6 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId6 = new ParagraphStyleId() { Val = "NoSpacing" };

                ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
                Color color16 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };

                paragraphMarkRunProperties6.Append(color16);

                paragraphProperties6.Append(paragraphStyleId6);
                paragraphProperties6.Append(paragraphMarkRunProperties6);

                paragraph7.Append(paragraphProperties6);

                tableCell4.Append(tableCellProperties4);
                tableCell4.Append(sdtBlock3);
                tableCell4.Append(sdtBlock4);
                tableCell4.Append(paragraph7);

                tableRow4.Append(tableCell4);

                table2.Append(tableProperties2);
                table2.Append(tableGrid2);
                table2.Append(tableRow4);

                Paragraph paragraph8 = new Paragraph() { RsidParagraphAddition = "001824A2", RsidRunAdditionDefault = "001824A2", ParagraphId = "542FB46A", TextId = "4BC3BE4E" };

                Run run6 = new Run();
                Break break1 = new Break() { Type = BreakValues.Page };

                run6.Append(break1);

                paragraph8.Append(run6);

                sdtContentBlock1.Append(paragraph1);
                sdtContentBlock1.Append(table1);
                sdtContentBlock1.Append(table2);
                sdtContentBlock1.Append(paragraph8);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
    }
}
