using DocumentFormat.OpenXml.Wordprocessing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a cover page within a Word document.
    /// </summary>
    public partial class WordCoverPage {

        private static SdtBlock CoverPageElement {
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
                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "006F1781", RsidRunAdditionDefault = "00C41751", ParagraphId = "2CCCF375", TextId = "225807DC" };

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                Picture picture1 = new Picture() { AnchorId = "7BAC436C" };

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
                V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

                shapetype1.Append(stroke1);
                shapetype1.Append(path1);

                V.Shape shape1 = new V.Shape() { Id = "Text Box 1", Style = "position:absolute;margin-left:0;margin-top:0;width:553.9pt;height:256.3pt;z-index:-251658752;visibility:visible;mso-wrap-style:square;mso-width-percent:906;mso-height-percent:0;mso-top-percent:510;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;mso-width-percent:906;mso-height-percent:0;mso-top-percent:510;mso-width-relative:page;mso-height-relative:margin;v-text-anchor:top", Alternate = "Cover page content layout", OptionalString = "_x0000_s1026", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "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\ncmV2LnhtbFBLBQYAAAAABAAEAPMAAADCBQAAAAA=\n" };

                V.TextBox textBox1 = new V.TextBox() { Style = "mso-fit-shape-to-text:t", Inset = "0,0,0,0" };

                TextBoxContent textBoxContent1 = new TextBoxContent();

                Table table1 = new Table();

                TableProperties tableProperties1 = new TableProperties();
                TableWidth tableWidth1 = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };

                TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
                TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 0, Type = TableWidthValues.Dxa };
                TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 0, Type = TableWidthValues.Dxa };

                tableCellMarginDefault1.Append(tableCellLeftMargin1);
                tableCellMarginDefault1.Append(tableCellRightMargin1);
                TableLook tableLook1 = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

                tableProperties1.Append(tableWidth1);
                tableProperties1.Append(tableCellMarginDefault1);
                tableProperties1.Append(tableLook1);

                TableGrid tableGrid1 = new TableGrid();
                GridColumn gridColumn1 = new GridColumn() { Width = "821" };
                GridColumn gridColumn2 = new GridColumn() { Width = "10272" };

                tableGrid1.Append(gridColumn1);
                tableGrid1.Append(gridColumn2);

                TableRow tableRow1 = new TableRow() { RsidTableRowAddition = "006F1781", ParagraphId = "4F40A134", TextId = "77777777" };

                TableRowProperties tableRowProperties1 = new TableRowProperties();
                TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)2376U };

                tableRowProperties1.Append(tableRowHeight1);

                TableCell tableCell1 = new TableCell();

                TableCellProperties tableCellProperties1 = new TableCellProperties();
                TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "370", Type = TableWidthUnitValues.Pct };
                Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "4472C4", ThemeFill = ThemeColorValues.Accent1 };

                tableCellProperties1.Append(tableCellWidth1);
                tableCellProperties1.Append(shading1);
                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "006F1781", RsidRunAdditionDefault = "00C41751", ParagraphId = "45BE470C", TextId = "77777777" };

                tableCell1.Append(tableCellProperties1);
                tableCell1.Append(paragraph2);

                SdtCell sdtCell1 = new SdtCell();

                SdtProperties sdtProperties2 = new SdtProperties();

                RunProperties runProperties2 = new RunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi };
                Color color1 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize1 = new FontSize() { Val = "96" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "96" };

                runProperties2.Append(runFonts1);
                runProperties2.Append(color1);
                runProperties2.Append(fontSize1);
                runProperties2.Append(fontSizeComplexScript1);
                SdtAlias sdtAlias1 = new SdtAlias() { Val = "Title" };
                Tag tag1 = new Tag() { Val = "" };
                SdtId sdtId2 = new SdtId();
                ShowingPlaceholder showingPlaceholder1 = new ShowingPlaceholder();
                DataBinding dataBinding1 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:title[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText1 = new SdtContentText();

                sdtProperties2.Append(runProperties2);
                sdtProperties2.Append(sdtAlias1);
                sdtProperties2.Append(tag1);
                sdtProperties2.Append(sdtId2);
                sdtProperties2.Append(showingPlaceholder1);
                sdtProperties2.Append(dataBinding1);
                sdtProperties2.Append(sdtContentText1);
                SdtEndCharProperties sdtEndCharProperties2 = new SdtEndCharProperties();

                SdtContentCell sdtContentCell1 = new SdtContentCell();

                TableCell tableCell2 = new TableCell();

                TableCellProperties tableCellProperties2 = new TableCellProperties();
                TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "4630", Type = TableWidthUnitValues.Pct };
                Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "404040", ThemeFill = ThemeColorValues.Text1, ThemeFillTint = "BF" };

                tableCellProperties2.Append(tableCellWidth2);
                tableCellProperties2.Append(shading2);

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "006F1781", RsidRunAdditionDefault = "00C41751", ParagraphId = "2FB85FD1", TextId = "77777777" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240", Line = "216", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation1 = new Indentation() { Left = "360", Right = "360" };
                ContextualSpacing contextualSpacing1 = new ContextualSpacing();

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                RunFonts runFonts2 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi };
                Color color2 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize2 = new FontSize() { Val = "96" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "96" };

                paragraphMarkRunProperties1.Append(runFonts2);
                paragraphMarkRunProperties1.Append(color2);
                paragraphMarkRunProperties1.Append(fontSize2);
                paragraphMarkRunProperties1.Append(fontSizeComplexScript2);

                paragraphProperties1.Append(spacingBetweenLines1);
                paragraphProperties1.Append(indentation1);
                paragraphProperties1.Append(contextualSpacing1);
                paragraphProperties1.Append(paragraphMarkRunProperties1);

                Run run2 = new Run();

                RunProperties runProperties3 = new RunProperties();
                RunFonts runFonts3 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi };
                Color color3 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize3 = new FontSize() { Val = "96" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "96" };

                runProperties3.Append(runFonts3);
                runProperties3.Append(color3);
                runProperties3.Append(fontSize3);
                runProperties3.Append(fontSizeComplexScript3);
                Text text1 = new Text();
                text1.Text = "[Document title]";

                run2.Append(runProperties3);
                run2.Append(text1);

                paragraph3.Append(paragraphProperties1);
                paragraph3.Append(run2);

                tableCell2.Append(tableCellProperties2);
                tableCell2.Append(paragraph3);

                sdtContentCell1.Append(tableCell2);

                sdtCell1.Append(sdtProperties2);
                sdtCell1.Append(sdtEndCharProperties2);
                sdtCell1.Append(sdtContentCell1);

                tableRow1.Append(tableRowProperties1);
                tableRow1.Append(tableCell1);
                tableRow1.Append(sdtCell1);

                TableRow tableRow2 = new TableRow() { RsidTableRowAddition = "006F1781", ParagraphId = "519110BA", TextId = "77777777" };

                TableRowProperties tableRowProperties2 = new TableRowProperties();
                TableRowHeight tableRowHeight2 = new TableRowHeight() { Val = (UInt32Value)648U, HeightType = HeightRuleValues.Exact };

                tableRowProperties2.Append(tableRowHeight2);

                TableCell tableCell3 = new TableCell();

                TableCellProperties tableCellProperties3 = new TableCellProperties();
                TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "370", Type = TableWidthUnitValues.Pct };
                Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "4472C4", ThemeFill = ThemeColorValues.Accent1 };

                tableCellProperties3.Append(tableCellWidth3);
                tableCellProperties3.Append(shading3);
                Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "006F1781", RsidRunAdditionDefault = "00C41751", ParagraphId = "7D751E48", TextId = "77777777" };

                tableCell3.Append(tableCellProperties3);
                tableCell3.Append(paragraph4);

                TableCell tableCell4 = new TableCell();

                TableCellProperties tableCellProperties4 = new TableCellProperties();
                TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "4630", Type = TableWidthUnitValues.Pct };
                Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "404040", ThemeFill = ThemeColorValues.Text1, ThemeFillTint = "BF" };
                TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Bottom };

                tableCellProperties4.Append(tableCellWidth4);
                tableCellProperties4.Append(shading4);
                tableCellProperties4.Append(tableCellVerticalAlignment1);

                Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "006F1781", RsidRunAdditionDefault = "00C41751", ParagraphId = "11179A8F", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                Indentation indentation2 = new Indentation() { Left = "360", Right = "360" };

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                Color color4 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize4 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties2.Append(color4);
                paragraphMarkRunProperties2.Append(fontSize4);
                paragraphMarkRunProperties2.Append(fontSizeComplexScript4);

                paragraphProperties2.Append(indentation2);
                paragraphProperties2.Append(paragraphMarkRunProperties2);

                paragraph5.Append(paragraphProperties2);

                tableCell4.Append(tableCellProperties4);
                tableCell4.Append(paragraph5);

                tableRow2.Append(tableRowProperties2);
                tableRow2.Append(tableCell3);
                tableRow2.Append(tableCell4);

                TableRow tableRow3 = new TableRow() { RsidTableRowAddition = "006F1781", ParagraphId = "1DA3945C", TextId = "77777777" };

                TableCell tableCell5 = new TableCell();

                TableCellProperties tableCellProperties5 = new TableCellProperties();
                TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "370", Type = TableWidthUnitValues.Pct };
                Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "4472C4", ThemeFill = ThemeColorValues.Accent1 };

                tableCellProperties5.Append(tableCellWidth5);
                tableCellProperties5.Append(shading5);
                Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "006F1781", RsidRunAdditionDefault = "00C41751", ParagraphId = "453B92ED", TextId = "77777777" };

                tableCell5.Append(tableCellProperties5);
                tableCell5.Append(paragraph6);

                TableCell tableCell6 = new TableCell();

                TableCellProperties tableCellProperties6 = new TableCellProperties();
                TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "4630", Type = TableWidthUnitValues.Pct };
                Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "404040", ThemeFill = ThemeColorValues.Text1, ThemeFillTint = "BF" };
                TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Bottom };

                tableCellProperties6.Append(tableCellWidth6);
                tableCellProperties6.Append(shading6);
                tableCellProperties6.Append(tableCellVerticalAlignment2);

                Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "006F1781", RsidRunAdditionDefault = "00C41751", ParagraphId = "19806265", TextId = "6804FBDF" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Line = "288", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation3 = new Indentation() { Left = "360", Right = "360" };

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                Color color5 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize5 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties3.Append(color5);
                paragraphMarkRunProperties3.Append(fontSize5);
                paragraphMarkRunProperties3.Append(fontSizeComplexScript5);

                paragraphProperties3.Append(spacingBetweenLines2);
                paragraphProperties3.Append(indentation3);
                paragraphProperties3.Append(paragraphMarkRunProperties3);

                SdtRun sdtRun1 = new SdtRun();

                SdtProperties sdtProperties3 = new SdtProperties();

                RunProperties runProperties4 = new RunProperties();
                Color color6 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize6 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "28" };

                runProperties4.Append(color6);
                runProperties4.Append(fontSize6);
                runProperties4.Append(fontSizeComplexScript6);
                SdtAlias sdtAlias2 = new SdtAlias() { Val = "Author" };
                Tag tag2 = new Tag() { Val = "" };
                SdtId sdtId3 = new SdtId();
                ShowingPlaceholder showingPlaceholder2 = new ShowingPlaceholder();
                DataBinding dataBinding2 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns0:creator[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText2 = new SdtContentText();

                sdtProperties3.Append(runProperties4);
                sdtProperties3.Append(sdtAlias2);
                sdtProperties3.Append(tag2);
                sdtProperties3.Append(sdtId3);
                sdtProperties3.Append(showingPlaceholder2);
                sdtProperties3.Append(dataBinding2);
                sdtProperties3.Append(sdtContentText2);
                SdtEndCharProperties sdtEndCharProperties3 = new SdtEndCharProperties();

                SdtContentRun sdtContentRun1 = new SdtContentRun();

                Run run3 = new Run();

                RunProperties runProperties5 = new RunProperties();
                Color color7 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize7 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "28" };

                runProperties5.Append(color7);
                runProperties5.Append(fontSize7);
                runProperties5.Append(fontSizeComplexScript7);
                Text text2 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text2.Text = "     ";

                run3.Append(runProperties5);
                run3.Append(text2);

                sdtContentRun1.Append(run3);

                sdtRun1.Append(sdtProperties3);
                sdtRun1.Append(sdtEndCharProperties3);
                sdtRun1.Append(sdtContentRun1);

                paragraph7.Append(paragraphProperties3);
                paragraph7.Append(sdtRun1);

                SdtBlock sdtBlock2 = new SdtBlock();

                SdtProperties sdtProperties4 = new SdtProperties();

                RunProperties runProperties6 = new RunProperties();
                Color color8 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize8 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "28" };

                runProperties6.Append(color8);
                runProperties6.Append(fontSize8);
                runProperties6.Append(fontSizeComplexScript8);
                SdtAlias sdtAlias3 = new SdtAlias() { Val = "Course title" };
                Tag tag3 = new Tag() { Val = "" };
                SdtId sdtId4 = new SdtId();
                ShowingPlaceholder showingPlaceholder3 = new ShowingPlaceholder();
                DataBinding dataBinding3 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://purl.org/dc/elements/1.1/\' xmlns:ns1=\'http://schemas.openxmlformats.org/package/2006/metadata/core-properties\' ", XPath = "/ns1:coreProperties[1]/ns1:category[1]", StoreItemId = "{6C3C8BC8-F283-45AE-878A-BAB7291924A1}" };
                SdtContentText sdtContentText3 = new SdtContentText();

                sdtProperties4.Append(runProperties6);
                sdtProperties4.Append(sdtAlias3);
                sdtProperties4.Append(tag3);
                sdtProperties4.Append(sdtId4);
                sdtProperties4.Append(showingPlaceholder3);
                sdtProperties4.Append(dataBinding3);
                sdtProperties4.Append(sdtContentText3);
                SdtEndCharProperties sdtEndCharProperties4 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock2 = new SdtContentBlock();

                Paragraph paragraph8 = new Paragraph() { RsidParagraphAddition = "006F1781", RsidRunAdditionDefault = "00C41751", ParagraphId = "6D4B7D6D", TextId = "77777777" };

                ParagraphProperties paragraphProperties4 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Line = "288", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation4 = new Indentation() { Left = "360", Right = "360" };

                ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
                Color color9 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize9 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties4.Append(color9);
                paragraphMarkRunProperties4.Append(fontSize9);
                paragraphMarkRunProperties4.Append(fontSizeComplexScript9);

                paragraphProperties4.Append(spacingBetweenLines3);
                paragraphProperties4.Append(indentation4);
                paragraphProperties4.Append(paragraphMarkRunProperties4);

                Run run4 = new Run();

                RunProperties runProperties7 = new RunProperties();
                Color color10 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize10 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "28" };

                runProperties7.Append(color10);
                runProperties7.Append(fontSize10);
                runProperties7.Append(fontSizeComplexScript10);
                Text text3 = new Text();
                text3.Text = "[Course title]";

                run4.Append(runProperties7);
                run4.Append(text3);

                paragraph8.Append(paragraphProperties4);
                paragraph8.Append(run4);

                sdtContentBlock2.Append(paragraph8);

                sdtBlock2.Append(sdtProperties4);
                sdtBlock2.Append(sdtEndCharProperties4);
                sdtBlock2.Append(sdtContentBlock2);

                SdtBlock sdtBlock3 = new SdtBlock();

                SdtProperties sdtProperties5 = new SdtProperties();

                RunProperties runProperties8 = new RunProperties();
                Color color11 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize11 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "28" };

                runProperties8.Append(color11);
                runProperties8.Append(fontSize11);
                runProperties8.Append(fontSizeComplexScript11);
                SdtAlias sdtAlias4 = new SdtAlias() { Val = "Date" };
                Tag tag4 = new Tag() { Val = "" };
                SdtId sdtId5 = new SdtId();
                ShowingPlaceholder showingPlaceholder4 = new ShowingPlaceholder();
                DataBinding dataBinding4 = new DataBinding() { PrefixMappings = "xmlns:ns0=\'http://schemas.microsoft.com/office/2006/coverPageProps\' ", XPath = "/ns0:CoverPageProperties[1]/ns0:PublishDate[1]", StoreItemId = "{55AF091B-3C7A-41E3-B477-F2FDAA23CFDA}" };

                SdtContentDate sdtContentDate1 = new SdtContentDate();
                DateFormat dateFormat1 = new DateFormat() { Val = "M/d/yy" };
                LanguageId languageId1 = new LanguageId() { Val = "en-US" };
                SdtDateMappingType sdtDateMappingType1 = new SdtDateMappingType() { Val = DateFormatValues.DateTime };
                Calendar calendar1 = new Calendar() { Val = CalendarValues.Gregorian };

                sdtContentDate1.Append(dateFormat1);
                sdtContentDate1.Append(languageId1);
                sdtContentDate1.Append(sdtDateMappingType1);
                sdtContentDate1.Append(calendar1);

                sdtProperties5.Append(runProperties8);
                sdtProperties5.Append(sdtAlias4);
                sdtProperties5.Append(tag4);
                sdtProperties5.Append(sdtId5);
                sdtProperties5.Append(showingPlaceholder4);
                sdtProperties5.Append(dataBinding4);
                sdtProperties5.Append(sdtContentDate1);
                SdtEndCharProperties sdtEndCharProperties5 = new SdtEndCharProperties();

                SdtContentBlock sdtContentBlock3 = new SdtContentBlock();

                Paragraph paragraph9 = new Paragraph() { RsidParagraphAddition = "006F1781", RsidRunAdditionDefault = "00C41751", ParagraphId = "0417BC68", TextId = "77777777" };

                ParagraphProperties paragraphProperties5 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { After = "240", Line = "288", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation5 = new Indentation() { Left = "360", Right = "360" };

                ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
                Color color12 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize12 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties5.Append(color12);
                paragraphMarkRunProperties5.Append(fontSize12);
                paragraphMarkRunProperties5.Append(fontSizeComplexScript12);

                paragraphProperties5.Append(spacingBetweenLines4);
                paragraphProperties5.Append(indentation5);
                paragraphProperties5.Append(paragraphMarkRunProperties5);

                Run run5 = new Run();

                RunProperties runProperties9 = new RunProperties();
                Color color13 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize13 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "28" };

                runProperties9.Append(color13);
                runProperties9.Append(fontSize13);
                runProperties9.Append(fontSizeComplexScript13);
                Text text4 = new Text();
                text4.Text = "[Date]";

                run5.Append(runProperties9);
                run5.Append(text4);

                paragraph9.Append(paragraphProperties5);
                paragraph9.Append(run5);

                sdtContentBlock3.Append(paragraph9);

                sdtBlock3.Append(sdtProperties5);
                sdtBlock3.Append(sdtEndCharProperties5);
                sdtBlock3.Append(sdtContentBlock3);

                tableCell6.Append(tableCellProperties6);
                tableCell6.Append(paragraph7);
                tableCell6.Append(sdtBlock2);
                tableCell6.Append(sdtBlock3);

                tableRow3.Append(tableCell5);
                tableRow3.Append(tableCell6);

                table1.Append(tableProperties1);
                table1.Append(tableGrid1);
                table1.Append(tableRow1);
                table1.Append(tableRow2);
                table1.Append(tableRow3);
                Paragraph paragraph10 = new Paragraph() { RsidParagraphAddition = "006F1781", RsidRunAdditionDefault = "00C41751", ParagraphId = "277F5834", TextId = "77777777" };

                textBoxContent1.Append(table1);
                textBoxContent1.Append(paragraph10);

                textBox1.Append(textBoxContent1);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

                shape1.Append(textBox1);
                shape1.Append(textWrap1);

                picture1.Append(shapetype1);
                picture1.Append(shape1);

                run1.Append(runProperties1);
                run1.Append(picture1);

                paragraph1.Append(run1);

                Paragraph paragraph11 = new Paragraph() { RsidParagraphAddition = "006F1781", RsidRunAdditionDefault = "00C41751", ParagraphId = "3A3E7D25", TextId = "0A010A30" };

                Run run6 = new Run();
                Break break1 = new Break() { Type = BreakValues.Page };

                run6.Append(break1);

                paragraph11.Append(run6);

                sdtContentBlock1.Append(paragraph1);
                sdtContentBlock1.Append(paragraph11);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
    }
}
