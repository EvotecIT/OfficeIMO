using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public enum WordListStyle {
        Bulleted,
        ArticleSections,
        Headings111,
        HeadingIA1,
        Chapters,
        BulletedChars,
        Heading1ai,
        Headings111Shifted
    }
    public static class WordListStyles {
        public static AbstractNum GetStyle(WordListStyle style) {
            switch (style) {
                case WordListStyle.Bulleted: return Bulleted;
                case WordListStyle.ArticleSections: return ArticleSections;
                case WordListStyle.Headings111: return Headings111;
                case WordListStyle.HeadingIA1: return HeadingIA1;
                case WordListStyle.Chapters: return Chapters;
                case WordListStyle.BulletedChars: return BulletedChars;
                case WordListStyle.Heading1ai: return Heading1ai;
                case WordListStyle.Headings111Shifted: return Headings111Shifted;
            }
            throw new ArgumentOutOfRangeException(nameof(style));
        }

        /// <summary>
        /// Generates a random NSID value
        /// </summary>
        /// <returns></returns>
        private static string GenerateNsidValue() {
            Random random = new Random();
            int randomNumber = random.Next();
            string nsidValue = randomNumber.ToString("X8");
            return nsidValue;
        }

        private static AbstractNum ArticleSections {
            get {
                AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 0 };
                abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
                Nsid nsid1 = new Nsid() { Val = GenerateNsidValue() };
                MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
                TemplateCode templateCode1 = new TemplateCode() { Val = "04090023" };

                Level level1 = new Level() { LevelIndex = 0 };
                StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.UpperRoman };
                ParagraphStyleIdInLevel paragraphStyleIdInLevel1 = new ParagraphStyleIdInLevel() { Val = "Heading1" };
                LevelText levelText1 = new LevelText() { Val = "Article %1." };
                LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
                Indentation indentation1 = new Indentation() { Left = "0", FirstLine = "0" };

                previousParagraphProperties1.Append(indentation1);

                level1.Append(startNumberingValue1);
                level1.Append(numberingFormat1);
                level1.Append(paragraphStyleIdInLevel1);
                level1.Append(levelText1);
                level1.Append(levelJustification1);
                level1.Append(previousParagraphProperties1);

                Level level2 = new Level() { LevelIndex = 1 };
                StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.DecimalZero };
                ParagraphStyleIdInLevel paragraphStyleIdInLevel2 = new ParagraphStyleIdInLevel() { Val = "Heading2" };
                IsLegalNumberingStyle isLegalNumberingStyle1 = new IsLegalNumberingStyle();
                LevelText levelText2 = new LevelText() { Val = "Section %1.%2" };
                LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
                Indentation indentation2 = new Indentation() { Left = "0", FirstLine = "0" };

                previousParagraphProperties2.Append(indentation2);

                level2.Append(startNumberingValue2);
                level2.Append(numberingFormat2);
                level2.Append(paragraphStyleIdInLevel2);
                level2.Append(isLegalNumberingStyle1);
                level2.Append(levelText2);
                level2.Append(levelJustification2);
                level2.Append(previousParagraphProperties2);

                Level level3 = new Level() { LevelIndex = 2 };
                StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
                ParagraphStyleIdInLevel paragraphStyleIdInLevel3 = new ParagraphStyleIdInLevel() { Val = "Heading3" };
                LevelText levelText3 = new LevelText() { Val = "(%3)" };
                LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
                Indentation indentation3 = new Indentation() { Left = "720", Hanging = "432" };

                previousParagraphProperties3.Append(indentation3);

                level3.Append(startNumberingValue3);
                level3.Append(numberingFormat3);
                level3.Append(paragraphStyleIdInLevel3);
                level3.Append(levelText3);
                level3.Append(levelJustification3);
                level3.Append(previousParagraphProperties3);

                Level level4 = new Level() { LevelIndex = 3 };
                StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
                ParagraphStyleIdInLevel paragraphStyleIdInLevel4 = new ParagraphStyleIdInLevel() { Val = "Heading4" };
                LevelText levelText4 = new LevelText() { Val = "(%4)" };
                LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Right };

                PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
                Indentation indentation4 = new Indentation() { Left = "864", Hanging = "144" };

                previousParagraphProperties4.Append(indentation4);

                level4.Append(startNumberingValue4);
                level4.Append(numberingFormat4);
                level4.Append(paragraphStyleIdInLevel4);
                level4.Append(levelText4);
                level4.Append(levelJustification4);
                level4.Append(previousParagraphProperties4);

                Level level5 = new Level() { LevelIndex = 4 };
                StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                ParagraphStyleIdInLevel paragraphStyleIdInLevel5 = new ParagraphStyleIdInLevel() { Val = "Heading5" };
                LevelText levelText5 = new LevelText() { Val = "%5)" };
                LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
                Indentation indentation5 = new Indentation() { Left = "1008", Hanging = "432" };

                previousParagraphProperties5.Append(indentation5);

                level5.Append(startNumberingValue5);
                level5.Append(numberingFormat5);
                level5.Append(paragraphStyleIdInLevel5);
                level5.Append(levelText5);
                level5.Append(levelJustification5);
                level5.Append(previousParagraphProperties5);

                Level level6 = new Level() { LevelIndex = 5 };
                StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
                ParagraphStyleIdInLevel paragraphStyleIdInLevel6 = new ParagraphStyleIdInLevel() { Val = "Heading6" };
                LevelText levelText6 = new LevelText() { Val = "%6)" };
                LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
                Indentation indentation6 = new Indentation() { Left = "1152", Hanging = "432" };

                previousParagraphProperties6.Append(indentation6);

                level6.Append(startNumberingValue6);
                level6.Append(numberingFormat6);
                level6.Append(paragraphStyleIdInLevel6);
                level6.Append(levelText6);
                level6.Append(levelJustification6);
                level6.Append(previousParagraphProperties6);

                Level level7 = new Level() { LevelIndex = 6 };
                StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
                ParagraphStyleIdInLevel paragraphStyleIdInLevel7 = new ParagraphStyleIdInLevel() { Val = "Heading7" };
                LevelText levelText7 = new LevelText() { Val = "%7)" };
                LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Right };

                PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
                Indentation indentation7 = new Indentation() { Left = "1296", Hanging = "288" };

                previousParagraphProperties7.Append(indentation7);

                level7.Append(startNumberingValue7);
                level7.Append(numberingFormat7);
                level7.Append(paragraphStyleIdInLevel7);
                level7.Append(levelText7);
                level7.Append(levelJustification7);
                level7.Append(previousParagraphProperties7);

                Level level8 = new Level() { LevelIndex = 7 };
                StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
                ParagraphStyleIdInLevel paragraphStyleIdInLevel8 = new ParagraphStyleIdInLevel() { Val = "Heading8" };
                LevelText levelText8 = new LevelText() { Val = "%8." };
                LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
                Indentation indentation8 = new Indentation() { Left = "1440", Hanging = "432" };

                previousParagraphProperties8.Append(indentation8);

                level8.Append(startNumberingValue8);
                level8.Append(numberingFormat8);
                level8.Append(paragraphStyleIdInLevel8);
                level8.Append(levelText8);
                level8.Append(levelJustification8);
                level8.Append(previousParagraphProperties8);

                Level level9 = new Level() { LevelIndex = 8 };
                StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
                ParagraphStyleIdInLevel paragraphStyleIdInLevel9 = new ParagraphStyleIdInLevel() { Val = "Heading9" };
                LevelText levelText9 = new LevelText() { Val = "%9." };
                LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Right };

                PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
                Indentation indentation9 = new Indentation() { Left = "1584", Hanging = "144" };

                previousParagraphProperties9.Append(indentation9);

                level9.Append(startNumberingValue9);
                level9.Append(numberingFormat9);
                level9.Append(paragraphStyleIdInLevel9);
                level9.Append(levelText9);
                level9.Append(levelJustification9);
                level9.Append(previousParagraphProperties9);

                abstractNum1.Append(nsid1);
                abstractNum1.Append(multiLevelType1);
                abstractNum1.Append(templateCode1);
                abstractNum1.Append(level1);
                abstractNum1.Append(level2);
                abstractNum1.Append(level3);
                abstractNum1.Append(level4);
                abstractNum1.Append(level5);
                abstractNum1.Append(level6);
                abstractNum1.Append(level7);
                abstractNum1.Append(level8);
                abstractNum1.Append(level9);
                return abstractNum1;

            }
        }

        private static AbstractNum Headings111 {
            get {
                AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 1 };
                abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
                Nsid nsid1 = new Nsid() { Val = GenerateNsidValue() };
                MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
                TemplateCode templateCode1 = new TemplateCode() { Val = "04090025" };

                Level level1 = new Level() { LevelIndex = 0 };
                StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText1 = new LevelText() { Val = "%1" };
                LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
                Indentation indentation1 = new Indentation() { Left = "432", Hanging = "432" };

                previousParagraphProperties1.Append(indentation1);

                level1.Append(startNumberingValue1);
                level1.Append(numberingFormat1);
                level1.Append(levelText1);
                level1.Append(levelJustification1);
                level1.Append(previousParagraphProperties1);

                Level level2 = new Level() { LevelIndex = 1 };
                StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText2 = new LevelText() { Val = "%1.%2" };
                LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
                Indentation indentation2 = new Indentation() { Left = "576", Hanging = "576" };

                previousParagraphProperties2.Append(indentation2);

                level2.Append(startNumberingValue2);
                level2.Append(numberingFormat2);
                level2.Append(levelText2);
                level2.Append(levelJustification2);
                level2.Append(previousParagraphProperties2);

                Level level3 = new Level() { LevelIndex = 2 };
                StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText3 = new LevelText() { Val = "%1.%2.%3" };
                LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
                Indentation indentation3 = new Indentation() { Left = "720", Hanging = "720" };

                previousParagraphProperties3.Append(indentation3);

                level3.Append(startNumberingValue3);
                level3.Append(numberingFormat3);
                level3.Append(levelText3);
                level3.Append(levelJustification3);
                level3.Append(previousParagraphProperties3);

                Level level4 = new Level() { LevelIndex = 3 };
                StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText4 = new LevelText() { Val = "%1.%2.%3.%4" };
                LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
                Indentation indentation4 = new Indentation() { Left = "864", Hanging = "864" };

                previousParagraphProperties4.Append(indentation4);

                level4.Append(startNumberingValue4);
                level4.Append(numberingFormat4);
                level4.Append(levelText4);
                level4.Append(levelJustification4);
                level4.Append(previousParagraphProperties4);

                Level level5 = new Level() { LevelIndex = 4 };
                StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText5 = new LevelText() { Val = "%1.%2.%3.%4.%5" };
                LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
                Indentation indentation5 = new Indentation() { Left = "1008", Hanging = "1008" };

                previousParagraphProperties5.Append(indentation5);

                level5.Append(startNumberingValue5);
                level5.Append(numberingFormat5);
                level5.Append(levelText5);
                level5.Append(levelJustification5);
                level5.Append(previousParagraphProperties5);

                Level level6 = new Level() { LevelIndex = 5 };
                StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText6 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6" };
                LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
                Indentation indentation6 = new Indentation() { Left = "1152", Hanging = "1152" };

                previousParagraphProperties6.Append(indentation6);

                level6.Append(startNumberingValue6);
                level6.Append(numberingFormat6);
                level6.Append(levelText6);
                level6.Append(levelJustification6);
                level6.Append(previousParagraphProperties6);

                Level level7 = new Level() { LevelIndex = 6 };
                StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText7 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7" };
                LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
                Indentation indentation7 = new Indentation() { Left = "1296", Hanging = "1296" };

                previousParagraphProperties7.Append(indentation7);

                level7.Append(startNumberingValue7);
                level7.Append(numberingFormat7);
                level7.Append(levelText7);
                level7.Append(levelJustification7);
                level7.Append(previousParagraphProperties7);

                Level level8 = new Level() { LevelIndex = 7 };
                StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText8 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8" };
                LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
                Indentation indentation8 = new Indentation() { Left = "1440", Hanging = "1440" };

                previousParagraphProperties8.Append(indentation8);

                level8.Append(startNumberingValue8);
                level8.Append(numberingFormat8);
                level8.Append(levelText8);
                level8.Append(levelJustification8);
                level8.Append(previousParagraphProperties8);

                Level level9 = new Level() { LevelIndex = 8 };
                StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText9 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8.%9" };
                LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
                Indentation indentation9 = new Indentation() { Left = "1584", Hanging = "1584" };

                previousParagraphProperties9.Append(indentation9);

                level9.Append(startNumberingValue9);
                level9.Append(numberingFormat9);
                level9.Append(levelText9);
                level9.Append(levelJustification9);
                level9.Append(previousParagraphProperties9);

                abstractNum1.Append(nsid1);
                abstractNum1.Append(multiLevelType1);
                abstractNum1.Append(templateCode1);
                abstractNum1.Append(level1);
                abstractNum1.Append(level2);
                abstractNum1.Append(level3);
                abstractNum1.Append(level4);
                abstractNum1.Append(level5);
                abstractNum1.Append(level6);
                abstractNum1.Append(level7);
                abstractNum1.Append(level8);
                abstractNum1.Append(level9);
                return abstractNum1;
            }
        }

        private static AbstractNum HeadingIA1 {
            get {
                AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 2 };
                abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
                Nsid nsid1 = new Nsid() { Val = GenerateNsidValue() };
                MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
                TemplateCode templateCode1 = new TemplateCode() { Val = "04090027" };

                Level level1 = new Level() { LevelIndex = 0 };
                StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.UpperRoman };
                LevelText levelText1 = new LevelText() { Val = "%1." };
                LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
                Indentation indentation1 = new Indentation() { Left = "0", FirstLine = "0" };

                previousParagraphProperties1.Append(indentation1);

                level1.Append(startNumberingValue1);
                level1.Append(numberingFormat1);
                level1.Append(levelText1);
                level1.Append(levelJustification1);
                level1.Append(previousParagraphProperties1);

                Level level2 = new Level() { LevelIndex = 1 };
                StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.UpperLetter };
                LevelText levelText2 = new LevelText() { Val = "%2." };
                LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
                Indentation indentation2 = new Indentation() { Left = "720", FirstLine = "0" };

                previousParagraphProperties2.Append(indentation2);

                level2.Append(startNumberingValue2);
                level2.Append(numberingFormat2);
                level2.Append(levelText2);
                level2.Append(levelJustification2);
                level2.Append(previousParagraphProperties2);

                Level level3 = new Level() { LevelIndex = 2 };
                StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText3 = new LevelText() { Val = "%3." };
                LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
                Indentation indentation3 = new Indentation() { Left = "1440", FirstLine = "0" };

                previousParagraphProperties3.Append(indentation3);

                level3.Append(startNumberingValue3);
                level3.Append(numberingFormat3);
                level3.Append(levelText3);
                level3.Append(levelJustification3);
                level3.Append(previousParagraphProperties3);

                Level level4 = new Level() { LevelIndex = 3 };
                StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
                LevelText levelText4 = new LevelText() { Val = "%4)" };
                LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
                Indentation indentation4 = new Indentation() { Left = "2160", FirstLine = "0" };

                previousParagraphProperties4.Append(indentation4);

                level4.Append(startNumberingValue4);
                level4.Append(numberingFormat4);
                level4.Append(levelText4);
                level4.Append(levelJustification4);
                level4.Append(previousParagraphProperties4);

                Level level5 = new Level() { LevelIndex = 4 };
                StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText5 = new LevelText() { Val = "(%5)" };
                LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
                Indentation indentation5 = new Indentation() { Left = "2880", FirstLine = "0" };

                previousParagraphProperties5.Append(indentation5);

                level5.Append(startNumberingValue5);
                level5.Append(numberingFormat5);
                level5.Append(levelText5);
                level5.Append(levelJustification5);
                level5.Append(previousParagraphProperties5);

                Level level6 = new Level() { LevelIndex = 5 };
                StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
                LevelText levelText6 = new LevelText() { Val = "(%6)" };
                LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
                Indentation indentation6 = new Indentation() { Left = "3600", FirstLine = "0" };

                previousParagraphProperties6.Append(indentation6);

                level6.Append(startNumberingValue6);
                level6.Append(numberingFormat6);
                level6.Append(levelText6);
                level6.Append(levelJustification6);
                level6.Append(previousParagraphProperties6);

                Level level7 = new Level() { LevelIndex = 6 };
                StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
                LevelText levelText7 = new LevelText() { Val = "(%7)" };
                LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
                Indentation indentation7 = new Indentation() { Left = "4320", FirstLine = "0" };

                previousParagraphProperties7.Append(indentation7);

                level7.Append(startNumberingValue7);
                level7.Append(numberingFormat7);
                level7.Append(levelText7);
                level7.Append(levelJustification7);
                level7.Append(previousParagraphProperties7);

                Level level8 = new Level() { LevelIndex = 7 };
                StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
                LevelText levelText8 = new LevelText() { Val = "(%8)" };
                LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
                Indentation indentation8 = new Indentation() { Left = "5040", FirstLine = "0" };

                previousParagraphProperties8.Append(indentation8);

                level8.Append(startNumberingValue8);
                level8.Append(numberingFormat8);
                level8.Append(levelText8);
                level8.Append(levelJustification8);
                level8.Append(previousParagraphProperties8);

                Level level9 = new Level() { LevelIndex = 8 };
                StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
                LevelText levelText9 = new LevelText() { Val = "(%9)" };
                LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
                Indentation indentation9 = new Indentation() { Left = "5760", FirstLine = "0" };

                previousParagraphProperties9.Append(indentation9);

                level9.Append(startNumberingValue9);
                level9.Append(numberingFormat9);
                level9.Append(levelText9);
                level9.Append(levelJustification9);
                level9.Append(previousParagraphProperties9);

                abstractNum1.Append(nsid1);
                abstractNum1.Append(multiLevelType1);
                abstractNum1.Append(templateCode1);
                abstractNum1.Append(level1);
                abstractNum1.Append(level2);
                abstractNum1.Append(level3);
                abstractNum1.Append(level4);
                abstractNum1.Append(level5);
                abstractNum1.Append(level6);
                abstractNum1.Append(level7);
                abstractNum1.Append(level8);
                abstractNum1.Append(level9);
                return abstractNum1;

            }
        }

        private static AbstractNum Chapters {
            get {
                AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 3 };
                abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
                Nsid nsid1 = new Nsid() { Val = GenerateNsidValue() };
                MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
                TemplateCode templateCode1 = new TemplateCode() { Val = "04090029" };

                Level level1 = new Level() { LevelIndex = 0 };
                StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelSuffix levelSuffix1 = new LevelSuffix() { Val = LevelSuffixValues.Space };
                LevelText levelText1 = new LevelText() { Val = "Chapter %1" };
                LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
                Indentation indentation1 = new Indentation() { Left = "0", FirstLine = "0" };

                previousParagraphProperties1.Append(indentation1);

                level1.Append(startNumberingValue1);
                level1.Append(numberingFormat1);
                level1.Append(levelSuffix1);
                level1.Append(levelText1);
                level1.Append(levelJustification1);
                level1.Append(previousParagraphProperties1);

                Level level2 = new Level() { LevelIndex = 1 };
                StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.None };
                LevelSuffix levelSuffix2 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
                LevelText levelText2 = new LevelText() { Val = "" };
                LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
                Indentation indentation2 = new Indentation() { Left = "0", FirstLine = "0" };

                previousParagraphProperties2.Append(indentation2);

                level2.Append(startNumberingValue2);
                level2.Append(numberingFormat2);
                level2.Append(levelSuffix2);
                level2.Append(levelText2);
                level2.Append(levelJustification2);
                level2.Append(previousParagraphProperties2);

                Level level3 = new Level() { LevelIndex = 2 };
                StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.None };
                LevelSuffix levelSuffix3 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
                LevelText levelText3 = new LevelText() { Val = "" };
                LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
                Indentation indentation3 = new Indentation() { Left = "0", FirstLine = "0" };

                previousParagraphProperties3.Append(indentation3);

                level3.Append(startNumberingValue3);
                level3.Append(numberingFormat3);
                level3.Append(levelSuffix3);
                level3.Append(levelText3);
                level3.Append(levelJustification3);
                level3.Append(previousParagraphProperties3);

                Level level4 = new Level() { LevelIndex = 3 };
                StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.None };
                LevelSuffix levelSuffix4 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
                LevelText levelText4 = new LevelText() { Val = "" };
                LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
                Indentation indentation4 = new Indentation() { Left = "0", FirstLine = "0" };

                previousParagraphProperties4.Append(indentation4);

                level4.Append(startNumberingValue4);
                level4.Append(numberingFormat4);
                level4.Append(levelSuffix4);
                level4.Append(levelText4);
                level4.Append(levelJustification4);
                level4.Append(previousParagraphProperties4);

                Level level5 = new Level() { LevelIndex = 4 };
                StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.None };
                LevelSuffix levelSuffix5 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
                LevelText levelText5 = new LevelText() { Val = "" };
                LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
                Indentation indentation5 = new Indentation() { Left = "0", FirstLine = "0" };

                previousParagraphProperties5.Append(indentation5);

                level5.Append(startNumberingValue5);
                level5.Append(numberingFormat5);
                level5.Append(levelSuffix5);
                level5.Append(levelText5);
                level5.Append(levelJustification5);
                level5.Append(previousParagraphProperties5);

                Level level6 = new Level() { LevelIndex = 5 };
                StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.None };
                LevelSuffix levelSuffix6 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
                LevelText levelText6 = new LevelText() { Val = "" };
                LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
                Indentation indentation6 = new Indentation() { Left = "0", FirstLine = "0" };

                previousParagraphProperties6.Append(indentation6);

                level6.Append(startNumberingValue6);
                level6.Append(numberingFormat6);
                level6.Append(levelSuffix6);
                level6.Append(levelText6);
                level6.Append(levelJustification6);
                level6.Append(previousParagraphProperties6);

                Level level7 = new Level() { LevelIndex = 6 };
                StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.None };
                LevelSuffix levelSuffix7 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
                LevelText levelText7 = new LevelText() { Val = "" };
                LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
                Indentation indentation7 = new Indentation() { Left = "0", FirstLine = "0" };

                previousParagraphProperties7.Append(indentation7);

                level7.Append(startNumberingValue7);
                level7.Append(numberingFormat7);
                level7.Append(levelSuffix7);
                level7.Append(levelText7);
                level7.Append(levelJustification7);
                level7.Append(previousParagraphProperties7);

                Level level8 = new Level() { LevelIndex = 7 };
                StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.None };
                LevelSuffix levelSuffix8 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
                LevelText levelText8 = new LevelText() { Val = "" };
                LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
                Indentation indentation8 = new Indentation() { Left = "0", FirstLine = "0" };

                previousParagraphProperties8.Append(indentation8);

                level8.Append(startNumberingValue8);
                level8.Append(numberingFormat8);
                level8.Append(levelSuffix8);
                level8.Append(levelText8);
                level8.Append(levelJustification8);
                level8.Append(previousParagraphProperties8);

                Level level9 = new Level() { LevelIndex = 8 };
                StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.None };
                LevelSuffix levelSuffix9 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
                LevelText levelText9 = new LevelText() { Val = "" };
                LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
                Indentation indentation9 = new Indentation() { Left = "0", FirstLine = "0" };

                previousParagraphProperties9.Append(indentation9);

                level9.Append(startNumberingValue9);
                level9.Append(numberingFormat9);
                level9.Append(levelSuffix9);
                level9.Append(levelText9);
                level9.Append(levelJustification9);
                level9.Append(previousParagraphProperties9);

                abstractNum1.Append(nsid1);
                abstractNum1.Append(multiLevelType1);
                abstractNum1.Append(templateCode1);
                abstractNum1.Append(level1);
                abstractNum1.Append(level2);
                abstractNum1.Append(level3);
                abstractNum1.Append(level4);
                abstractNum1.Append(level5);
                abstractNum1.Append(level6);
                abstractNum1.Append(level7);
                abstractNum1.Append(level8);
                abstractNum1.Append(level9);
                return abstractNum1;

            }
        }

        private static AbstractNum BulletedChars {
            get {
                AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 4 };
                abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
                Nsid nsid1 = new Nsid() { Val = GenerateNsidValue() };
                MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
                TemplateCode templateCode1 = new TemplateCode() { Val = "04090021" };

                Level level1 = new Level() { LevelIndex = 0 };
                StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText1 = new LevelText() { Val = "v" };
                LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
                Indentation indentation1 = new Indentation() { Left = "360", Hanging = "360" };

                previousParagraphProperties1.Append(indentation1);

                NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
                RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties1.Append(runFonts1);

                level1.Append(startNumberingValue1);
                level1.Append(numberingFormat1);
                level1.Append(levelText1);
                level1.Append(levelJustification1);
                level1.Append(previousParagraphProperties1);
                level1.Append(numberingSymbolRunProperties1);

                Level level2 = new Level() { LevelIndex = 1 };
                StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText2 = new LevelText() { Val = "Ø" };
                LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
                Indentation indentation2 = new Indentation() { Left = "720", Hanging = "360" };

                previousParagraphProperties2.Append(indentation2);

                NumberingSymbolRunProperties numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();
                RunFonts runFonts2 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties2.Append(runFonts2);

                level2.Append(startNumberingValue2);
                level2.Append(numberingFormat2);
                level2.Append(levelText2);
                level2.Append(levelJustification2);
                level2.Append(previousParagraphProperties2);
                level2.Append(numberingSymbolRunProperties2);

                Level level3 = new Level() { LevelIndex = 2 };
                StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText3 = new LevelText() { Val = "§" };
                LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
                Indentation indentation3 = new Indentation() { Left = "1080", Hanging = "360" };

                previousParagraphProperties3.Append(indentation3);

                NumberingSymbolRunProperties numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();
                RunFonts runFonts3 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties3.Append(runFonts3);

                level3.Append(startNumberingValue3);
                level3.Append(numberingFormat3);
                level3.Append(levelText3);
                level3.Append(levelJustification3);
                level3.Append(previousParagraphProperties3);
                level3.Append(numberingSymbolRunProperties3);

                Level level4 = new Level() { LevelIndex = 3 };
                StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText4 = new LevelText() { Val = "·" };
                LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
                Indentation indentation4 = new Indentation() { Left = "1440", Hanging = "360" };

                previousParagraphProperties4.Append(indentation4);

                NumberingSymbolRunProperties numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();
                RunFonts runFonts4 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

                numberingSymbolRunProperties4.Append(runFonts4);

                level4.Append(startNumberingValue4);
                level4.Append(numberingFormat4);
                level4.Append(levelText4);
                level4.Append(levelJustification4);
                level4.Append(previousParagraphProperties4);
                level4.Append(numberingSymbolRunProperties4);

                Level level5 = new Level() { LevelIndex = 4 };
                StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText5 = new LevelText() { Val = "¨" };
                LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
                Indentation indentation5 = new Indentation() { Left = "1800", Hanging = "360" };

                previousParagraphProperties5.Append(indentation5);

                NumberingSymbolRunProperties numberingSymbolRunProperties5 = new NumberingSymbolRunProperties();
                RunFonts runFonts5 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

                numberingSymbolRunProperties5.Append(runFonts5);

                level5.Append(startNumberingValue5);
                level5.Append(numberingFormat5);
                level5.Append(levelText5);
                level5.Append(levelJustification5);
                level5.Append(previousParagraphProperties5);
                level5.Append(numberingSymbolRunProperties5);

                Level level6 = new Level() { LevelIndex = 5 };
                StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText6 = new LevelText() { Val = "Ø" };
                LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
                Indentation indentation6 = new Indentation() { Left = "2160", Hanging = "360" };

                previousParagraphProperties6.Append(indentation6);

                NumberingSymbolRunProperties numberingSymbolRunProperties6 = new NumberingSymbolRunProperties();
                RunFonts runFonts6 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties6.Append(runFonts6);

                level6.Append(startNumberingValue6);
                level6.Append(numberingFormat6);
                level6.Append(levelText6);
                level6.Append(levelJustification6);
                level6.Append(previousParagraphProperties6);
                level6.Append(numberingSymbolRunProperties6);

                Level level7 = new Level() { LevelIndex = 6 };
                StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText7 = new LevelText() { Val = "§" };
                LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
                Indentation indentation7 = new Indentation() { Left = "2520", Hanging = "360" };

                previousParagraphProperties7.Append(indentation7);

                NumberingSymbolRunProperties numberingSymbolRunProperties7 = new NumberingSymbolRunProperties();
                RunFonts runFonts7 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties7.Append(runFonts7);

                level7.Append(startNumberingValue7);
                level7.Append(numberingFormat7);
                level7.Append(levelText7);
                level7.Append(levelJustification7);
                level7.Append(previousParagraphProperties7);
                level7.Append(numberingSymbolRunProperties7);

                Level level8 = new Level() { LevelIndex = 7 };
                StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText8 = new LevelText() { Val = "·" };
                LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
                Indentation indentation8 = new Indentation() { Left = "2880", Hanging = "360" };

                previousParagraphProperties8.Append(indentation8);

                NumberingSymbolRunProperties numberingSymbolRunProperties8 = new NumberingSymbolRunProperties();
                RunFonts runFonts8 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

                numberingSymbolRunProperties8.Append(runFonts8);

                level8.Append(startNumberingValue8);
                level8.Append(numberingFormat8);
                level8.Append(levelText8);
                level8.Append(levelJustification8);
                level8.Append(previousParagraphProperties8);
                level8.Append(numberingSymbolRunProperties8);

                Level level9 = new Level() { LevelIndex = 8 };
                StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText9 = new LevelText() { Val = "¨" };
                LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
                Indentation indentation9 = new Indentation() { Left = "3240", Hanging = "360" };

                previousParagraphProperties9.Append(indentation9);

                NumberingSymbolRunProperties numberingSymbolRunProperties9 = new NumberingSymbolRunProperties();
                RunFonts runFonts9 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

                numberingSymbolRunProperties9.Append(runFonts9);

                level9.Append(startNumberingValue9);
                level9.Append(numberingFormat9);
                level9.Append(levelText9);
                level9.Append(levelJustification9);
                level9.Append(previousParagraphProperties9);
                level9.Append(numberingSymbolRunProperties9);

                abstractNum1.Append(nsid1);
                abstractNum1.Append(multiLevelType1);
                abstractNum1.Append(templateCode1);
                abstractNum1.Append(level1);
                abstractNum1.Append(level2);
                abstractNum1.Append(level3);
                abstractNum1.Append(level4);
                abstractNum1.Append(level5);
                abstractNum1.Append(level6);
                abstractNum1.Append(level7);
                abstractNum1.Append(level8);
                abstractNum1.Append(level9);
                return abstractNum1;

            }
        }

        private static AbstractNum Heading1ai {
            get {
                AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 5 };
                abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
                Nsid nsid1 = new Nsid() { Val = GenerateNsidValue() };
                MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
                TemplateCode templateCode1 = new TemplateCode() { Val = "0409001D" };

                Level level1 = new Level() { LevelIndex = 0 };
                StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText1 = new LevelText() { Val = "%1)" };
                LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
                Indentation indentation1 = new Indentation() { Left = "360", Hanging = "360" };

                previousParagraphProperties1.Append(indentation1);

                level1.Append(startNumberingValue1);
                level1.Append(numberingFormat1);
                level1.Append(levelText1);
                level1.Append(levelJustification1);
                level1.Append(previousParagraphProperties1);

                Level level2 = new Level() { LevelIndex = 1 };
                StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
                LevelText levelText2 = new LevelText() { Val = "%2)" };
                LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
                Indentation indentation2 = new Indentation() { Left = "720", Hanging = "360" };

                previousParagraphProperties2.Append(indentation2);

                level2.Append(startNumberingValue2);
                level2.Append(numberingFormat2);
                level2.Append(levelText2);
                level2.Append(levelJustification2);
                level2.Append(previousParagraphProperties2);

                Level level3 = new Level() { LevelIndex = 2 };
                StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
                LevelText levelText3 = new LevelText() { Val = "%3)" };
                LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
                Indentation indentation3 = new Indentation() { Left = "1080", Hanging = "360" };

                previousParagraphProperties3.Append(indentation3);

                level3.Append(startNumberingValue3);
                level3.Append(numberingFormat3);
                level3.Append(levelText3);
                level3.Append(levelJustification3);
                level3.Append(previousParagraphProperties3);

                Level level4 = new Level() { LevelIndex = 3 };
                StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText4 = new LevelText() { Val = "(%4)" };
                LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
                Indentation indentation4 = new Indentation() { Left = "1440", Hanging = "360" };

                previousParagraphProperties4.Append(indentation4);

                level4.Append(startNumberingValue4);
                level4.Append(numberingFormat4);
                level4.Append(levelText4);
                level4.Append(levelJustification4);
                level4.Append(previousParagraphProperties4);

                Level level5 = new Level() { LevelIndex = 4 };
                StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
                LevelText levelText5 = new LevelText() { Val = "(%5)" };
                LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
                Indentation indentation5 = new Indentation() { Left = "1800", Hanging = "360" };

                previousParagraphProperties5.Append(indentation5);

                level5.Append(startNumberingValue5);
                level5.Append(numberingFormat5);
                level5.Append(levelText5);
                level5.Append(levelJustification5);
                level5.Append(previousParagraphProperties5);

                Level level6 = new Level() { LevelIndex = 5 };
                StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
                LevelText levelText6 = new LevelText() { Val = "(%6)" };
                LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
                Indentation indentation6 = new Indentation() { Left = "2160", Hanging = "360" };

                previousParagraphProperties6.Append(indentation6);

                level6.Append(startNumberingValue6);
                level6.Append(numberingFormat6);
                level6.Append(levelText6);
                level6.Append(levelJustification6);
                level6.Append(previousParagraphProperties6);

                Level level7 = new Level() { LevelIndex = 6 };
                StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText7 = new LevelText() { Val = "%7." };
                LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
                Indentation indentation7 = new Indentation() { Left = "2520", Hanging = "360" };

                previousParagraphProperties7.Append(indentation7);

                level7.Append(startNumberingValue7);
                level7.Append(numberingFormat7);
                level7.Append(levelText7);
                level7.Append(levelJustification7);
                level7.Append(previousParagraphProperties7);

                Level level8 = new Level() { LevelIndex = 7 };
                StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.LowerLetter };
                LevelText levelText8 = new LevelText() { Val = "%8." };
                LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
                Indentation indentation8 = new Indentation() { Left = "2880", Hanging = "360" };

                previousParagraphProperties8.Append(indentation8);

                level8.Append(startNumberingValue8);
                level8.Append(numberingFormat8);
                level8.Append(levelText8);
                level8.Append(levelJustification8);
                level8.Append(previousParagraphProperties8);

                Level level9 = new Level() { LevelIndex = 8 };
                StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
                LevelText levelText9 = new LevelText() { Val = "%9." };
                LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
                Indentation indentation9 = new Indentation() { Left = "3240", Hanging = "360" };

                previousParagraphProperties9.Append(indentation9);

                level9.Append(startNumberingValue9);
                level9.Append(numberingFormat9);
                level9.Append(levelText9);
                level9.Append(levelJustification9);
                level9.Append(previousParagraphProperties9);

                abstractNum1.Append(nsid1);
                abstractNum1.Append(multiLevelType1);
                abstractNum1.Append(templateCode1);
                abstractNum1.Append(level1);
                abstractNum1.Append(level2);
                abstractNum1.Append(level3);
                abstractNum1.Append(level4);
                abstractNum1.Append(level5);
                abstractNum1.Append(level6);
                abstractNum1.Append(level7);
                abstractNum1.Append(level8);
                abstractNum1.Append(level9);
                return abstractNum1;
            }
        }

        private static AbstractNum Headings111Shifted {
            get {
                AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 6 };
                abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
                Nsid nsid1 = new Nsid() { Val = GenerateNsidValue() };
                MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.Multilevel };
                TemplateCode templateCode1 = new TemplateCode() { Val = "0409001F" };

                Level level1 = new Level() { LevelIndex = 0 };
                StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText1 = new LevelText() { Val = "%1." };
                LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
                Indentation indentation1 = new Indentation() { Left = "360", Hanging = "360" };

                previousParagraphProperties1.Append(indentation1);

                level1.Append(startNumberingValue1);
                level1.Append(numberingFormat1);
                level1.Append(levelText1);
                level1.Append(levelJustification1);
                level1.Append(previousParagraphProperties1);

                Level level2 = new Level() { LevelIndex = 1 };
                StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText2 = new LevelText() { Val = "%1.%2." };
                LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
                Indentation indentation2 = new Indentation() { Left = "792", Hanging = "432" };

                previousParagraphProperties2.Append(indentation2);

                level2.Append(startNumberingValue2);
                level2.Append(numberingFormat2);
                level2.Append(levelText2);
                level2.Append(levelJustification2);
                level2.Append(previousParagraphProperties2);

                Level level3 = new Level() { LevelIndex = 2 };
                StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText3 = new LevelText() { Val = "%1.%2.%3." };
                LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
                Indentation indentation3 = new Indentation() { Left = "1224", Hanging = "504" };

                previousParagraphProperties3.Append(indentation3);

                level3.Append(startNumberingValue3);
                level3.Append(numberingFormat3);
                level3.Append(levelText3);
                level3.Append(levelJustification3);
                level3.Append(previousParagraphProperties3);

                Level level4 = new Level() { LevelIndex = 3 };
                StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText4 = new LevelText() { Val = "%1.%2.%3.%4." };
                LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
                Indentation indentation4 = new Indentation() { Left = "1728", Hanging = "648" };

                previousParagraphProperties4.Append(indentation4);

                level4.Append(startNumberingValue4);
                level4.Append(numberingFormat4);
                level4.Append(levelText4);
                level4.Append(levelJustification4);
                level4.Append(previousParagraphProperties4);

                Level level5 = new Level() { LevelIndex = 4 };
                StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText5 = new LevelText() { Val = "%1.%2.%3.%4.%5." };
                LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
                Indentation indentation5 = new Indentation() { Left = "2232", Hanging = "792" };

                previousParagraphProperties5.Append(indentation5);

                level5.Append(startNumberingValue5);
                level5.Append(numberingFormat5);
                level5.Append(levelText5);
                level5.Append(levelJustification5);
                level5.Append(previousParagraphProperties5);

                Level level6 = new Level() { LevelIndex = 5 };
                StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText6 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6." };
                LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
                Indentation indentation6 = new Indentation() { Left = "2736", Hanging = "936" };

                previousParagraphProperties6.Append(indentation6);

                level6.Append(startNumberingValue6);
                level6.Append(numberingFormat6);
                level6.Append(levelText6);
                level6.Append(levelJustification6);
                level6.Append(previousParagraphProperties6);

                Level level7 = new Level() { LevelIndex = 6 };
                StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText7 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7." };
                LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
                Indentation indentation7 = new Indentation() { Left = "3240", Hanging = "1080" };

                previousParagraphProperties7.Append(indentation7);

                level7.Append(startNumberingValue7);
                level7.Append(numberingFormat7);
                level7.Append(levelText7);
                level7.Append(levelJustification7);
                level7.Append(previousParagraphProperties7);

                Level level8 = new Level() { LevelIndex = 7 };
                StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText8 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8." };
                LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
                Indentation indentation8 = new Indentation() { Left = "3744", Hanging = "1224" };

                previousParagraphProperties8.Append(indentation8);

                level8.Append(startNumberingValue8);
                level8.Append(numberingFormat8);
                level8.Append(levelText8);
                level8.Append(levelJustification8);
                level8.Append(previousParagraphProperties8);

                Level level9 = new Level() { LevelIndex = 8 };
                StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
                LevelText levelText9 = new LevelText() { Val = "%1.%2.%3.%4.%5.%6.%7.%8.%9." };
                LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
                Indentation indentation9 = new Indentation() { Left = "4320", Hanging = "1440" };

                previousParagraphProperties9.Append(indentation9);

                level9.Append(startNumberingValue9);
                level9.Append(numberingFormat9);
                level9.Append(levelText9);
                level9.Append(levelJustification9);
                level9.Append(previousParagraphProperties9);

                abstractNum1.Append(nsid1);
                abstractNum1.Append(multiLevelType1);
                abstractNum1.Append(templateCode1);
                abstractNum1.Append(level1);
                abstractNum1.Append(level2);
                abstractNum1.Append(level3);
                abstractNum1.Append(level4);
                abstractNum1.Append(level5);
                abstractNum1.Append(level6);
                abstractNum1.Append(level7);
                abstractNum1.Append(level8);
                abstractNum1.Append(level9);
                return abstractNum1;
            }
        }

        private static AbstractNum Bulleted {
            get {
                AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 7 };
                abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
                Nsid nsid1 = new Nsid() { Val = GenerateNsidValue() };
                MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
                TemplateCode templateCode1 = new TemplateCode() { Val = "934E79A6" };

                Level level1 = new Level() { LevelIndex = 0, TemplateCode = "04090001" };
                StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText1 = new LevelText() { Val = "·" };
                LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
                Indentation indentation1 = new Indentation() { Left = "720", Hanging = "360" };

                previousParagraphProperties1.Append(indentation1);

                NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
                RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

                numberingSymbolRunProperties1.Append(runFonts1);

                level1.Append(startNumberingValue1);
                level1.Append(numberingFormat1);
                level1.Append(levelText1);
                level1.Append(levelJustification1);
                level1.Append(previousParagraphProperties1);
                level1.Append(numberingSymbolRunProperties1);

                Level level2 = new Level() { LevelIndex = 1, TemplateCode = "04090003" };
                StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText2 = new LevelText() { Val = "o" };
                LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
                Indentation indentation2 = new Indentation() { Left = "1440", Hanging = "360" };

                previousParagraphProperties2.Append(indentation2);

                NumberingSymbolRunProperties numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();
                RunFonts runFonts2 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" };

                numberingSymbolRunProperties2.Append(runFonts2);

                level2.Append(startNumberingValue2);
                level2.Append(numberingFormat2);
                level2.Append(levelText2);
                level2.Append(levelJustification2);
                level2.Append(previousParagraphProperties2);
                level2.Append(numberingSymbolRunProperties2);

                Level level3 = new Level() { LevelIndex = 2, TemplateCode = "04090005", Tentative = true };
                StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText3 = new LevelText() { Val = "§" };
                LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
                Indentation indentation3 = new Indentation() { Left = "2160", Hanging = "360" };

                previousParagraphProperties3.Append(indentation3);

                NumberingSymbolRunProperties numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();
                RunFonts runFonts3 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties3.Append(runFonts3);

                level3.Append(startNumberingValue3);
                level3.Append(numberingFormat3);
                level3.Append(levelText3);
                level3.Append(levelJustification3);
                level3.Append(previousParagraphProperties3);
                level3.Append(numberingSymbolRunProperties3);

                Level level4 = new Level() { LevelIndex = 3, TemplateCode = "04090001", Tentative = true };
                StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText4 = new LevelText() { Val = "·" };
                LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
                Indentation indentation4 = new Indentation() { Left = "2880", Hanging = "360" };

                previousParagraphProperties4.Append(indentation4);

                NumberingSymbolRunProperties numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();
                RunFonts runFonts4 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

                numberingSymbolRunProperties4.Append(runFonts4);

                level4.Append(startNumberingValue4);
                level4.Append(numberingFormat4);
                level4.Append(levelText4);
                level4.Append(levelJustification4);
                level4.Append(previousParagraphProperties4);
                level4.Append(numberingSymbolRunProperties4);

                Level level5 = new Level() { LevelIndex = 4, TemplateCode = "04090003", Tentative = true };
                StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText5 = new LevelText() { Val = "o" };
                LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
                Indentation indentation5 = new Indentation() { Left = "3600", Hanging = "360" };

                previousParagraphProperties5.Append(indentation5);

                NumberingSymbolRunProperties numberingSymbolRunProperties5 = new NumberingSymbolRunProperties();
                RunFonts runFonts5 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" };

                numberingSymbolRunProperties5.Append(runFonts5);

                level5.Append(startNumberingValue5);
                level5.Append(numberingFormat5);
                level5.Append(levelText5);
                level5.Append(levelJustification5);
                level5.Append(previousParagraphProperties5);
                level5.Append(numberingSymbolRunProperties5);

                Level level6 = new Level() { LevelIndex = 5, TemplateCode = "04090005", Tentative = true };
                StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText6 = new LevelText() { Val = "§" };
                LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
                Indentation indentation6 = new Indentation() { Left = "4320", Hanging = "360" };

                previousParagraphProperties6.Append(indentation6);

                NumberingSymbolRunProperties numberingSymbolRunProperties6 = new NumberingSymbolRunProperties();
                RunFonts runFonts6 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties6.Append(runFonts6);

                level6.Append(startNumberingValue6);
                level6.Append(numberingFormat6);
                level6.Append(levelText6);
                level6.Append(levelJustification6);
                level6.Append(previousParagraphProperties6);
                level6.Append(numberingSymbolRunProperties6);

                Level level7 = new Level() { LevelIndex = 6, TemplateCode = "04090001", Tentative = true };
                StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText7 = new LevelText() { Val = "·" };
                LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
                Indentation indentation7 = new Indentation() { Left = "5040", Hanging = "360" };

                previousParagraphProperties7.Append(indentation7);

                NumberingSymbolRunProperties numberingSymbolRunProperties7 = new NumberingSymbolRunProperties();
                RunFonts runFonts7 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" };

                numberingSymbolRunProperties7.Append(runFonts7);

                level7.Append(startNumberingValue7);
                level7.Append(numberingFormat7);
                level7.Append(levelText7);
                level7.Append(levelJustification7);
                level7.Append(previousParagraphProperties7);
                level7.Append(numberingSymbolRunProperties7);

                Level level8 = new Level() { LevelIndex = 7, TemplateCode = "04090003", Tentative = true };
                StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText8 = new LevelText() { Val = "o" };
                LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
                Indentation indentation8 = new Indentation() { Left = "5760", Hanging = "360" };

                previousParagraphProperties8.Append(indentation8);

                NumberingSymbolRunProperties numberingSymbolRunProperties8 = new NumberingSymbolRunProperties();
                RunFonts runFonts8 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" };

                numberingSymbolRunProperties8.Append(runFonts8);

                level8.Append(startNumberingValue8);
                level8.Append(numberingFormat8);
                level8.Append(levelText8);
                level8.Append(levelJustification8);
                level8.Append(previousParagraphProperties8);
                level8.Append(numberingSymbolRunProperties8);

                Level level9 = new Level() { LevelIndex = 8, TemplateCode = "04090005", Tentative = true };
                StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText9 = new LevelText() { Val = "§" };
                LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
                Indentation indentation9 = new Indentation() { Left = "6480", Hanging = "360" };

                previousParagraphProperties9.Append(indentation9);

                NumberingSymbolRunProperties numberingSymbolRunProperties9 = new NumberingSymbolRunProperties();
                RunFonts runFonts9 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties9.Append(runFonts9);

                level9.Append(startNumberingValue9);
                level9.Append(numberingFormat9);
                level9.Append(levelText9);
                level9.Append(levelJustification9);
                level9.Append(previousParagraphProperties9);
                level9.Append(numberingSymbolRunProperties9);

                abstractNum1.Append(nsid1);
                abstractNum1.Append(multiLevelType1);
                abstractNum1.Append(templateCode1);
                abstractNum1.Append(level1);
                abstractNum1.Append(level2);
                abstractNum1.Append(level3);
                abstractNum1.Append(level4);
                abstractNum1.Append(level5);
                abstractNum1.Append(level6);
                abstractNum1.Append(level7);
                abstractNum1.Append(level8);
                abstractNum1.Append(level9);
                return abstractNum1;
            }
        }

    }
}
