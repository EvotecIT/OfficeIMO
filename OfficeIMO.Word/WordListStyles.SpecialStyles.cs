using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public static partial class WordListStyles {
    private static AbstractNum Chapters {
        get {
            AbstractNum abstractNum1 = CreateNewAbstractNum();
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

}

