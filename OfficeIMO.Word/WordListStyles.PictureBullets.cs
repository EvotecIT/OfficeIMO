using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word;

/// <summary>
/// Defines built-in list styles.
/// </summary>
public static partial class WordListStyles {
    internal static AbstractNum CreatePictureBulletStyle(int pictureBulletId) {
        AbstractNum abstractNum1 = CreateNewAbstractNum();
        abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
        Nsid nsid1 = new Nsid() { Val = GenerateNsidValue() };
        MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
        TemplateCode templateCode1 = new TemplateCode() { Val = GenerateNsidValue() };

        Level level1 = new Level() { LevelIndex = 0 };
        StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
        LevelPictureBulletId levelPictureBulletId1 = new LevelPictureBulletId() { Val = pictureBulletId };
        LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
        Indentation indentation1 = new Indentation() { Left = "720", Hanging = "360" };
        previousParagraphProperties1.Append(indentation1);

        level1.Append(startNumberingValue1);
        level1.Append(numberingFormat1);
        level1.Append(levelPictureBulletId1);
        level1.Append(levelJustification1);
        level1.Append(previousParagraphProperties1);

        abstractNum1.Append(nsid1);
        abstractNum1.Append(multiLevelType1);
        abstractNum1.Append(templateCode1);
        abstractNum1.Append(level1);
        return abstractNum1;
    }
}
