using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Defines built-in list styles.
/// </summary>
public static partial class WordListStyles {
    // API methods moved to WordListStyles.Api.cs



    private static AbstractNum Custom {
        get {
            AbstractNum abstractNum1 = CreateNewAbstractNum();
            abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid = new Nsid() { Val = GenerateNsidValue() };
            MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode1 = new TemplateCode() { Val = GenerateNsidValue() };

            abstractNum1.Append(nsid);
            abstractNum1.Append(multiLevelType1);
            abstractNum1.Append(templateCode1);

            return abstractNum1;
        }
    }
}
