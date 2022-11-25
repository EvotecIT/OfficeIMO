using System;
using System.IO;
using System.Xml;
using VerifyTests;
using VerifyXunit;

namespace OfficeIMO.VerifyTests;

[UsesVerify]
public abstract class VerifyTestBase {

    private static XmlWriterSettings _xmlWriterSettings = new() {
        Indent = true,
        NewLineOnAttributes = false,
        IndentChars = "  ",
        ConformanceLevel = ConformanceLevel.Document
    };

    static VerifyTestBase() {
        // To disable Visual Studio popping up on every test execution.
        Environment.SetEnvironmentVariable("DiffEngine_Disabled", "true");
        Environment.SetEnvironmentVariable("Verify_DisableClipboard", "true");
    }

    protected static VerifySettings GetSettings() {
        var settings = new VerifySettings();
        settings.UseDirectory("verified");
        return settings;
    }

    protected static string FormatXml(string value) {
        using var textReader = new StringReader(value);
        using var xmlReader = XmlReader.Create(
            textReader, new XmlReaderSettings { ConformanceLevel = _xmlWriterSettings.ConformanceLevel } );
        using var textWriter = new StringWriter();
        using (var xmlWriter = XmlWriter.Create(textWriter, _xmlWriterSettings))
            xmlWriter.WriteNode(xmlReader, true);
        return textWriter.ToString();
    }
}
