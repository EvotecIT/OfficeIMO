using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;

internal static class SchemaValidation {
    internal static int Run(string inputPath, string schemaPath) {
        var secure = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 64 * 1024 * 1024 };
        using var inputReader = XmlReader.Create(inputPath, secure);
        var input = XDocument.Load(inputReader);
        using var schemaReader = XmlReader.Create(schemaPath, secure);
        var schemaXml = XDocument.Load(schemaReader);
        const string published = "http://schemas.microsoft.com/project/2007";
        const string application = "http://schemas.microsoft.com/project";
        bool alias = (string?)schemaXml.Root?.Attribute("targetNamespace") == published && input.Root?.Name.NamespaceName == application;
        if (alias) {
            foreach (var attribute in schemaXml.Descendants().Attributes().Where(a => a.Value == published)) attribute.Value = application;
        }
        var schemas = new XmlSchemaSet { XmlResolver = null };
        using var reader = schemaXml.CreateReader();
        schemas.Add(null, reader);
        schemas.Compile();
        var errors = new List<string>();
        input.Validate(schemas, (_, e) => errors.Add(e.Message));
        Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(new {
            file = Path.GetFileName(inputPath), schema = Path.GetFileName(schemaPath), namespaceAliasApplied = alias, errors
        }));
        return errors.Count == 0 ? 0 : 1;
    }
}
