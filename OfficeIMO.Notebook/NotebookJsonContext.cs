using System.Text.Json.Serialization;

namespace OfficeIMO.Notebook;

[JsonSerializable(typeof(string))]
internal sealed partial class NotebookJsonContext : JsonSerializerContext { }
