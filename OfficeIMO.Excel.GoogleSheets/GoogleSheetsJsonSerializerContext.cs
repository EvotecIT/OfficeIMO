using System.Text.Json.Serialization;

namespace OfficeIMO.Excel.GoogleSheets {
    [JsonSourceGenerationOptions(
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNameCaseInsensitive = true)]
    [JsonSerializable(typeof(GoogleSheetsApiCreateSpreadsheetPayload))]
    [JsonSerializable(typeof(GoogleSheetsApiBatchUpdatePayload))]
    [JsonSerializable(typeof(GoogleSheetsApiBatchUpdateValuesPayload))]
    [JsonSerializable(typeof(GoogleSheetsApiCreateSpreadsheetResponse))]
    [JsonSerializable(typeof(GoogleSheetsApiSpreadsheetMetadataResponse))]
    [JsonSerializable(typeof(GoogleSheetsNativeSpreadsheet))]
    [JsonSerializable(typeof(object))]
    internal sealed partial class GoogleSheetsJsonSerializerContext : JsonSerializerContext {
    }
}
