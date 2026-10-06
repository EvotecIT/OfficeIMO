using System.Text.Json.Serialization;

namespace OfficeIMO.Studio.Infrastructure.Preferences;

[JsonSerializable(typeof(StudioSessionSnapshot))]
internal sealed partial class StudioSessionJsonContext : JsonSerializerContext;
