namespace OfficeIMO.Data.Conversions;

/// <summary>Event-stream helpers shared by both table passes.</summary>
internal static class TableEventWalker {
    /// <summary>Reads the next event that is not a comment; tables have no comments, so their omission is reported once.</summary>
    public static DataTokenType Next(DataEventReader reader) {
        while (reader.Read()) {
            if (reader.TokenType != DataTokenType.Comment) return reader.TokenType;
            reader.Context.Report(OfficeDiagnosticCodes.CommentsOmitted, OfficeConversionLossKind.Omission,
                "Comments were dropped because tables have no comments.");
        }

        return DataTokenType.None;
    }

    /// <summary>Skips the value whose first event is current.</summary>
    public static void SkipValue(DataEventReader reader) {
        if (reader.TokenType is not (DataTokenType.StartObject or DataTokenType.StartArray)) return;
        int depth = 1;
        while (depth > 0) {
            switch (Next(reader)) {
                case DataTokenType.StartObject:
                case DataTokenType.StartArray:
                    depth++;
                    break;
                case DataTokenType.EndObject:
                case DataTokenType.EndArray:
                    depth--;
                    break;
                case DataTokenType.None:
                    return;
            }
        }
    }

    /// <summary>Moves from the current value to the value a path selects; false when the path is absent.</summary>
    public static bool Navigate(DataEventReader reader, IReadOnlyList<DataPathSegment>? path) {
        if (path == null) return true;
        foreach (DataPathSegment segment in path) {
            if (segment.IsIndex) {
                if (reader.TokenType != DataTokenType.StartArray) return false;
                int index = 0;
                bool found = false;
                while (Next(reader) is not (DataTokenType.EndArray or DataTokenType.None)) {
                    if (index++ == segment.Index) {
                        found = true;
                        break;
                    }

                    SkipValue(reader);
                }

                if (!found) return false;
            } else {
                if (reader.TokenType != DataTokenType.StartObject) return false;
                bool found = false;
                while (Next(reader) == DataTokenType.PropertyName) {
                    bool match = reader.Text.SequenceEqual(segment.Name.AsSpan());
                    Next(reader);
                    if (match) {
                        found = true;
                        break;
                    }

                    SkipValue(reader);
                }

                if (!found) return false;
            }
        }

        return true;
    }

    public static void SkipToEndOfDocument(DataEventReader reader) {
        DataTokenType type;
        do {
            type = Next(reader);
        } while (type is not (DataTokenType.EndDocument or DataTokenType.None));
    }
}
