namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteObject(StringBuilder builder, RtfObject rtfObject, int? defaultLanguageId, int unicodeSkipCount) {
        builder.Append(@"{\object");
        builder.Append(rtfObject.Kind switch {
            RtfObjectKind.Embedded => @"\objemb",
            RtfObjectKind.Linked => @"\objlink",
            RtfObjectKind.AutoLinked => @"\objautlink",
            RtfObjectKind.Subscription => @"\objsub",
            RtfObjectKind.Publisher => @"\objpub",
            RtfObjectKind.IconEmbedded => @"\objicemb",
            _ => string.Empty
        });
        AppendOptionalTwips(builder, @"\objw", rtfObject.Width);
        AppendOptionalTwips(builder, @"\objh", rtfObject.Height);
        AppendOptionalTwips(builder, @"\objscalex", rtfObject.ScaleX);
        AppendOptionalTwips(builder, @"\objscaley", rtfObject.ScaleY);
        WriteObjectTextDestination(builder, "objclass", rtfObject.ClassName, unicodeSkipCount);
        WriteObjectTextDestination(builder, "objname", rtfObject.Name, unicodeSkipCount);
        if (rtfObject.Data.Length > 0) {
            builder.Append(@"{\*\objdata ");
            WriteHexBytes(builder, rtfObject.Data);
            builder.Append('}');
        }

        WriteObjectResult(builder, rtfObject, defaultLanguageId, unicodeSkipCount);
        builder.Append('}');
    }

    private static void WriteObjectTextDestination(StringBuilder builder, string destination, string? value, int unicodeSkipCount) {
        if (string.IsNullOrEmpty(value)) return;
        builder.Append(@"{\*\");
        builder.Append(destination);
        builder.Append(' ');
        builder.Append(EscapeText(value!, unicodeSkipCount));
        builder.Append('}');
    }

    private static void WriteObjectResult(StringBuilder builder, RtfObject rtfObject, int? defaultLanguageId, int unicodeSkipCount) {
        if (rtfObject.ResultImage == null && rtfObject.Result.Inlines.Count == 0) return;

        builder.Append(@"{\result ");
        if (rtfObject.ResultImage != null) {
            WriteImage(builder, rtfObject.ResultImage);
        } else {
            var state = new RunWriteState(defaultLanguageId);
            foreach (IRtfInline inline in rtfObject.Result.Inlines) {
                WriteInline(builder, inline, state, defaultLanguageId, unicodeSkipCount);
            }

            ResetRunState(builder, state);
        }

        builder.Append('}');
    }
}
