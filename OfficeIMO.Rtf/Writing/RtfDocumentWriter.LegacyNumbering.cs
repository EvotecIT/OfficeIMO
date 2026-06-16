namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteLegacyNumbering(StringBuilder builder, RtfLegacyNumbering numbering, int unicodeSkipCount) {
        if (numbering == null || !numbering.HasAnyValue) {
            return;
        }

        builder.Append(@"{\*\pn");
        WriteLegacyNumberingLevel(builder, numbering);
        WriteLegacyNumberingStyle(builder, numbering.NumberStyle);
        WriteLegacyNumberingCharacterFormat(builder, numbering);
        WriteLegacyNumberingText(builder, "pntxtb", numbering.TextBefore, unicodeSkipCount);
        WriteLegacyNumberingText(builder, "pntxta", numbering.TextAfter, unicodeSkipCount);
        AppendOptionalBinary(builder, @"\pnnumonce", numbering.NumberEachCellOnce);
        AppendOptionalBinary(builder, @"\pnacross", numbering.NumberAcrossRows);
        AppendOptionalTwips(builder, @"\pnindent", numbering.IndentTwips);
        AppendOptionalTwips(builder, @"\pnsp", numbering.SpaceTwips);
        AppendOptionalBinary(builder, @"\pnprev", numbering.IncludePreviousLevels);
        WriteLegacyNumberingAlignment(builder, numbering.Alignment);
        AppendOptionalTwips(builder, @"\pnstart", numbering.StartAt);
        AppendOptionalBinary(builder, @"\pnhang", numbering.HangingIndent);
        AppendOptionalBinary(builder, @"\pnrestart", numbering.RestartAfterSection);
        builder.Append('}');
    }

    private static void WriteLegacyNumberingLevel(StringBuilder builder, RtfLegacyNumbering numbering) {
        switch (numbering.LevelKind) {
            case RtfLegacyNumberingLevelKind.Level:
                builder.Append(@"\pnlvl");
                if (numbering.Level.HasValue) {
                    builder.Append(numbering.Level.Value.ToString(CultureInfo.InvariantCulture));
                }

                break;
            case RtfLegacyNumberingLevelKind.Bullet:
                builder.Append(@"\pnlvlblt");
                break;
            case RtfLegacyNumberingLevelKind.Body:
                builder.Append(@"\pnlvlbody");
                break;
            case RtfLegacyNumberingLevelKind.Continue:
                builder.Append(@"\pnlvlcont");
                break;
        }
    }

    private static void WriteLegacyNumberingStyle(StringBuilder builder, RtfLegacyNumberingStyle style) {
        string? control = style switch {
            RtfLegacyNumberingStyle.Cardinal => @"\pncard",
            RtfLegacyNumberingStyle.Decimal => @"\pndec",
            RtfLegacyNumberingStyle.UpperLetter => @"\pnucltr",
            RtfLegacyNumberingStyle.UpperRoman => @"\pnucrm",
            RtfLegacyNumberingStyle.LowerLetter => @"\pnlcltr",
            RtfLegacyNumberingStyle.LowerRoman => @"\pnlcrm",
            RtfLegacyNumberingStyle.Ordinal => @"\pnord",
            RtfLegacyNumberingStyle.OrdinalText => @"\pnordt",
            _ => null
        };

        if (control != null) {
            builder.Append(control);
        }
    }

    private static void WriteLegacyNumberingCharacterFormat(StringBuilder builder, RtfLegacyNumbering numbering) {
        AppendOptionalTwips(builder, @"\pnf", numbering.FontId);
        AppendOptionalTwips(builder, @"\pnfs", numbering.FontSizeHalfPoints);
        AppendOptionalBinary(builder, @"\pnb", numbering.Bold);
        AppendOptionalBinary(builder, @"\pni", numbering.Italic);
        AppendOptionalBinary(builder, @"\pncaps", numbering.AllCaps);
        AppendOptionalBinary(builder, @"\pnscaps", numbering.SmallCaps);
        WriteLegacyNumberingUnderline(builder, numbering.UnderlineStyle);
        AppendOptionalBinary(builder, @"\pnstrike", numbering.Strike);
        AppendOptionalTwips(builder, @"\pncf", numbering.ForegroundColorIndex);
    }

    private static void WriteLegacyNumberingUnderline(StringBuilder builder, RtfUnderlineStyle? underlineStyle) {
        if (!underlineStyle.HasValue) {
            return;
        }

        builder.Append(underlineStyle.Value switch {
            RtfUnderlineStyle.Dotted => @"\pnuld",
            RtfUnderlineStyle.Double => @"\pnuldb",
            RtfUnderlineStyle.Words => @"\pnulw",
            RtfUnderlineStyle.None => @"\pnulnone",
            _ => @"\pnul"
        });
    }

    private static void WriteLegacyNumberingText(StringBuilder builder, string destination, string? text, int unicodeSkipCount) {
        if (text == null) {
            return;
        }

        builder.Append(@"{\");
        builder.Append(destination);
        builder.Append(' ');
        builder.Append(EscapeText(text, unicodeSkipCount));
        builder.Append('}');
    }

    private static void WriteLegacyNumberingAlignment(StringBuilder builder, RtfLegacyNumberingAlignment? alignment) {
        if (!alignment.HasValue) {
            return;
        }

        builder.Append(alignment.Value switch {
            RtfLegacyNumberingAlignment.Center => @"\pnqc",
            RtfLegacyNumberingAlignment.Right => @"\pnqr",
            _ => @"\pnql"
        });
    }
}
