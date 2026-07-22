namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    public static string Write(RtfDocument document, RtfWriteOptions options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        options ??= new RtfWriteOptions();
        RtfTableTraversalGuard.ValidateDocument(document);
        int unicodeSkipCount = GetUnicodeSkipCount(document.Settings);

        var builder = new StringBuilder();
        builder.Append(@"{\rtf1");
        WriteDocumentCharacterSet(builder, document.Settings);
        builder.Append(@"\deff");
        builder.Append((document.Settings.DefaultFontId ?? 0).ToString(CultureInfo.InvariantCulture));
        WritePageSetup(builder, document.PageSetup, isSection: false);
        WriteNoteSettings(builder, document.NoteSettings);
        WriteDocumentSettings(builder, document.Settings);
        WriteHtmlEncapsulation(builder, document, options, unicodeSkipCount);
        if (options.IncludeGenerator) {
            builder.Append(@"{\*\generator ");
            builder.Append(EscapeText(string.IsNullOrWhiteSpace(document.Info.Generator) ? "OfficeIMO.Rtf" : document.Info.Generator!.Trim(), unicodeSkipCount));
            builder.Append(";}");
        }

        WriteFontTable(builder, document, options, unicodeSkipCount);
        WriteFileTable(builder, document, unicodeSkipCount);
        WriteXmlNamespaceTable(builder, document, unicodeSkipCount);
        WriteColorTable(builder, document);
        WriteStyleSheet(builder, document, unicodeSkipCount);
        WriteListTables(builder, document, unicodeSkipCount);
        WriteRevisionTable(builder, document, unicodeSkipCount);
        WriteRevisionSaveIdTable(builder, document);
        WriteInfo(builder, document, unicodeSkipCount);
        WriteUserProperties(builder, document, unicodeSkipCount);
        WriteDocumentVariables(builder, document, unicodeSkipCount);
        WriteHeaderFooters(builder, document, unicodeSkipCount);
        HashSet<RtfNote> referencedNotes = RtfNoteReferenceCollector.Collect(document);
        WriteDetachedNotes(builder, document, referencedNotes, document.Settings.DefaultLanguageId, unicodeSkipCount);
        builder.AppendLine();

        if (document.Sections.Count > 0) {
            foreach (RtfSection section in document.Sections) {
                WriteSection(builder, section, document.Settings.DefaultLanguageId, unicodeSkipCount);
            }
        } else {
            foreach (IRtfBlock block in document.Blocks) {
                WriteBlock(builder, block, document.Settings.DefaultLanguageId, unicodeSkipCount);
            }
        }

        builder.Append('}');
        return builder.ToString();
    }

    private static int GetUnicodeSkipCount(RtfDocumentSettings settings) => settings.UnicodeSkipCount ?? 1;

    private static void WriteHtmlEncapsulation(StringBuilder builder, RtfDocument document, RtfWriteOptions options, int unicodeSkipCount) {
        RtfHtmlEncapsulation? encapsulation = document.HtmlEncapsulation;
        if (!options.IncludeHtmlEncapsulation || encapsulation == null || string.IsNullOrEmpty(encapsulation.Html)) return;

        builder.Append(@"\fromhtml");
        builder.Append(encapsulation.Version.ToString(CultureInfo.InvariantCulture));
        builder.Append(@"{\*\htmltag ");
        builder.Append(EscapeText(encapsulation.Html, unicodeSkipCount));
        builder.Append('}');
    }

    private static void WriteDocumentCharacterSet(StringBuilder builder, RtfDocumentSettings settings) {
        builder.Append(settings.CharacterSet switch {
            RtfDocumentCharacterSet.Mac => @"\mac",
            RtfDocumentCharacterSet.Pc => @"\pc",
            RtfDocumentCharacterSet.Pca => @"\pca",
            _ => @"\ansi"
        });

        AppendOptionalTwips(builder, @"\ansicpg", settings.AnsiCodePage);
    }

    private static void WriteSection(StringBuilder builder, RtfSection section, int? defaultLanguageId, int unicodeSkipCount) {
        WriteSectionStart(builder, section);
        foreach (IRtfBlock block in section.Blocks) {
            WriteBlock(builder, block, defaultLanguageId, unicodeSkipCount);
        }

        builder.Append(@"\sect");
        builder.AppendLine();
    }

    private static void WriteBlock(StringBuilder builder, IRtfBlock block, int? defaultLanguageId, int unicodeSkipCount) {
        switch (block) {
            case RtfParagraph paragraph:
                WriteParagraph(builder, paragraph, defaultLanguageId, unicodeSkipCount);
                break;
            case RtfTable table:
                WriteTable(builder, table, defaultLanguageId, unicodeSkipCount);
                break;
            case RtfImage image:
                WriteImage(builder, image);
                break;
            case RtfObject rtfObject:
                WriteObject(builder, rtfObject, defaultLanguageId, unicodeSkipCount);
                builder.AppendLine();
                break;
            case RtfShape shape:
                WriteShape(builder, shape, defaultLanguageId, unicodeSkipCount);
                builder.AppendLine();
                break;
        }
    }

    private static void WriteDocumentSettings(StringBuilder builder, RtfDocumentSettings settings) {
        if (!settings.HasAnyValue) return;

        AppendOptionalTwips(builder, @"\uc", settings.UnicodeSkipCount);
        AppendOptionalTwips(builder, @"\deftab", settings.DefaultTabWidthTwips);
        AppendOptionalTwips(builder, @"\deflang", settings.DefaultLanguageId);
        AppendOptionalTwips(builder, @"\deflangfe", settings.DefaultFarEastLanguageId);
        AppendOptionalTwips(builder, @"\adeflang", settings.DefaultAlternateLanguageId);
        AppendOptionalTwips(builder, @"\viewkind", settings.ViewKind);
        AppendOptionalTwips(builder, @"\viewscale", settings.ViewScale);
        AppendOptionalTwips(builder, @"\viewzk", settings.ZoomKind);
        AppendOptionalTwips(builder, @"\viewbksp", settings.ViewBackspaceBehavior);
        AppendOptionalToggle(builder, @"\widowctrl", settings.WidowOrphanControl);
        AppendOptionalToggle(builder, @"\hyphauto", settings.AutoHyphenation);
        AppendOptionalToggle(builder, @"\hyphcaps", settings.HyphenateCaps);
        AppendOptionalTwips(builder, @"\hyphconsec", settings.ConsecutiveHyphenLimit);
        AppendOptionalTwips(builder, @"\hyphhotz", settings.HyphenationZoneTwips);
        AppendOptionalToggle(builder, @"\facingp", settings.FacingPages);
        AppendOptionalToggle(builder, @"\margmirror", settings.MirrorMargins);
        AppendOptionalToggle(builder, @"\formprot", settings.FormProtection);
        AppendOptionalToggle(builder, @"\revprot", settings.RevisionProtection);
        AppendOptionalToggle(builder, @"\annotprot", settings.AnnotationProtection);
        AppendOptionalToggle(builder, @"\readprot", settings.ReadOnlyProtection);
        AppendOptionalToggle(builder, @"\revisions", settings.TrackRevisions);
        AppendOptionalTwips(builder, @"\revprop", settings.RevisionDisplayStyle);
        AppendOptionalTwips(builder, @"\revbar", settings.RevisionBarPlacement);
        AppendOptionalTwips(builder, @"\dghspace", settings.DrawingGridHorizontalSpacingTwips);
        AppendOptionalTwips(builder, @"\dgvspace", settings.DrawingGridVerticalSpacingTwips);
        AppendOptionalTwips(builder, @"\dghorigin", settings.DrawingGridHorizontalOriginTwips);
        AppendOptionalTwips(builder, @"\dgvorigin", settings.DrawingGridVerticalOriginTwips);
        AppendOptionalTwips(builder, @"\dghshow", settings.DrawingGridHorizontalShow);
        AppendOptionalTwips(builder, @"\dgvshow", settings.DrawingGridVerticalShow);
        AppendOptionalToggle(builder, @"\dgsnap", settings.SnapToDrawingGrid);
        AppendOptionalToggle(builder, @"\dgmargin", settings.DrawingGridUsesMargins);
        if (settings.Direction.HasValue) {
            builder.Append(settings.Direction.Value == RtfTextDirection.RightToLeft ? @"\rtldoc" : @"\ltrdoc");
        }
    }

    private static void WriteHeaderFooters(StringBuilder builder, RtfDocument document, int unicodeSkipCount) {
        foreach (RtfHeaderFooter headerFooter in document.HeaderFooters) {
            builder.Append(@"{\");
            builder.Append(GetHeaderFooterControlWord(headerFooter.Kind));
            foreach (RtfParagraph paragraph in headerFooter.Paragraphs) {
                WriteParagraph(builder, paragraph, document.Settings.DefaultLanguageId, unicodeSkipCount);
            }

            builder.Append('}');
        }
    }

    private static string GetHeaderFooterControlWord(RtfHeaderFooterKind kind) {
        switch (kind) {
            case RtfHeaderFooterKind.LeftHeader:
                return "headerl";
            case RtfHeaderFooterKind.RightHeader:
                return "headerr";
            case RtfHeaderFooterKind.FirstHeader:
                return "headerf";
            case RtfHeaderFooterKind.Footer:
                return "footer";
            case RtfHeaderFooterKind.LeftFooter:
                return "footerl";
            case RtfHeaderFooterKind.RightFooter:
                return "footerr";
            case RtfHeaderFooterKind.FirstFooter:
                return "footerf";
            default:
                return "header";
        }
    }

    private static void WriteInfo(StringBuilder builder, RtfDocument document, int unicodeSkipCount) {
        RtfDocumentInfo info = document.Info;
        if (!HasText(info.Title) && !HasText(info.Subject) && !HasText(info.Author) &&
            !HasText(info.Manager) && !HasText(info.Company) && !HasText(info.Operator) &&
            !HasText(info.Category) && !HasText(info.Keywords) && !HasText(info.Comments) &&
            !HasText(info.HyperlinkBase) && !info.Created.HasValue && !info.Revised.HasValue &&
            !info.Printed.HasValue && !info.BackedUp.HasValue && !info.EditingMinutes.HasValue &&
            !info.NumberOfPages.HasValue && !info.NumberOfWords.HasValue &&
            !info.NumberOfCharacters.HasValue && !info.NumberOfCharactersWithSpaces.HasValue &&
            !info.InternalVersion.HasValue) {
            return;
        }

        builder.Append(@"{\info");
        WriteInfoValue(builder, "title", info.Title, unicodeSkipCount);
        WriteInfoValue(builder, "subject", info.Subject, unicodeSkipCount);
        WriteInfoValue(builder, "author", info.Author, unicodeSkipCount);
        WriteInfoValue(builder, "manager", info.Manager, unicodeSkipCount);
        WriteInfoValue(builder, "company", info.Company, unicodeSkipCount);
        WriteInfoValue(builder, "operator", info.Operator, unicodeSkipCount);
        WriteInfoValue(builder, "category", info.Category, unicodeSkipCount);
        WriteInfoValue(builder, "keywords", info.Keywords, unicodeSkipCount);
        WriteInfoValue(builder, "comment", info.Comments, unicodeSkipCount);
        WriteInfoValue(builder, "hlinkbase", info.HyperlinkBase, unicodeSkipCount);
        WriteInfoTimestamp(builder, "creatim", info.Created);
        WriteInfoTimestamp(builder, "revtim", info.Revised);
        WriteInfoTimestamp(builder, "printim", info.Printed);
        WriteInfoTimestamp(builder, "buptim", info.BackedUp);
        AppendOptionalTwips(builder, @"\edmins", info.EditingMinutes);
        AppendOptionalTwips(builder, @"\nofpages", info.NumberOfPages);
        AppendOptionalTwips(builder, @"\nofwords", info.NumberOfWords);
        AppendOptionalTwips(builder, @"\nofchars", info.NumberOfCharacters);
        AppendOptionalTwips(builder, @"\nofcharsws", info.NumberOfCharactersWithSpaces);
        AppendOptionalTwips(builder, @"\vern", info.InternalVersion);
        builder.Append('}');
    }

    private static bool HasText(string? value) => !string.IsNullOrEmpty(value);

    private static void WriteInfoValue(StringBuilder builder, string name, string? value, int unicodeSkipCount) {
        if (string.IsNullOrEmpty(value)) return;
        builder.Append(@"{\");
        builder.Append(name);
        builder.Append(' ');
        builder.Append(EscapeText(value!, unicodeSkipCount));
        builder.Append('}');
    }

    private static void WriteInfoTimestamp(StringBuilder builder, string name, DateTime? value) {
        if (!value.HasValue) return;

        DateTime timestamp = value.Value;
        builder.Append(@"{\");
        builder.Append(name);
        AppendOptionalTwips(builder, @"\yr", timestamp.Year);
        AppendOptionalTwips(builder, @"\mo", timestamp.Month);
        AppendOptionalTwips(builder, @"\dy", timestamp.Day);
        AppendOptionalTwips(builder, @"\hr", timestamp.Hour);
        AppendOptionalTwips(builder, @"\min", timestamp.Minute);
        AppendOptionalTwips(builder, @"\sec", timestamp.Second);
        builder.Append('}');
    }

    private static void WriteParagraph(StringBuilder builder, RtfParagraph paragraph, int? defaultLanguageId, int unicodeSkipCount) {
        WriteListText(builder, paragraph.ListText, defaultLanguageId, unicodeSkipCount);
        WriteParagraphStart(builder, paragraph, inTable: false, unicodeSkipCount);

        var state = new RunWriteState(defaultLanguageId);
        foreach (IRtfInline inline in paragraph.Inlines) {
            WriteInline(builder, inline, state, defaultLanguageId, unicodeSkipCount);
        }

        ResetRunState(builder, state);
        builder.Append(@"\par");
        builder.AppendLine();
    }

    private static void WriteParagraphStart(StringBuilder builder, RtfParagraph paragraph, bool inTable, int unicodeSkipCount) {
        builder.Append(@"\pard");
        if (inTable) {
            builder.Append(@"\intbl");
        }

        if (paragraph.Direction.HasValue) {
            builder.Append(paragraph.Direction.Value == RtfTextDirection.RightToLeft ? @"\rtlpar" : @"\ltrpar");
        }

        if (paragraph.PageBreakBefore) {
            builder.Append(@"\pagebb");
        }

        if (paragraph.KeepWithNext) {
            builder.Append(@"\keepn");
        }

        if (paragraph.KeepLinesTogether) {
            builder.Append(@"\keep");
        }

        if (paragraph.SuppressLineNumbers) {
            builder.Append(@"\noline");
        }

        if (paragraph.AutoHyphenation.HasValue) {
            builder.Append(paragraph.AutoHyphenation.Value ? @"\hyphpar" : @"\hyphpar0");
        }

        if (paragraph.ContextualSpacing.HasValue) {
            builder.Append(paragraph.ContextualSpacing.Value ? @"\contextualspace" : @"\contextualspace0");
        }

        if (paragraph.AdjustRightIndent.HasValue) {
            builder.Append(paragraph.AdjustRightIndent.Value ? @"\adjustright" : @"\adjustright0");
        }

        if (paragraph.SnapToLineGrid.HasValue) {
            builder.Append(paragraph.SnapToLineGrid.Value ? @"\nosnaplinegrid0" : @"\nosnaplinegrid");
        }

        if (paragraph.WidowControl.HasValue) {
            builder.Append(paragraph.WidowControl.Value ? @"\widctlpar" : @"\nowidctlpar");
        }

        AppendOptionalTwips(builder, @"\outlinelevel", paragraph.OutlineLevel);
        AppendOptionalTwips(builder, @"\pararsid", paragraph.RevisionSaveId);

        if (paragraph.StyleId.HasValue) {
            builder.Append(@"\s");
            builder.Append(paragraph.StyleId.Value.ToString(CultureInfo.InvariantCulture));
        }

        WriteLegacyNumbering(builder, paragraph.LegacyNumbering, unicodeSkipCount);
        if (paragraph.ListId.HasValue) {
            builder.Append(paragraph.ListKind == RtfListKind.Bullet ? @"\pn\pnlvlblt" : @"\pn\pnlvlbody");
            builder.Append(@"\ls");
            builder.Append(paragraph.ListId.Value.ToString(CultureInfo.InvariantCulture));
            builder.Append(@"\ilvl");
            builder.Append((paragraph.ListLevel ?? 0).ToString(CultureInfo.InvariantCulture));
        }

        WriteParagraphFrame(builder, paragraph.Frame);
        WriteTabStops(builder, paragraph.TabStops);
        AppendOptionalTwips(builder, @"\li", paragraph.LeftIndentTwips);
        AppendOptionalTwips(builder, @"\ri", paragraph.RightIndentTwips);
        AppendOptionalTwips(builder, @"\fi", paragraph.FirstLineIndentTwips);
        AppendOptionalTwips(builder, @"\sb", paragraph.SpaceBeforeTwips);
        AppendOptionalTwips(builder, @"\sa", paragraph.SpaceAfterTwips);
        AppendOptionalBinary(builder, @"\sbauto", paragraph.SpaceBeforeAuto);
        AppendOptionalBinary(builder, @"\saauto", paragraph.SpaceAfterAuto);
        AppendOptionalTwips(builder, @"\sl", paragraph.LineSpacingTwips);
        AppendOptionalBinary(builder, @"\slmult", paragraph.LineSpacingMultiple);
        AppendOptionalTwips(builder, @"\cbpat", paragraph.BackgroundColorIndex);
        AppendOptionalTwips(builder, @"\cfpat", paragraph.ShadingForegroundColorIndex);
        AppendOptionalTwips(builder, @"\shading", paragraph.ShadingPatternPercent);
        WriteParagraphShadingPattern(builder, paragraph.ShadingPattern);
        WriteParagraphBorder(builder, @"\brdrt", paragraph.TopBorder);
        WriteParagraphBorder(builder, @"\brdrl", paragraph.LeftBorder);
        WriteParagraphBorder(builder, @"\brdrb", paragraph.BottomBorder);
        WriteParagraphBorder(builder, @"\brdrr", paragraph.RightBorder);
        builder.Append(paragraph.Alignment switch {
            RtfTextAlignment.Center => @"\qc",
            RtfTextAlignment.Right => @"\qr",
            RtfTextAlignment.Justify => @"\qj",
            _ => @"\ql"
        });
        builder.Append(' ');
    }

    private static void WriteTabStops(StringBuilder builder, IEnumerable<RtfTabStop> tabStops) {
        foreach (RtfTabStop tabStop in tabStops) {
            builder.Append(tabStop.Leader switch {
                RtfTabLeader.Dots => @"\tldot",
                RtfTabLeader.MiddleDots => @"\tlmdot",
                RtfTabLeader.Hyphen => @"\tlhyph",
                RtfTabLeader.Underline => @"\tlul",
                RtfTabLeader.ThickLine => @"\tlth",
                RtfTabLeader.EqualSign => @"\tleq",
                _ => string.Empty
            });

            if (tabStop.Alignment == RtfTabAlignment.Bar) {
                builder.Append(@"\tb");
                builder.Append(tabStop.PositionTwips.ToString(CultureInfo.InvariantCulture));
                continue;
            }

            builder.Append(tabStop.Alignment switch {
                RtfTabAlignment.Center => @"\tqc",
                RtfTabAlignment.Right => @"\tqr",
                RtfTabAlignment.Decimal => @"\tqdec",
                _ => string.Empty
            });
            builder.Append(@"\tx");
            builder.Append(tabStop.PositionTwips.ToString(CultureInfo.InvariantCulture));
        }
    }

    private static void WriteParagraphShadingPattern(StringBuilder builder, RtfShadingPattern pattern) {
        string? control = pattern switch {
            RtfShadingPattern.Horizontal => @"\bghoriz",
            RtfShadingPattern.Vertical => @"\bgvert",
            RtfShadingPattern.ForwardDiagonal => @"\bgfdiag",
            RtfShadingPattern.BackwardDiagonal => @"\bgbdiag",
            RtfShadingPattern.Cross => @"\bgcross",
            RtfShadingPattern.DiagonalCross => @"\bgdcross",
            RtfShadingPattern.DarkHorizontal => @"\bgdkhoriz",
            RtfShadingPattern.DarkVertical => @"\bgdkvert",
            RtfShadingPattern.DarkForwardDiagonal => @"\bgdkfdiag",
            RtfShadingPattern.DarkBackwardDiagonal => @"\bgdkbdiag",
            RtfShadingPattern.DarkCross => @"\bgdkcross",
            RtfShadingPattern.DarkDiagonalCross => @"\bgdkdcross",
            _ => null
        };
        if (control != null) {
            builder.Append(control);
        }
    }

    private static void WriteParagraphBorder(StringBuilder builder, string sideControl, RtfParagraphBorder border) {
        if (!border.HasAnyValue) return;

        builder.Append(sideControl);
        builder.Append(border.Style switch {
            RtfParagraphBorderStyle.Double => @"\brdrdb",
            RtfParagraphBorderStyle.Dotted => @"\brdrdot",
            RtfParagraphBorderStyle.Dashed => @"\brdrdash",
            RtfParagraphBorderStyle.None => @"\brdrnil",
            _ => @"\brdrs"
        });
        AppendOptionalTwips(builder, @"\brdrw", border.Width);
        AppendOptionalTwips(builder, @"\brdrcf", border.ColorIndex);
    }

    private static void AppendOptionalTwips(StringBuilder builder, string control, int? value) {
        if (!value.HasValue) return;
        builder.Append(control);
        builder.Append(value.Value.ToString(CultureInfo.InvariantCulture));
    }

    private static void AppendOptionalToggle(StringBuilder builder, string control, bool? value) {
        if (!value.HasValue) return;
        builder.Append(control);
        if (!value.Value) {
            builder.Append('0');
        }
    }

    private static void AppendOptionalBinary(StringBuilder builder, string control, bool? value) {
        if (!value.HasValue) return;
        builder.Append(control);
        builder.Append(value.Value ? '1' : '0');
    }

    private static void WriteImage(StringBuilder builder, RtfImage image) {
        builder.Append(@"{\pict");
        builder.Append(image.Format switch {
            RtfImageFormat.Png => @"\pngblip",
            RtfImageFormat.Jpeg => @"\jpegblip",
            RtfImageFormat.Dib => @"\dibitmap0",
            RtfImageFormat.Wmf => @"\wmetafile8",
            RtfImageFormat.Emf => @"\emfblip",
            _ => @"\bliptag0"
        });

        AppendOptionalTwips(builder, @"\picw", image.SourceWidth);
        AppendOptionalTwips(builder, @"\pich", image.SourceHeight);
        AppendOptionalTwips(builder, @"\picwgoal", image.DesiredWidthTwips);
        AppendOptionalTwips(builder, @"\pichgoal", image.DesiredHeightTwips);
        builder.AppendLine();
        WriteHexBytes(builder, image.Data);
        builder.Append('}');
    }

    private static void WriteHexBytes(StringBuilder builder, byte[] data) {
        for (int i = 0; i < data.Length; i++) {
            builder.Append(data[i].ToString("x2", CultureInfo.InvariantCulture));
            if ((i + 1) % 32 == 0) {
                builder.AppendLine();
            }
        }
    }

    private static void WriteRun(StringBuilder builder, RtfRun run, RunWriteState state, int? defaultLanguageId, int unicodeSkipCount) {
        if (run.Hyperlink != null) {
            builder.Append(@"{\field{\*\fldinst HYPERLINK """);
            builder.Append(EscapeText(run.Hyperlink.ToString(), unicodeSkipCount));
            builder.Append(@"""}{\fldrslt ");
            WriteRunPrefix(builder, run, state);
            builder.Append(EscapeText(run.Text, unicodeSkipCount));
            builder.Append("}}");
            if (run.Note != null) {
                WriteNote(builder, run.Note, defaultLanguageId, unicodeSkipCount);
            }

            return;
        }

        WriteRunPrefix(builder, run, state);
        builder.Append(EscapeText(run.Text, unicodeSkipCount));
        if (run.Note != null) {
            WriteNote(builder, run.Note, defaultLanguageId, unicodeSkipCount);
        }
    }

    private static void WriteInline(StringBuilder builder, IRtfInline inline, RunWriteState state, int? defaultLanguageId, int unicodeSkipCount) {
        switch (inline) {
            case RtfRun run:
                WriteRun(builder, run, state, defaultLanguageId, unicodeSkipCount);
                break;
            case RtfBookmarkMarker marker:
                WriteBookmarkMarker(builder, marker, unicodeSkipCount);
                break;
            case RtfField field:
                WriteField(builder, field, defaultLanguageId, unicodeSkipCount);
                break;
            case RtfGeneratedText generatedText:
                WriteGeneratedText(builder, generatedText);
                if (generatedText.Note != null) {
                    WriteNote(builder, generatedText.Note, defaultLanguageId, unicodeSkipCount);
                }

                break;
            case RtfBreak rtfBreak:
                WriteBreak(builder, rtfBreak);
                break;
            case RtfObject rtfObject:
                ResetRunState(builder, state);
                WriteObject(builder, rtfObject, defaultLanguageId, unicodeSkipCount);
                break;
            case RtfShape shape:
                ResetRunState(builder, state);
                WriteShape(builder, shape, defaultLanguageId, unicodeSkipCount);
                break;
            case RtfImage image:
                ResetRunState(builder, state);
                WriteImage(builder, image);
                break;
        }
    }

    private static void WriteBreak(StringBuilder builder, RtfBreak rtfBreak) {
        builder.Append(rtfBreak.Kind switch {
            RtfBreakKind.SoftLine => @"\softline ",
            RtfBreakKind.Page => @"\page ",
            RtfBreakKind.SoftPage => @"\softpage ",
            RtfBreakKind.Column => @"\column ",
            _ => @"\line "
        });
    }

    private static void WriteGeneratedText(StringBuilder builder, RtfGeneratedText generatedText) {
        builder.Append(generatedText.Kind switch {
            RtfGeneratedTextKind.SectionNumber => @"\sectnum ",
            RtfGeneratedTextKind.CurrentDate => @"\chdate ",
            RtfGeneratedTextKind.CurrentDateLong => @"\chdpl ",
            RtfGeneratedTextKind.CurrentDateAbbreviated => @"\chdpa ",
            RtfGeneratedTextKind.CurrentTime => @"\chtime ",
            RtfGeneratedTextKind.NoteReference => @"\chftn ",
            _ => @"\chpgn "
        });
    }

    private static void WriteField(StringBuilder builder, RtfField field, int? defaultLanguageId, int unicodeSkipCount) {
        builder.Append(@"{\field{\*\fldinst ");
        builder.Append(EscapeText(field.Instruction, unicodeSkipCount));
        builder.Append('}');
        WriteFormFieldData(builder, field.FormFieldData, unicodeSkipCount);
        builder.Append(@"{\fldrslt ");
        var state = new RunWriteState(defaultLanguageId);
        foreach (IRtfInline inline in field.Result.Inlines) {
            WriteInline(builder, inline, state, defaultLanguageId, unicodeSkipCount);
        }

        ResetRunState(builder, state);
        builder.Append("}}");
    }

    private static void WriteBookmarkMarker(StringBuilder builder, RtfBookmarkMarker marker, int unicodeSkipCount) {
        builder.Append(@"{\*\");
        builder.Append(marker.Kind == RtfBookmarkMarkerKind.Start ? "bkmkstart " : "bkmkend ");
        builder.Append(EscapeText(marker.Name, unicodeSkipCount));
        builder.Append('}');
    }

    private static string EscapeText(string text) {
        return RtfTextEncoding.EncodeText(text);
    }

    private static string EscapeText(string text, int unicodeSkipCount) {
        return RtfTextEncoding.EncodeText(text, unicodeSkipCount);
    }

}
