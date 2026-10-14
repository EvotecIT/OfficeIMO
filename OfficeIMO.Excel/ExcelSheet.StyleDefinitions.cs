using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Creates or replaces a custom named style from a standalone definition. Existing cells
        /// using the unchanged named style receive the replacement formatting.
        /// </summary>
        public ExcelNamedStyleInfo DefineNamedStyle(string name, ExcelStyleDefinition definition, bool hidden = false) {
            string normalizedName = ExcelSheet.ValidateNamedStyleName(name);
            if (definition == null) throw new ArgumentNullException(nameof(definition));
            ExcelStyleDefinition snapshot = definition.Snapshot();
            ExcelNamedStyleInfo? result = null;
            Locking.ExecuteWrite(EnsureLock(), () => {
                var part = WorkbookPartRoot.WorkbookStylesPart
                    ?? WorkbookPartRoot.AddNewPart<DocumentFormat.OpenXml.Packaging.WorkbookStylesPart>();
                Stylesheet stylesheet = part.Stylesheet ??= CreateDefaultStylesheet();
                result = ExcelSheet.DefineDeclaredNamedStyle(stylesheet, normalizedName, snapshot, hidden);
                stylesheet.Save();
                MarkPackageDirty();
            });
            return result!;
        }
    }

    public partial class ExcelSheet {
        internal static CellFormat CreateDefinitionFormat(Stylesheet stylesheet, ExcelStyleDefinition definition) {
            EnsureDefaultStylePrimitives(stylesheet);
            CellFormat format = GetBaseCellFormat(stylesheet, 0U);
            format.FontId = GetOrCreateFontVariant(stylesheet, format.FontId?.Value, font => {
                SetBold(font, definition.Bold);
                SetItalic(font, definition.Italic);
                SetUnderline(font, definition.Underline?.ToOpenXml());
                if (definition.FontName != null) SetFontName(font, definition.FontName);
                if (definition.FontSize.HasValue) SetFontSize(font, definition.FontSize.Value);
                if (definition.FontColor.HasValue) SetFontColor(font, definition.FontColor.Value.ToArgbHex());
            });
            format.ApplyFont = true;
            if (definition.BackgroundColor.HasValue) {
                string argb = definition.BackgroundColor.Value.ToArgbHex();
                format.FillId = GetOrCreateFill(stylesheet, new Fill(new PatternFill {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor { Rgb = argb },
                    BackgroundColor = new BackgroundColor { Rgb = argb }
                }));
                format.ApplyFill = true;
            }
            if (definition.WrapText || definition.HorizontalAlignment.HasValue || definition.VerticalAlignment.HasValue) {
                format.Alignment = new Alignment {
                    WrapText = definition.WrapText,
                    Horizontal = definition.HorizontalAlignment?.ToOpenXml(),
                    Vertical = definition.VerticalAlignment?.ToOpenXml()
                };
                format.ApplyAlignment = true;
            }
            if (definition.NumberFormat != null) {
                format.NumberFormatId = GetOrCreateNumberFormatId(stylesheet, definition.NumberFormat);
                format.ApplyNumberFormat = true;
            }
            return format;
        }

        internal static uint AddDefinitionCellFormat(Stylesheet stylesheet, CellFormat format) => AppendOrReuseCellFormat(stylesheet, format);

        internal static uint AddDefinitionNumberFormat(Stylesheet stylesheet, string code) => GetOrCreateNumberFormatId(stylesheet, code);

        internal static ExcelNamedStyleInfo DefineDeclaredNamedStyle(Stylesheet stylesheet, string name, ExcelStyleDefinition definition, bool hidden) {
            EnsureNamedStyleContainers(stylesheet);
            CellStyle? style = stylesheet.CellStyles!.Elements<CellStyle>()
                .FirstOrDefault(item => string.Equals(item.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
            if (style?.BuiltinId != null) throw new InvalidOperationException($"Built-in named style '{name}' cannot be replaced.");
            if (style?.FormatId?.Value is uint existingFormatId && stylesheet.CellStyles.Elements<CellStyle>()
                .Any(item => !ReferenceEquals(item, style) && item.FormatId?.Value == existingFormatId)) {
                throw new InvalidOperationException($"Named style '{name}' shares its base format with another style and cannot be redefined safely.");
            }
            CellFormat format = CreateDefinitionFormat(stylesheet, definition);
            format.FormatId = null;
            uint formatId = ReplaceOrAppendNamedStyleFormat(stylesheet, style, format);
            style ??= stylesheet.CellStyles.AppendChild(new CellStyle());
            style.Name = name;
            style.FormatId = formatId;
            style.Hidden = hidden;
            stylesheet.CellStyles.Count = (uint)stylesheet.CellStyles.Count();
            return new ExcelNamedStyleInfo(name, formatId, null, hidden);
        }
    }
}
