using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>Format-neutral AutoFilter criterion family.</summary>
    public enum ExcelAutoFilterCriteriaKind {
        /// <summary>Explicit values, blanks, or date groups.</summary>
        Values,
        /// <summary>One or two comparison criteria.</summary>
        Custom,
        /// <summary>Top/bottom count or percentage.</summary>
        TopBottom,
        /// <summary>Dynamic date or average criterion.</summary>
        Dynamic,
        /// <summary>Cell or font color criterion.</summary>
        Color,
        /// <summary>Conditional-format icon criterion.</summary>
        Icon,
        /// <summary>Unrecognized preserved criterion.</summary>
        Unsupported
    }

    /// <summary>One format-neutral custom AutoFilter comparison.</summary>
    public sealed class ExcelAutoFilterConditionInfo {
        internal ExcelAutoFilterConditionInfo(string? @operator, string? value) {
            Operator = @operator;
            Value = value;
        }
        /// <summary>Comparison operator name.</summary>
        public string? Operator { get; }
        /// <summary>Authored criterion value.</summary>
        public string? Value { get; }
    }

    /// <summary>One date grouping entry in an AutoFilter value criterion.</summary>
    public sealed class ExcelAutoFilterDateGroupInfo {
        internal ExcelAutoFilterDateGroupInfo(DateGroupItem item) {
            Grouping = item.DateTimeGrouping?.InnerText;
            Year = item.Year?.Value;
            Month = item.Month?.Value;
            Day = item.Day?.Value;
            Hour = item.Hour?.Value;
            Minute = item.Minute?.Value;
            Second = item.Second?.Value;
        }
        /// <summary>Date grouping precision.</summary>
        public string? Grouping { get; }
        /// <summary>Year component.</summary>
        public ushort? Year { get; }
        /// <summary>Month component.</summary>
        public ushort? Month { get; }
        /// <summary>Day component.</summary>
        public ushort? Day { get; }
        /// <summary>Hour component.</summary>
        public ushort? Hour { get; }
        /// <summary>Minute component.</summary>
        public ushort? Minute { get; }
        /// <summary>Second component.</summary>
        public ushort? Second { get; }
    }

    /// <summary>Complete public state for one AutoFilter column.</summary>
    public sealed class ExcelAutoFilterColumnInfo {
        internal ExcelAutoFilterColumnInfo(
            uint columnOffset,
            ExcelAutoFilterCriteriaKind kind,
            IReadOnlyList<string> values,
            IReadOnlyList<ExcelAutoFilterConditionInfo> conditions,
            IReadOnlyList<ExcelAutoFilterDateGroupInfo> dateGroups,
            bool matchAll,
            bool includeBlank,
            bool hiddenButton,
            bool showButton,
            bool? top,
            bool? percent,
            double? topValue,
            string? dynamicType,
            double? dynamicValue,
            double? dynamicMaximum,
            uint? differentialFormatId,
            bool? cellColor,
            string? iconSet,
            uint? iconId) {
            ColumnOffset = columnOffset;
            Kind = kind;
            Values = values;
            Conditions = conditions;
            DateGroups = dateGroups;
            MatchAll = matchAll;
            IncludeBlank = includeBlank;
            HiddenButton = hiddenButton;
            ShowButton = showButton;
            Top = top;
            Percent = percent;
            TopValue = topValue;
            DynamicType = dynamicType;
            DynamicValue = dynamicValue;
            DynamicMaximum = dynamicMaximum;
            DifferentialFormatId = differentialFormatId;
            CellColor = cellColor;
            IconSet = iconSet;
            IconId = iconId;
        }

        /// <summary>Zero-based offset inside the AutoFilter range.</summary>
        public uint ColumnOffset { get; }
        /// <summary>Criterion family.</summary>
        public ExcelAutoFilterCriteriaKind Kind { get; }
        /// <summary>Explicit values.</summary>
        public IReadOnlyList<string> Values { get; }
        /// <summary>Custom comparisons.</summary>
        public IReadOnlyList<ExcelAutoFilterConditionInfo> Conditions { get; }
        /// <summary>Date-group values.</summary>
        public IReadOnlyList<ExcelAutoFilterDateGroupInfo> DateGroups { get; }
        /// <summary>Whether all custom comparisons must match.</summary>
        public bool MatchAll { get; }
        /// <summary>Whether blank values are included.</summary>
        public bool IncludeBlank { get; }
        /// <summary>Whether the filter button is hidden.</summary>
        public bool HiddenButton { get; }
        /// <summary>Whether the filter button is shown.</summary>
        public bool ShowButton { get; }
        /// <summary>Top rather than bottom criterion.</summary>
        public bool? Top { get; }
        /// <summary>Percentage rather than item count criterion.</summary>
        public bool? Percent { get; }
        /// <summary>Top/bottom threshold.</summary>
        public double? TopValue { get; }
        /// <summary>Dynamic criterion type.</summary>
        public string? DynamicType { get; }
        /// <summary>Dynamic criterion value.</summary>
        public double? DynamicValue { get; }
        /// <summary>Dynamic criterion maximum.</summary>
        public double? DynamicMaximum { get; }
        /// <summary>Color differential-format index.</summary>
        public uint? DifferentialFormatId { get; }
        /// <summary>Whether a color filter targets the cell color.</summary>
        public bool? CellColor { get; }
        /// <summary>Icon set name.</summary>
        public string? IconSet { get; }
        /// <summary>Icon index.</summary>
        public uint? IconId { get; }
    }

    /// <summary>Worksheet or table AutoFilter state.</summary>
    public sealed class ExcelAutoFilterInfo {
        internal ExcelAutoFilterInfo(string range, string? tableName, IReadOnlyList<ExcelAutoFilterColumnInfo> columns) {
            Range = range;
            TableName = tableName;
            Columns = columns;
        }
        /// <summary>A1 filter range.</summary>
        public string Range { get; }
        /// <summary>Owning table name, or null for a worksheet filter.</summary>
        public string? TableName { get; }
        /// <summary>Configured filter columns.</summary>
        public IReadOnlyList<ExcelAutoFilterColumnInfo> Columns { get; }
        /// <summary>Whether this filter is owned by a table.</summary>
        public bool IsTableFilter => TableName != null;
    }

    public partial class ExcelSheet {
        /// <summary>Reads worksheet and table AutoFilter criteria/state.</summary>
        public IReadOnlyList<ExcelAutoFilterInfo> GetAutoFilters() {
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                var result = new List<ExcelAutoFilterInfo>();
                AutoFilter? worksheetFilter = WorksheetRoot.GetFirstChild<AutoFilter>();
                if (worksheetFilter != null) result.Add(ReadAutoFilter(worksheetFilter, null));
                foreach (var tablePart in _worksheetPart.TableDefinitionParts) {
                    Table? table = tablePart.Table;
                    AutoFilter? tableFilter = table?.GetFirstChild<AutoFilter>();
                    if (tableFilter != null) {
                        result.Add(ReadAutoFilter(tableFilter, table!.DisplayName?.Value ?? table.Name?.Value));
                    }
                }
                return new ReadOnlyCollection<ExcelAutoFilterInfo>(result);
            });
        }

        /// <summary>Applies an explicit blank criterion by zero-based AutoFilter column offset.</summary>
        public void AutoFilterBlanks(string range, uint columnOffset) => ApplyAutoFilterBlankCriteria(range, columnOffset);

        /// <summary>Applies a top/bottom count or percentage criterion by zero-based AutoFilter column offset.</summary>
        public void AutoFilterTopBottom(string range, uint columnOffset, ushort value, bool top = true, bool percent = false) =>
            ApplyAutoFilterTop10Criteria(range, columnOffset, value, top, percent);

        /// <summary>Clears criteria for one zero-based worksheet AutoFilter column offset.</summary>
        public bool ClearAutoFilterColumn(uint columnOffset) {
            bool removed = false;
            WriteLock(() => {
                AutoFilter? filter = WorksheetRoot.GetFirstChild<AutoFilter>();
                FilterColumn? column = filter?.Elements<FilterColumn>()
                    .FirstOrDefault(item => item.ColumnId?.Value == columnOffset);
                if (column == null) return;
                column.Remove();
                WorksheetRoot.Save();
                removed = true;
            });
            return removed;
        }

        private static ExcelAutoFilterInfo ReadAutoFilter(AutoFilter filter, string? tableName) {
            var columns = filter.Elements<FilterColumn>()
                .Select(ReadAutoFilterColumn)
                .OrderBy(column => column.ColumnOffset)
                .ToArray();
            return new ExcelAutoFilterInfo(
                filter.Reference?.Value ?? string.Empty,
                tableName,
                new ReadOnlyCollection<ExcelAutoFilterColumnInfo>(columns));
        }

        private static ExcelAutoFilterColumnInfo ReadAutoFilterColumn(FilterColumn column) {
            Filters? values = column.GetFirstChild<Filters>();
            CustomFilters? custom = column.GetFirstChild<CustomFilters>();
            Top10? top = column.GetFirstChild<Top10>();
            DynamicFilter? dynamic = column.GetFirstChild<DynamicFilter>();
            ColorFilter? color = column.GetFirstChild<ColorFilter>();
            IconFilter? icon = column.GetFirstChild<IconFilter>();
            ExcelAutoFilterCriteriaKind kind = values != null ? ExcelAutoFilterCriteriaKind.Values
                : custom != null ? ExcelAutoFilterCriteriaKind.Custom
                : top != null ? ExcelAutoFilterCriteriaKind.TopBottom
                : dynamic != null ? ExcelAutoFilterCriteriaKind.Dynamic
                : color != null ? ExcelAutoFilterCriteriaKind.Color
                : icon != null ? ExcelAutoFilterCriteriaKind.Icon
                : ExcelAutoFilterCriteriaKind.Unsupported;
            return new ExcelAutoFilterColumnInfo(
                column.ColumnId?.Value ?? 0U,
                kind,
                new ReadOnlyCollection<string>(values?.Elements<Filter>().Select(item => item.Val?.Value ?? string.Empty).ToArray() ?? System.Array.Empty<string>()),
                new ReadOnlyCollection<ExcelAutoFilterConditionInfo>(custom?.Elements<CustomFilter>().Select(item => new ExcelAutoFilterConditionInfo(item.Operator?.InnerText, item.Val?.Value)).ToArray() ?? System.Array.Empty<ExcelAutoFilterConditionInfo>()),
                new ReadOnlyCollection<ExcelAutoFilterDateGroupInfo>(values?.Elements<DateGroupItem>().Select(item => new ExcelAutoFilterDateGroupInfo(item)).ToArray() ?? System.Array.Empty<ExcelAutoFilterDateGroupInfo>()),
                custom?.And?.Value == true,
                values?.Blank?.Value == true,
                column.HiddenButton?.Value == true,
                column.ShowButton?.Value ?? true,
                top?.Top?.Value,
                top?.Percent?.Value,
                top?.Val?.Value,
                dynamic?.Type?.InnerText,
                dynamic?.Val?.Value,
                dynamic?.MaxVal?.Value,
                color?.FormatId?.Value,
                color?.CellColor?.Value,
                icon?.IconSet?.InnerText,
                icon?.IconId?.Value);
        }
    }
}
