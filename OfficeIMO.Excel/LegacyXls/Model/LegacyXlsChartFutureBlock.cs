namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes preserve-only StartBlock and EndBlock chart future-record scope metadata.
    /// </summary>
    public sealed class LegacyXlsChartFutureBlock {
        internal LegacyXlsChartFutureBlock(
            bool isStart,
            ushort objectKind,
            ushort? objectContext,
            ushort? objectInstance1,
            ushort? objectInstance2) {
            IsStart = isStart;
            ObjectKind = objectKind;
            ObjectContext = objectContext;
            ObjectInstance1 = objectInstance1;
            ObjectInstance2 = objectInstance2;
        }

        /// <summary>Gets whether this record starts a future-record block.</summary>
        public bool IsStart { get; }

        /// <summary>Gets whether this record ends a future-record block.</summary>
        public bool IsEnd => !IsStart;

        /// <summary>Gets the future-record block direction.</summary>
        public string DirectionName => IsStart ? "StartBlock" : "EndBlock";

        /// <summary>Gets the raw future-record object kind.</summary>
        public ushort ObjectKind { get; }

        /// <summary>Gets the decoded future-record object kind name.</summary>
        public string ObjectKindName => ObjectKind switch {
            0x0000 => "AxisGroup",
            0x0002 => "AttachedLabel",
            0x0004 => "Axis",
            0x0005 => "ChartGroup",
            0x0006 => "DataTable",
            0x0007 => "Frame",
            0x0009 => "Legend",
            0x000A => "LegendException",
            0x000C => "Series",
            0x000D => "Sheet",
            0x000E => "DataFormat",
            0x000F => "DropBar",
            _ => $"Unknown:0x{ObjectKind:X4}"
        };

        /// <summary>Gets the raw StartBlock object context, when present.</summary>
        public ushort? ObjectContext { get; }

        /// <summary>Gets the raw StartBlock first object instance, when present.</summary>
        public ushort? ObjectInstance1 { get; }

        /// <summary>Gets the raw StartBlock second object instance, when present.</summary>
        public ushort? ObjectInstance2 { get; }

        /// <summary>Gets a compact key describing the future-record block scope.</summary>
        public string ScopeKey {
            get {
                if (!IsStart) {
                    return $"Kind:{ObjectKindName}";
                }

                return $"Kind:{ObjectKindName};Context:0x{ObjectContext!.Value:X4};Instance1:0x{ObjectInstance1!.Value:X4};Instance2:0x{ObjectInstance2!.Value:X4}";
            }
        }
    }
}
