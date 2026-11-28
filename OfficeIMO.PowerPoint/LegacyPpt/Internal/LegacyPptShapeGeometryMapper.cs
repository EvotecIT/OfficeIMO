using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Maps OfficeArt MSOSPT values to equivalent DrawingML preset geometries.</summary>
    internal static class LegacyPptShapeGeometryMapper {
        internal static bool TryGetPreset(ushort shapeType, out OfficePresetShapeType preset) {
            OfficePresetShapeType? value = shapeType switch {
                1 => OfficePresetShapeType.Rectangle,
                2 => OfficePresetShapeType.RoundRectangle,
                3 => OfficePresetShapeType.Ellipse,
                4 => OfficePresetShapeType.Diamond,
                5 => OfficePresetShapeType.Triangle,
                6 => OfficePresetShapeType.RightTriangle,
                7 => OfficePresetShapeType.Parallelogram,
                8 => OfficePresetShapeType.Trapezoid,
                9 => OfficePresetShapeType.Hexagon,
                10 => OfficePresetShapeType.Octagon,
                11 => OfficePresetShapeType.Plus,
                12 => OfficePresetShapeType.Star5,
                13 or 14 => OfficePresetShapeType.RightArrow,
                15 => OfficePresetShapeType.HomePlate,
                16 => OfficePresetShapeType.Cube,
                17 => OfficePresetShapeType.WedgeRoundRectangleCallout,
                18 => OfficePresetShapeType.Star16,
                19 => OfficePresetShapeType.Arc,
                20 => OfficePresetShapeType.Line,
                21 => OfficePresetShapeType.Plaque,
                22 => OfficePresetShapeType.Can,
                23 => OfficePresetShapeType.Donut,
                32 => OfficePresetShapeType.StraightConnector1,
                33 => OfficePresetShapeType.BentConnector2,
                34 => OfficePresetShapeType.BentConnector3,
                35 => OfficePresetShapeType.BentConnector4,
                36 => OfficePresetShapeType.BentConnector5,
                37 => OfficePresetShapeType.CurvedConnector2,
                38 => OfficePresetShapeType.CurvedConnector3,
                39 => OfficePresetShapeType.CurvedConnector4,
                40 => OfficePresetShapeType.CurvedConnector5,
                41 => OfficePresetShapeType.Callout1,
                42 => OfficePresetShapeType.Callout2,
                43 => OfficePresetShapeType.Callout3,
                44 => OfficePresetShapeType.AccentCallout1,
                45 => OfficePresetShapeType.AccentCallout2,
                46 => OfficePresetShapeType.AccentCallout3,
                47 => OfficePresetShapeType.BorderCallout1,
                48 => OfficePresetShapeType.BorderCallout2,
                49 => OfficePresetShapeType.BorderCallout3,
                50 => OfficePresetShapeType.AccentBorderCallout1,
                51 => OfficePresetShapeType.AccentBorderCallout2,
                52 => OfficePresetShapeType.AccentBorderCallout3,
                53 => OfficePresetShapeType.Ribbon,
                54 => OfficePresetShapeType.Ribbon2,
                55 => OfficePresetShapeType.Chevron,
                56 => OfficePresetShapeType.Pentagon,
                57 => OfficePresetShapeType.NoSmoking,
                58 => OfficePresetShapeType.Star8,
                59 => OfficePresetShapeType.Star16,
                60 => OfficePresetShapeType.Star32,
                61 => OfficePresetShapeType.WedgeRectangleCallout,
                62 => OfficePresetShapeType.WedgeRoundRectangleCallout,
                63 => OfficePresetShapeType.WedgeEllipseCallout,
                64 => OfficePresetShapeType.Wave,
                65 => OfficePresetShapeType.FoldedCorner,
                66 => OfficePresetShapeType.LeftArrow,
                67 => OfficePresetShapeType.DownArrow,
                68 => OfficePresetShapeType.UpArrow,
                69 => OfficePresetShapeType.LeftRightArrow,
                70 => OfficePresetShapeType.UpDownArrow,
                71 => OfficePresetShapeType.IrregularSeal1,
                72 => OfficePresetShapeType.IrregularSeal2,
                73 => OfficePresetShapeType.LightningBolt,
                74 => OfficePresetShapeType.Heart,
                76 => OfficePresetShapeType.QuadArrow,
                77 => OfficePresetShapeType.LeftArrowCallout,
                78 => OfficePresetShapeType.RightArrowCallout,
                79 => OfficePresetShapeType.UpArrowCallout,
                80 => OfficePresetShapeType.DownArrowCallout,
                81 => OfficePresetShapeType.LeftRightArrowCallout,
                82 => OfficePresetShapeType.UpDownArrowCallout,
                83 => OfficePresetShapeType.QuadArrowCallout,
                84 => OfficePresetShapeType.Bevel,
                85 => OfficePresetShapeType.LeftBracket,
                86 => OfficePresetShapeType.RightBracket,
                87 => OfficePresetShapeType.LeftBrace,
                88 => OfficePresetShapeType.RightBrace,
                89 => OfficePresetShapeType.LeftUpArrow,
                90 => OfficePresetShapeType.BentUpArrow,
                91 => OfficePresetShapeType.BentArrow,
                92 => OfficePresetShapeType.Star24,
                93 => OfficePresetShapeType.StripedRightArrow,
                94 => OfficePresetShapeType.NotchedRightArrow,
                95 => OfficePresetShapeType.BlockArc,
                96 => OfficePresetShapeType.SmileyFace,
                97 => OfficePresetShapeType.VerticalScroll,
                98 => OfficePresetShapeType.HorizontalScroll,
                99 or 100 => OfficePresetShapeType.CircularArrow,
                101 => OfficePresetShapeType.UTurnArrow,
                102 => OfficePresetShapeType.CurvedRightArrow,
                103 => OfficePresetShapeType.CurvedLeftArrow,
                104 => OfficePresetShapeType.CurvedUpArrow,
                105 => OfficePresetShapeType.CurvedDownArrow,
                106 => OfficePresetShapeType.CloudCallout,
                107 => OfficePresetShapeType.EllipseRibbon,
                108 => OfficePresetShapeType.EllipseRibbon2,
                109 => OfficePresetShapeType.FlowChartProcess,
                110 => OfficePresetShapeType.FlowChartDecision,
                111 => OfficePresetShapeType.FlowChartInputOutput,
                112 => OfficePresetShapeType.FlowChartPredefinedProcess,
                113 => OfficePresetShapeType.FlowChartInternalStorage,
                114 => OfficePresetShapeType.FlowChartDocument,
                115 => OfficePresetShapeType.FlowChartMultidocument,
                116 => OfficePresetShapeType.FlowChartTerminator,
                117 => OfficePresetShapeType.FlowChartPreparation,
                118 => OfficePresetShapeType.FlowChartManualInput,
                119 => OfficePresetShapeType.FlowChartManualOperation,
                120 => OfficePresetShapeType.FlowChartConnector,
                121 => OfficePresetShapeType.FlowChartPunchedCard,
                122 => OfficePresetShapeType.FlowChartPunchedTape,
                123 => OfficePresetShapeType.FlowChartSummingJunction,
                124 => OfficePresetShapeType.FlowChartOr,
                125 => OfficePresetShapeType.FlowChartCollate,
                126 => OfficePresetShapeType.FlowChartSort,
                127 => OfficePresetShapeType.FlowChartExtract,
                128 => OfficePresetShapeType.FlowChartMerge,
                129 => OfficePresetShapeType.FlowChartOfflineStorage,
                130 => OfficePresetShapeType.FlowChartOnlineStorage,
                131 => OfficePresetShapeType.FlowChartMagneticTape,
                132 => OfficePresetShapeType.FlowChartMagneticDisk,
                133 => OfficePresetShapeType.FlowChartMagneticDrum,
                134 => OfficePresetShapeType.FlowChartDisplay,
                135 => OfficePresetShapeType.FlowChartDelay,
                176 => OfficePresetShapeType.FlowChartAlternateProcess,
                177 => OfficePresetShapeType.FlowChartOffpageConnector,
                178 => OfficePresetShapeType.Callout1,
                179 => OfficePresetShapeType.AccentCallout1,
                180 => OfficePresetShapeType.BorderCallout1,
                181 => OfficePresetShapeType.AccentBorderCallout1,
                182 => OfficePresetShapeType.LeftRightUpArrow,
                183 => OfficePresetShapeType.Sun,
                184 => OfficePresetShapeType.Moon,
                185 => OfficePresetShapeType.BracketPair,
                186 => OfficePresetShapeType.BracePair,
                187 => OfficePresetShapeType.Star4,
                188 => OfficePresetShapeType.DoubleWave,
                189 => OfficePresetShapeType.ActionButtonBlank,
                190 => OfficePresetShapeType.ActionButtonHome,
                191 => OfficePresetShapeType.ActionButtonHelp,
                192 => OfficePresetShapeType.ActionButtonInformation,
                193 => OfficePresetShapeType.ActionButtonForwardNext,
                194 => OfficePresetShapeType.ActionButtonBackPrevious,
                195 => OfficePresetShapeType.ActionButtonEnd,
                196 => OfficePresetShapeType.ActionButtonBeginning,
                197 => OfficePresetShapeType.ActionButtonReturn,
                198 => OfficePresetShapeType.ActionButtonDocument,
                199 => OfficePresetShapeType.ActionButtonSound,
                200 => OfficePresetShapeType.ActionButtonMovie,
                _ => null
            };
            if (!value.HasValue) {
                preset = default;
                return false;
            }
            preset = value.Value;
            return true;
        }

        internal static bool IsConnector(ushort shapeType) => shapeType is >= 32 and <= 40;

        internal static bool TryGetShapeType(OfficePresetShapeType preset,
            out ushort shapeType) {
            // Prefer the exact canonical MSOSPT value when several legacy
            // values project to the same DrawingML preset.
            for (ushort candidate = 1; candidate <= 202; candidate++) {
                if (!IsApproximation(candidate)
                    && TryGetPreset(candidate, out OfficePresetShapeType mapped)
                    && mapped == preset) {
                    shapeType = candidate;
                    return true;
                }
            }
            for (ushort candidate = 1; candidate <= 202; candidate++) {
                if (TryGetPreset(candidate, out OfficePresetShapeType mapped)
                    && mapped == preset) {
                    shapeType = candidate;
                    return true;
                }
            }
            shapeType = 0;
            return false;
        }

        internal static bool IsApproximation(ushort shapeType) => shapeType is 14 or 17 or 18 or 100
            or >= 178 and <= 181;

        /// <summary>Writes the adjustment of a rounded rectangle or donut onto its <c>a:prstGeom</c>.</summary>
        internal static void ApplyExactPresetAdjustments(ushort shapeType,
            OfficeArtShapeGeometry source, OfficeXmlNode presetGeometry) {
            ArgumentNullException.ThrowIfNull(source);
            if (presetGeometry.IsNone) throw new ArgumentNullException(nameof(presetGeometry));
            if (shapeType is not 2 and not 23 || !source.AdjustmentValues[0].HasValue) return;

            long scaled = (long)Math.Round(source.AdjustmentValues[0]!.Value * (100000D / 21600D),
                MidpointRounding.AwayFromZero);
            OfficeXmlNode values = presetGeometry.Child(LegacyPptXml.A, "avLst");
            if (values.IsNone) values = presetGeometry.AppendNew(LegacyPptXml.A, "avLst");
            values.RemoveChildren(LegacyPptXml.A, "gd");
            OfficeXmlNode guide = values.AppendNew(LegacyPptXml.A, "gd");
            guide.SetAttribute("name", "adj");
            guide.SetAttribute("fmla", "val " + scaled.ToString(CultureInfo.InvariantCulture));
        }
    }
}
