using System.Globalization;
using OfficeIMO.Drawing.Binary;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Maps OfficeArt MSOSPT values to equivalent DrawingML preset geometries.</summary>
    internal static class LegacyPptShapeGeometryMapper {
        internal static bool TryGetPreset(ushort shapeType, out A.ShapeTypeValues preset) {
            A.ShapeTypeValues? value = shapeType switch {
                1 => A.ShapeTypeValues.Rectangle,
                2 => A.ShapeTypeValues.RoundRectangle,
                3 => A.ShapeTypeValues.Ellipse,
                4 => A.ShapeTypeValues.Diamond,
                5 => A.ShapeTypeValues.Triangle,
                6 => A.ShapeTypeValues.RightTriangle,
                7 => A.ShapeTypeValues.Parallelogram,
                8 => A.ShapeTypeValues.Trapezoid,
                9 => A.ShapeTypeValues.Hexagon,
                10 => A.ShapeTypeValues.Octagon,
                11 => A.ShapeTypeValues.Plus,
                12 => A.ShapeTypeValues.Star5,
                13 or 14 => A.ShapeTypeValues.RightArrow,
                15 => A.ShapeTypeValues.HomePlate,
                16 => A.ShapeTypeValues.Cube,
                17 => A.ShapeTypeValues.WedgeRoundRectangleCallout,
                18 => A.ShapeTypeValues.Star16,
                19 => A.ShapeTypeValues.Arc,
                20 => A.ShapeTypeValues.Line,
                21 => A.ShapeTypeValues.Plaque,
                22 => A.ShapeTypeValues.Can,
                23 => A.ShapeTypeValues.Donut,
                32 => A.ShapeTypeValues.StraightConnector1,
                33 => A.ShapeTypeValues.BentConnector2,
                34 => A.ShapeTypeValues.BentConnector3,
                35 => A.ShapeTypeValues.BentConnector4,
                36 => A.ShapeTypeValues.BentConnector5,
                37 => A.ShapeTypeValues.CurvedConnector2,
                38 => A.ShapeTypeValues.CurvedConnector3,
                39 => A.ShapeTypeValues.CurvedConnector4,
                40 => A.ShapeTypeValues.CurvedConnector5,
                41 => A.ShapeTypeValues.Callout1,
                42 => A.ShapeTypeValues.Callout2,
                43 => A.ShapeTypeValues.Callout3,
                44 => A.ShapeTypeValues.AccentCallout1,
                45 => A.ShapeTypeValues.AccentCallout2,
                46 => A.ShapeTypeValues.AccentCallout3,
                47 => A.ShapeTypeValues.BorderCallout1,
                48 => A.ShapeTypeValues.BorderCallout2,
                49 => A.ShapeTypeValues.BorderCallout3,
                50 => A.ShapeTypeValues.AccentBorderCallout1,
                51 => A.ShapeTypeValues.AccentBorderCallout2,
                52 => A.ShapeTypeValues.AccentBorderCallout3,
                53 => A.ShapeTypeValues.Ribbon,
                54 => A.ShapeTypeValues.Ribbon2,
                55 => A.ShapeTypeValues.Chevron,
                56 => A.ShapeTypeValues.Pentagon,
                57 => A.ShapeTypeValues.NoSmoking,
                58 => A.ShapeTypeValues.Star8,
                59 => A.ShapeTypeValues.Star16,
                60 => A.ShapeTypeValues.Star32,
                61 => A.ShapeTypeValues.WedgeRectangleCallout,
                62 => A.ShapeTypeValues.WedgeRoundRectangleCallout,
                63 => A.ShapeTypeValues.WedgeEllipseCallout,
                64 => A.ShapeTypeValues.Wave,
                65 => A.ShapeTypeValues.FoldedCorner,
                66 => A.ShapeTypeValues.LeftArrow,
                67 => A.ShapeTypeValues.DownArrow,
                68 => A.ShapeTypeValues.UpArrow,
                69 => A.ShapeTypeValues.LeftRightArrow,
                70 => A.ShapeTypeValues.UpDownArrow,
                71 => A.ShapeTypeValues.IrregularSeal1,
                72 => A.ShapeTypeValues.IrregularSeal2,
                73 => A.ShapeTypeValues.LightningBolt,
                74 => A.ShapeTypeValues.Heart,
                76 => A.ShapeTypeValues.QuadArrow,
                77 => A.ShapeTypeValues.LeftArrowCallout,
                78 => A.ShapeTypeValues.RightArrowCallout,
                79 => A.ShapeTypeValues.UpArrowCallout,
                80 => A.ShapeTypeValues.DownArrowCallout,
                81 => A.ShapeTypeValues.LeftRightArrowCallout,
                82 => A.ShapeTypeValues.UpDownArrowCallout,
                83 => A.ShapeTypeValues.QuadArrowCallout,
                84 => A.ShapeTypeValues.Bevel,
                85 => A.ShapeTypeValues.LeftBracket,
                86 => A.ShapeTypeValues.RightBracket,
                87 => A.ShapeTypeValues.LeftBrace,
                88 => A.ShapeTypeValues.RightBrace,
                89 => A.ShapeTypeValues.LeftUpArrow,
                90 => A.ShapeTypeValues.BentUpArrow,
                91 => A.ShapeTypeValues.BentArrow,
                92 => A.ShapeTypeValues.Star24,
                93 => A.ShapeTypeValues.StripedRightArrow,
                94 => A.ShapeTypeValues.NotchedRightArrow,
                95 => A.ShapeTypeValues.BlockArc,
                96 => A.ShapeTypeValues.SmileyFace,
                97 => A.ShapeTypeValues.VerticalScroll,
                98 => A.ShapeTypeValues.HorizontalScroll,
                99 or 100 => A.ShapeTypeValues.CircularArrow,
                101 => A.ShapeTypeValues.UTurnArrow,
                102 => A.ShapeTypeValues.CurvedRightArrow,
                103 => A.ShapeTypeValues.CurvedLeftArrow,
                104 => A.ShapeTypeValues.CurvedUpArrow,
                105 => A.ShapeTypeValues.CurvedDownArrow,
                106 => A.ShapeTypeValues.CloudCallout,
                107 => A.ShapeTypeValues.EllipseRibbon,
                108 => A.ShapeTypeValues.EllipseRibbon2,
                109 => A.ShapeTypeValues.FlowChartProcess,
                110 => A.ShapeTypeValues.FlowChartDecision,
                111 => A.ShapeTypeValues.FlowChartInputOutput,
                112 => A.ShapeTypeValues.FlowChartPredefinedProcess,
                113 => A.ShapeTypeValues.FlowChartInternalStorage,
                114 => A.ShapeTypeValues.FlowChartDocument,
                115 => A.ShapeTypeValues.FlowChartMultidocument,
                116 => A.ShapeTypeValues.FlowChartTerminator,
                117 => A.ShapeTypeValues.FlowChartPreparation,
                118 => A.ShapeTypeValues.FlowChartManualInput,
                119 => A.ShapeTypeValues.FlowChartManualOperation,
                120 => A.ShapeTypeValues.FlowChartConnector,
                121 => A.ShapeTypeValues.FlowChartPunchedCard,
                122 => A.ShapeTypeValues.FlowChartPunchedTape,
                123 => A.ShapeTypeValues.FlowChartSummingJunction,
                124 => A.ShapeTypeValues.FlowChartOr,
                125 => A.ShapeTypeValues.FlowChartCollate,
                126 => A.ShapeTypeValues.FlowChartSort,
                127 => A.ShapeTypeValues.FlowChartExtract,
                128 => A.ShapeTypeValues.FlowChartMerge,
                129 => A.ShapeTypeValues.FlowChartOfflineStorage,
                130 => A.ShapeTypeValues.FlowChartOnlineStorage,
                131 => A.ShapeTypeValues.FlowChartMagneticTape,
                132 => A.ShapeTypeValues.FlowChartMagneticDisk,
                133 => A.ShapeTypeValues.FlowChartMagneticDrum,
                134 => A.ShapeTypeValues.FlowChartDisplay,
                135 => A.ShapeTypeValues.FlowChartDelay,
                176 => A.ShapeTypeValues.FlowChartAlternateProcess,
                177 => A.ShapeTypeValues.FlowChartOffpageConnector,
                178 => A.ShapeTypeValues.Callout1,
                179 => A.ShapeTypeValues.AccentCallout1,
                180 => A.ShapeTypeValues.BorderCallout1,
                181 => A.ShapeTypeValues.AccentBorderCallout1,
                182 => A.ShapeTypeValues.LeftRightUpArrow,
                183 => A.ShapeTypeValues.Sun,
                184 => A.ShapeTypeValues.Moon,
                185 => A.ShapeTypeValues.BracketPair,
                186 => A.ShapeTypeValues.BracePair,
                187 => A.ShapeTypeValues.Star4,
                188 => A.ShapeTypeValues.DoubleWave,
                189 => A.ShapeTypeValues.ActionButtonBlank,
                190 => A.ShapeTypeValues.ActionButtonHome,
                191 => A.ShapeTypeValues.ActionButtonHelp,
                192 => A.ShapeTypeValues.ActionButtonInformation,
                193 => A.ShapeTypeValues.ActionButtonForwardNext,
                194 => A.ShapeTypeValues.ActionButtonBackPrevious,
                195 => A.ShapeTypeValues.ActionButtonEnd,
                196 => A.ShapeTypeValues.ActionButtonBeginning,
                197 => A.ShapeTypeValues.ActionButtonReturn,
                198 => A.ShapeTypeValues.ActionButtonDocument,
                199 => A.ShapeTypeValues.ActionButtonSound,
                200 => A.ShapeTypeValues.ActionButtonMovie,
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

        internal static bool TryGetShapeType(A.ShapeTypeValues preset,
            out ushort shapeType) {
            // Prefer the exact canonical MSOSPT value when several legacy
            // values project to the same DrawingML preset.
            for (ushort candidate = 1; candidate <= 202; candidate++) {
                if (!IsApproximation(candidate)
                    && TryGetPreset(candidate, out A.ShapeTypeValues mapped)
                    && mapped == preset) {
                    shapeType = candidate;
                    return true;
                }
            }
            for (ushort candidate = 1; candidate <= 202; candidate++) {
                if (TryGetPreset(candidate, out A.ShapeTypeValues mapped)
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

        internal static void ApplyExactPresetAdjustments(ushort shapeType,
            OfficeArtShapeGeometry source, A.PresetGeometry target) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (shapeType is not 2 and not 23 || !source.AdjustmentValues[0].HasValue) return;

            long scaled = (long)Math.Round(source.AdjustmentValues[0]!.Value * (100000D / 21600D),
                MidpointRounding.AwayFromZero);
            A.AdjustValueList values = target.AdjustValueList ??= new A.AdjustValueList();
            values.RemoveAllChildren<A.ShapeGuide>();
            values.Append(new A.ShapeGuide {
                Name = "adj",
                Formula = "val " + scaled.ToString(CultureInfo.InvariantCulture)
            });
        }
    }
}
