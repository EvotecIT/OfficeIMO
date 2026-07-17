using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a chart on a worksheet.
    /// </summary>
    public sealed partial class ExcelChart {
        private static void ApplyGridlines(OpenXmlCompositeElement axis, bool showMajor, bool showMinor,
            string? lineColor, double? lineWidthPoints) {
            if (showMajor) {
                C.MajorGridlines major = axis.GetFirstChild<C.MajorGridlines>() ?? new C.MajorGridlines();
                ApplyGridlineStyle(major, lineColor, lineWidthPoints);
                if (major.Parent == null) {
                    axis.Append(major);
                }
            } else {
                axis.GetFirstChild<C.MajorGridlines>()?.Remove();
            }

            if (showMinor) {
                C.MinorGridlines minor = axis.GetFirstChild<C.MinorGridlines>() ?? new C.MinorGridlines();
                ApplyGridlineStyle(minor, lineColor, lineWidthPoints);
                if (minor.Parent == null) {
                    axis.Append(minor);
                }
            } else {
                axis.GetFirstChild<C.MinorGridlines>()?.Remove();
            }
        }

        private static void ApplyTrendline(OpenXmlCompositeElement series, C.TrendlineValues type, int? order, int? period,
            double? forward, double? backward, double? intercept, bool displayEquation, bool displayRSquared,
            string? lineColor, double? lineWidthPoints) {
            if (!IsTrendlineSupportedSeries(series)) {
                throw new InvalidOperationException("Trendlines are only supported for line, bar/column, area, scatter, and bubble series.");
            }

            series.RemoveAllChildren<C.Trendline>();
            C.Trendline trendline = new C.Trendline();

            if (lineColor != null || lineWidthPoints != null) {
                C.ChartShapeProperties props = new C.ChartShapeProperties();
                string? normalizedLine = lineColor != null ? NormalizeHexColor(lineColor) : null;
                ApplyOptionalLine(props, normalizedLine, lineWidthPoints);
                trendline.Append(props);
            }

            trendline.Append(new C.TrendlineType { Val = type });

            if (type.Equals(C.TrendlineValues.Polynomial) && order != null) {
                trendline.Append(new C.PolynomialOrder { Val = (byte)order.Value });
            }
            if (type.Equals(C.TrendlineValues.MovingAverage) && period != null) {
                trendline.Append(new C.Period { Val = (uint)period.Value });
            }
            if (forward != null) {
                trendline.Append(new C.Forward { Val = forward.Value });
            }
            if (backward != null) {
                trendline.Append(new C.Backward { Val = backward.Value });
            }
            if (intercept != null) {
                trendline.Append(new C.Intercept { Val = intercept.Value });
            }

            if (displayRSquared) {
                trendline.Append(new C.DisplayRSquaredValue { Val = displayRSquared });
            }
            if (displayEquation) {
                trendline.Append(new C.DisplayEquation { Val = displayEquation });
            }

            InsertTrendline(series, trendline);
        }

        private static bool IsTrendlineSupportedSeries(OpenXmlCompositeElement series) {
            return series is C.LineChartSeries
                || series is C.BarChartSeries
                || series is C.AreaChartSeries
                || series is C.ScatterChartSeries
                || series is C.BubbleChartSeries;
        }

        private static void InsertTrendline(OpenXmlCompositeElement series, C.Trendline trendline) {
            OpenXmlElement? insertBefore = series.GetFirstChild<C.ErrorBars>();
            insertBefore ??= series.GetFirstChild<C.CategoryAxisData>();
            insertBefore ??= series.GetFirstChild<C.Values>();
            insertBefore ??= series.GetFirstChild<C.XValues>();
            insertBefore ??= series.GetFirstChild<C.YValues>();
            insertBefore ??= series.GetFirstChild<C.BubbleSize>();
            insertBefore ??= series.GetFirstChild<C.Smooth>();
            insertBefore ??= series.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                series.InsertBefore(trendline, insertBefore);
            } else {
                series.Append(trendline);
            }
        }

        private static void ApplyGridlineStyle(OpenXmlCompositeElement gridlines, string? lineColor, double? lineWidthPoints) {
            if (lineColor == null && lineWidthPoints == null) {
                return;
            }
            C.ChartShapeProperties props = gridlines.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
            string? normalizedLine = lineColor != null ? NormalizeHexColor(lineColor) : null;
            ApplyOptionalLine(props, normalizedLine, lineWidthPoints);
            if (props.Parent == null) {
                gridlines.Append(props);
            }
        }

        private static void ApplyAxisCrossing(OpenXmlCompositeElement axis, C.CrossesValues crosses, double? crossesAt) {
            axis.GetFirstChild<C.Crosses>()?.Remove();
            axis.GetFirstChild<C.CrossesAt>()?.Remove();

            OpenXmlElement crossing = crossesAt != null
                ? new C.CrossesAt { Val = crossesAt.Value }
                : new C.Crosses { Val = crosses };

            C.CrossingAxis? crossAxis = axis.GetFirstChild<C.CrossingAxis>();
            if (crossAxis != null) {
                axis.InsertAfter(crossing, crossAxis);
            } else {
                axis.Append(crossing);
            }
        }

        private static void ApplyDisplayUnitsLabel(C.DisplayUnits displayUnits, bool showLabel, string? labelText = null) {
            if (!showLabel) {
                displayUnits.GetFirstChild<C.DisplayUnitsLabel>()?.Remove();
                return;
            }

            C.DisplayUnitsLabel label = displayUnits.GetFirstChild<C.DisplayUnitsLabel>() ?? new C.DisplayUnitsLabel();
            if (label.GetFirstChild<C.Layout>() == null) {
                label.Append(new C.Layout());
            }
            if (labelText != null) {
                label.RemoveAllChildren<C.ChartText>();
                label.Append(CreateChartText(labelText));
            }
            if (label.Parent == null) {
                displayUnits.Append(label);
            }
        }

        private static C.ChartShapeProperties EnsureDataLabelShapeProperties(OpenXmlCompositeElement labels) {
            C.ChartShapeProperties props = labels.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
            if (props.Parent == null) {
                labels.Append(props);
            }
            return props;
        }

        private static void ApplyNoFill(OpenXmlCompositeElement props) {
            RemoveShapeFillChoices(props);
            InsertShapeFill(props, new A.NoFill());
        }

        private static void ApplyNoLine(OpenXmlCompositeElement props) {
            A.Outline outline = props.GetFirstChild<A.Outline>() ?? new A.Outline();
            outline.RemoveAllChildren<A.SolidFill>();
            outline.RemoveAllChildren<A.GradientFill>();
            outline.RemoveAllChildren<A.PatternFill>();
            outline.RemoveAllChildren<A.NoFill>();
            outline.PrependChild(new A.NoFill());
            if (outline.Parent == null) {
                OpenXmlElement? insertBefore = props.ChildElements.FirstOrDefault(child =>
                    child.LocalName == "effectLst"
                    || child.LocalName == "effectDag"
                    || child.LocalName == "scene3d"
                    || child.LocalName == "sp3d"
                    || child.LocalName == "extLst");
                if (insertBefore != null) {
                    props.InsertBefore(outline, insertBefore);
                } else {
                    props.Append(outline);
                }
            }
        }

        private static A.DefaultRunProperties EnsureTextPropertiesRunProperties(OpenXmlCompositeElement parent) {
            C.TextProperties textProps = parent.GetFirstChild<C.TextProperties>() ?? new C.TextProperties();
            if (textProps.GetFirstChild<A.BodyProperties>() == null) {
                textProps.Append(new A.BodyProperties());
            }
            if (textProps.GetFirstChild<A.ListStyle>() == null) {
                textProps.Append(new A.ListStyle());
            }

            A.Paragraph paragraph = textProps.GetFirstChild<A.Paragraph>() ?? new A.Paragraph();
            if (paragraph.Parent == null) {
                textProps.Append(paragraph);
            }

            A.ParagraphProperties paragraphProps = paragraph.GetFirstChild<A.ParagraphProperties>() ?? new A.ParagraphProperties();
            if (paragraphProps.Parent == null) {
                paragraph.Append(paragraphProps);
            }

            A.DefaultRunProperties runProps = paragraphProps.GetFirstChild<A.DefaultRunProperties>() ?? new A.DefaultRunProperties();
            if (runProps.Parent == null) {
                paragraphProps.Append(runProps);
            }

            if (textProps.Parent == null) {
                if (parent is C.CategoryAxis || parent is C.ValueAxis) {
                    OpenXmlElement? insertBefore = parent.GetFirstChild<C.CrossingAxis>();
                    insertBefore ??= parent.GetFirstChild<C.Crosses>();
                    insertBefore ??= parent.GetFirstChild<C.CrossesAt>();
                    insertBefore ??= parent.GetFirstChild<C.AutoLabeled>();
                    insertBefore ??= parent.GetFirstChild<C.LabelAlignment>();
                    insertBefore ??= parent.GetFirstChild<C.LabelOffset>();
                    insertBefore ??= parent.GetFirstChild<C.NoMultiLevelLabels>();
                    insertBefore ??= parent.GetFirstChild<C.CrossBetween>();
                    insertBefore ??= parent.GetFirstChild<C.MajorUnit>();
                    insertBefore ??= parent.GetFirstChild<C.MinorUnit>();
                    insertBefore ??= parent.GetFirstChild<C.DisplayUnits>();
                    insertBefore ??= parent.GetFirstChild<C.ExtensionList>();
                    if (insertBefore != null) {
                        parent.InsertBefore(textProps, insertBefore);
                    } else {
                        parent.Append(textProps);
                    }
                } else {
                    parent.Append(textProps);
                }
            }

            return runProps;
        }

        private static A.RunProperties EnsureChartTextRunProperties(C.ChartText chartText) {
            C.RichText richText = chartText.GetFirstChild<C.RichText>() ?? new C.RichText();
            if (richText.GetFirstChild<A.BodyProperties>() == null) {
                richText.Append(new A.BodyProperties());
            }
            if (richText.GetFirstChild<A.ListStyle>() == null) {
                richText.Append(new A.ListStyle());
            }

            A.Paragraph paragraph = richText.GetFirstChild<A.Paragraph>() ?? new A.Paragraph();
            if (paragraph.Parent == null) {
                richText.Append(paragraph);
            }

            A.Run run = paragraph.GetFirstChild<A.Run>() ?? new A.Run();
            if (run.Parent == null) {
                paragraph.Append(run);
            }

            A.RunProperties runProps = run.GetFirstChild<A.RunProperties>() ?? new A.RunProperties();
            if (runProps.Parent == null) {
                run.InsertAt(runProps, 0);
            } else if (runProps != run.FirstChild) {
                runProps.Remove();
                run.InsertAt(runProps, 0);
            }

            if (richText.Parent == null) {
                chartText.Append(richText);
            }

            return runProps;
        }

        private static C.ChartShapeProperties EnsureChartShapeProperties(OpenXmlCompositeElement series) {
            C.ChartShapeProperties props = series.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
            if (props.Parent != null) {
                props.Remove();
            }

            OpenXmlElement? anchor = series.Elements<C.SeriesText>().LastOrDefault();
            anchor ??= series.Elements<C.Order>().LastOrDefault();
            anchor ??= series.Elements<C.Index>().LastOrDefault();

            if (anchor != null) {
                series.InsertAfter(props, anchor);
            } else {
                series.PrependChild(props);
            }

            return props;
        }

        private static void ApplySolidFill(OpenXmlCompositeElement props, string color) {
            RemoveShapeFillChoices(props);
            InsertShapeFill(props, new A.SolidFill(new A.RgbColorModelHex { Val = color }));
        }

        private static void RemoveShapeFillChoices(OpenXmlCompositeElement props) {
            foreach (OpenXmlElement child in props.ChildElements.Where(child =>
                child.LocalName == "noFill"
                || child.LocalName == "solidFill"
                || child.LocalName == "gradFill"
                || child.LocalName == "blipFill"
                || child.LocalName == "pattFill"
                || child.LocalName == "grpFill").ToList()) {
                child.Remove();
            }
        }

        private static void InsertShapeFill(OpenXmlCompositeElement props, OpenXmlElement fill) {
            OpenXmlElement? insertBefore = props.ChildElements.FirstOrDefault(child =>
                child is A.Outline
                || child.LocalName == "effectLst"
                || child.LocalName == "effectDag"
                || child.LocalName == "scene3d"
                || child.LocalName == "sp3d"
                || child.LocalName == "extLst");
            if (insertBefore != null) {
                props.InsertBefore(fill, insertBefore);
            } else {
                props.Append(fill);
            }
        }

        private static void ApplyTextSolidFill(A.TextCharacterPropertiesType props, string color) {
            props.RemoveAllChildren<A.SolidFill>();
            props.RemoveAllChildren<A.NoFill>();
            props.RemoveAllChildren<A.GradientFill>();
            props.RemoveAllChildren<A.PatternFill>();

            var fill = new A.SolidFill(new A.RgbColorModelHex { Val = color });
            OpenXmlElement? insertBefore = props.ChildElements.FirstOrDefault(child =>
                child is A.LatinFont
                || child is A.EastAsianFont
                || child is A.ComplexScriptFont
                || child is A.SymbolFont
                || child is A.HyperlinkOnClick
                || child is A.HyperlinkOnMouseOver
                || child is A.ExtensionList);

            if (insertBefore != null) {
                props.InsertBefore(fill, insertBefore);
            } else {
                props.Append(fill);
            }
        }

        private static void ApplyLine(OpenXmlCompositeElement props, string color, double? widthPoints) {
            A.Outline outline = props.GetFirstChild<A.Outline>() ?? new A.Outline();
            outline.RemoveAllChildren<A.SolidFill>();
            outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
            if (widthPoints != null) {
                outline.Width = (int)Math.Round(widthPoints.Value * 12700d);
            }

            if (outline.Parent == null) {
                props.Append(outline);
            }
        }

        private static void ApplyOptionalLine(OpenXmlCompositeElement props, string? color, double? widthPoints) {
            A.Outline outline = props.GetFirstChild<A.Outline>() ?? new A.Outline();
            if (color != null) {
                outline.RemoveAllChildren<A.SolidFill>();
                outline.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
            }
            if (widthPoints != null) {
                outline.Width = (int)Math.Round(widthPoints.Value * 12700d);
            }
            if (outline.Parent == null) {
                props.Append(outline);
            }
        }

        private static void ApplyMarker(C.Marker marker, C.MarkerStyleValues style, int? size, string? fillColor, string? lineColor, double? lineWidthPoints) {
            marker.Symbol = new C.Symbol { Val = style };
            if (size != null) {
                marker.Size = new C.Size { Val = (byte)size.Value };
            }

            if (fillColor != null || lineColor != null || lineWidthPoints != null) {
                C.ChartShapeProperties props = marker.ChartShapeProperties ?? new C.ChartShapeProperties();
                if (fillColor != null) {
                    ApplySolidFill(props, NormalizeHexColor(fillColor));
                }
                if (lineColor != null || lineWidthPoints != null) {
                    string? normalizedLine = lineColor != null ? NormalizeHexColor(lineColor) : null;
                    ApplyOptionalLine(props, normalizedLine, lineWidthPoints);
                }
                if (props.Parent == null) {
                    marker.Append(props);
                }
            }
        }

    }
}
