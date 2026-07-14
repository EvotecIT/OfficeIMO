using System;
using System.Collections.Generic;
using System.Text;
using CodeGlyphX;
using CodeGlyphX.DataMatrix;
using CodeGlyphX.Rendering.Png;
using CodeGlyphX.Rendering.Svg;

internal static class SvgScenarioFactory {
    internal static IEnumerable<string> Names {
        get {
            yield return "QR default";
            yield return "QR styled circles";
            yield return "Data Matrix";
            yield return "DataBar Expanded stacked";
            yield return "Code 128 with text";
        }
    }

    internal static SvgScenario Create(string name) {
        string svg;
        switch (name) {
            case "QR default":
                QrCode defaultQr = QrCode.Encode("https://evotec.xyz/codeglyphx/benchmark");
                svg = SvgQrRenderer.Render(defaultQr.Modules, new QrSvgRenderOptions());
                break;
            case "QR styled circles":
                QrCode styledQr = QrCode.Encode("STYLED-QR-OFFICEIMO-BENCHMARK");
                svg = SvgQrRenderer.Render(styledQr.Modules, new QrSvgRenderOptions {
                    ModuleShape = QrPngModuleShape.Circle,
                    ModuleScale = 0.78,
                    DarkColor = "#2457A6",
                    LightColor = "#F7FAFF"
                });
                break;
            case "Data Matrix":
                BitMatrix dataMatrix = DataMatrixEncoder.Encode("OFFICEIMO-DATAMATRIX-BENCHMARK-1234567890");
                svg = MatrixSvgRenderer.Render(dataMatrix, new MatrixSvgRenderOptions());
                break;
            case "DataBar Expanded stacked":
                BitMatrix dataBar = MatrixBarcodeEncoder.Encode(BarcodeType.GS1DataBarExpandedStacked, "1234567890");
                svg = MatrixSvgRenderer.Render(dataBar, new MatrixSvgRenderOptions());
                break;
            case "Code 128 with text":
                const string label = "ORDER-1234-BENCHMARK";
                Barcode1D barcode = BarcodeEncoder.Encode(BarcodeType.Code128, label);
                svg = SvgBarcodeRenderer.Render(barcode, new BarcodeSvgRenderOptions { LabelText = label });
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(name), name, "Unknown SVG benchmark scenario.");
        }
        return new SvgScenario(name, Encoding.UTF8.GetBytes(svg));
    }
}

internal sealed class SvgScenario {
    internal string Name { get; }
    internal byte[] Svg { get; }

    internal SvgScenario(string name, byte[] svg) {
        Name = name;
        Svg = svg;
    }
}
