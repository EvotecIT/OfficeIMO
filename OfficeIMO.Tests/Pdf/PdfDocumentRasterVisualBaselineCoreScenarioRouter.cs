using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentRasterVisualBaselineTests {
    private static byte[] CreateCoreScenario(string scenarioName) {
        switch (scenarioName) {
            case "hello-world":
                return CreateHelloWorld();
            case "core-layout":
                return CreateCoreLayout();
            case "style-cheatsheet":
                return CreateStyleCheatsheet();
            case "links-rules":
                return CreateLinksAndRules();
            case "lists-tables":
                return CreateListsTables();
            case "table-style-gallery":
                return CreateTableStyleGallery();
            case "default-styles":
                return CreateDefaultStyles();
            case "styled-runs":
                return CreateStyledRuns();
            case "tabs-leaders":
                return CreateTabsLeaders();
            case "drawing-gallery":
                return CreateDrawingGallery();
            case "watermark":
                return CreateWatermark();
            case "image-watermark":
                return CreateImageWatermark();
            case "page-border":
                return CreatePageBorder();
            case "background-image":
                return CreateBackgroundImage();
            case "background-shapes":
                return CreateBackgroundShapes();
            case "row-columns":
                return CreateRowColumns();
            case "showcase-dashboard":
                return CreateShowcaseDashboard();
            default:
                throw new ArgumentOutOfRangeException(nameof(scenarioName), scenarioName, "Unknown PDF raster scenario.");
        }
    }
}
