using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word.SmartArt.Templates;

namespace OfficeIMO.Word {
    internal static class SmartArtBuiltIn {
        internal static (string relLayout, string relColors, string relStyle, string relData) AddParts(MainDocumentPart mainPart, SmartArtType type) {
            switch (type) {
                case SmartArtType.Cycle:
                    return AddCycle(mainPart);
                case SmartArtType.BasicProcess:
                    return AddBasicProcess(mainPart);
                case SmartArtType.CustomSmartArt1:
                    return AddCustom1(mainPart);
                case SmartArtType.CustomSmartArt2:
                    return AddCustom2(mainPart);
                case SmartArtType.Hierarchy:
                    return AddHierarchy(mainPart);
                case SmartArtType.PictureOrgChart:
                    return AddPictureOrgChart(mainPart);
                case SmartArtType.ContinuousBlockProcess:
                    return AddContinuousBlockProcess(mainPart);
                default:
                    return AddBasicProcess(mainPart);
            }
        }

        private static (string relLayout, string relColors, string relStyle, string relData) AddBasicProcess(MainDocumentPart mainPart) {
            var layoutPart = mainPart.AddNewPart<DiagramLayoutDefinitionPart>();
            SmartArtBasicProcessLayout.PopulateLayout(layoutPart);
            var colorsPart = mainPart.AddNewPart<DiagramColorsPart>();
            SmartArtCommonColors.PopulateColors(colorsPart);
            var stylePart = mainPart.AddNewPart<DiagramStylePart>();
            SmartArtCommonStyle.PopulateStyle(stylePart);
            var dataPart = mainPart.AddNewPart<DiagramDataPart>();
            SmartArtBasicProcessData.PopulateData(dataPart);

            return (
                mainPart.GetIdOfPart(layoutPart)!,
                mainPart.GetIdOfPart(colorsPart)!,
                mainPart.GetIdOfPart(stylePart)!,
                mainPart.GetIdOfPart(dataPart)!
            );
        }

        private static (string relLayout, string relColors, string relStyle, string relData) AddCycle(MainDocumentPart mainPart) {
            var layoutPart = mainPart.AddNewPart<DiagramLayoutDefinitionPart>();
            SmartArtCycleLayout.PopulateLayout(layoutPart);
            var colorsPart = mainPart.AddNewPart<DiagramColorsPart>();
            SmartArtCommonColors.PopulateColors(colorsPart);
            var stylePart = mainPart.AddNewPart<DiagramStylePart>();
            SmartArtCommonStyle.PopulateStyle(stylePart);
            var dataPart = mainPart.AddNewPart<DiagramDataPart>();
            SmartArtCycleData.PopulateData(dataPart);

            return (
                mainPart.GetIdOfPart(layoutPart)!,
                mainPart.GetIdOfPart(colorsPart)!,
                mainPart.GetIdOfPart(stylePart)!,
                mainPart.GetIdOfPart(dataPart)!
            );
        }

        private static (string relLayout, string relColors, string relStyle, string relData) AddHierarchy(MainDocumentPart mainPart) {
            var layoutPart = mainPart.AddNewPart<DiagramLayoutDefinitionPart>();
            SmartArtHierarchyLayout.PopulateLayout(layoutPart);
            var colorsPart = mainPart.AddNewPart<DiagramColorsPart>();
            SmartArtCommonColors.PopulateColors(colorsPart);
            var stylePart = mainPart.AddNewPart<DiagramStylePart>();
            SmartArtCommonStyle.PopulateStyle(stylePart);
            var dataPart = mainPart.AddNewPart<DiagramDataPart>();
            SmartArtHierarchyData.PopulateData(dataPart);

            return (
                mainPart.GetIdOfPart(layoutPart)!,
                mainPart.GetIdOfPart(colorsPart)!,
                mainPart.GetIdOfPart(stylePart)!,
                mainPart.GetIdOfPart(dataPart)!
            );
        }

        private static (string relLayout, string relColors, string relStyle, string relData) AddPictureOrgChart(MainDocumentPart mainPart) {
            var layoutPart = mainPart.AddNewPart<DiagramLayoutDefinitionPart>();
            SmartArtPictureOrgChartLayout.PopulateLayout(layoutPart);
            var colorsPart = mainPart.AddNewPart<DiagramColorsPart>();
            SmartArtCommonColors.PopulateColors(colorsPart);
            var stylePart = mainPart.AddNewPart<DiagramStylePart>();
            SmartArtCommonStyle.PopulateStyle(stylePart);
            var dataPart = mainPart.AddNewPart<DiagramDataPart>();
            SmartArtPictureOrgChartData.PopulateData(dataPart);

            return (
                mainPart.GetIdOfPart(layoutPart)!,
                mainPart.GetIdOfPart(colorsPart)!,
                mainPart.GetIdOfPart(stylePart)!,
                mainPart.GetIdOfPart(dataPart)!
            );
        }

        private static (string relLayout, string relColors, string relStyle, string relData) AddContinuousBlockProcess(MainDocumentPart mainPart) {
            var layoutPart = mainPart.AddNewPart<DiagramLayoutDefinitionPart>();
            SmartArtContinuousBlockProcessLayout.PopulateLayout(layoutPart);
            var colorsPart = mainPart.AddNewPart<DiagramColorsPart>();
            SmartArtCommonColors.PopulateColors(colorsPart);
            var stylePart = mainPart.AddNewPart<DiagramStylePart>();
            SmartArtCommonStyle.PopulateStyle(stylePart);
            var dataPart = mainPart.AddNewPart<DiagramDataPart>();
            SmartArtContinuousBlockProcessData.PopulateData(dataPart);

            return (
                mainPart.GetIdOfPart(layoutPart)!,
                mainPart.GetIdOfPart(colorsPart)!,
                mainPart.GetIdOfPart(stylePart)!,
                mainPart.GetIdOfPart(dataPart)!
            );
        }

        private static (string relLayout, string relColors, string relStyle, string relData) AddCustom1(MainDocumentPart mainPart) {
            var layoutPart = mainPart.AddNewPart<DiagramLayoutDefinitionPart>();
            SmartArt.Templates.SmartArtCustom1.PopulateLayout(layoutPart);
            var colorsPart = mainPart.AddNewPart<DiagramColorsPart>();
            SmartArt.Templates.SmartArtCustom1.PopulateColors(colorsPart);
            var stylePart = mainPart.AddNewPart<DiagramStylePart>();
            SmartArt.Templates.SmartArtCustom1.PopulateStyle(stylePart);

            // Optional persisted layout for exact positioning
            var persistPart = mainPart.AddNewPart<DiagramPersistLayoutPart>();
            SmartArt.Templates.SmartArtCustom1.PopulatePersistLayout(persistPart);
            var persistRel = mainPart.GetIdOfPart(persistPart)!;

            var dataPart = mainPart.AddNewPart<DiagramDataPart>();
            SmartArt.Templates.SmartArtCustom1.PopulateData(dataPart, persistRel);

            return (
                mainPart.GetIdOfPart(layoutPart)!,
                mainPart.GetIdOfPart(colorsPart)!,
                mainPart.GetIdOfPart(stylePart)!,
                mainPart.GetIdOfPart(dataPart)!
            );
        }

        private static (string relLayout, string relColors, string relStyle, string relData) AddCustom2(MainDocumentPart mainPart) {
            var layoutPart = mainPart.AddNewPart<DiagramLayoutDefinitionPart>();
            SmartArt.Templates.SmartArtCustom2.PopulateLayout(layoutPart);
            var colorsPart = mainPart.AddNewPart<DiagramColorsPart>();
            SmartArt.Templates.SmartArtCustom2.PopulateColors(colorsPart);
            var stylePart = mainPart.AddNewPart<DiagramStylePart>();
            SmartArt.Templates.SmartArtCustom2.PopulateStyle(stylePart);

            var persistPart = mainPart.AddNewPart<DiagramPersistLayoutPart>();
            SmartArt.Templates.SmartArtCustom2.PopulatePersistLayout(persistPart);
            var persistRel = mainPart.GetIdOfPart(persistPart)!;

            var dataPart = mainPart.AddNewPart<DiagramDataPart>();
            SmartArt.Templates.SmartArtCustom2.PopulateData(dataPart, persistRel);

            return (
                mainPart.GetIdOfPart(layoutPart)!,
                mainPart.GetIdOfPart(colorsPart)!,
                mainPart.GetIdOfPart(stylePart)!,
                mainPart.GetIdOfPart(dataPart)!
            );
        }
    }
}
