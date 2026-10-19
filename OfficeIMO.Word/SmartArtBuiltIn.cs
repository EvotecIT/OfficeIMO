using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word.SmartArt.Templates;

namespace OfficeIMO.Word {
    internal static class SmartArtBuiltIn {
        internal static (string relLayout, string relColors, string relStyle, string relData) AddParts(OpenXmlPart owner, WordSmartArtType type) {
            switch (type) {
                case WordSmartArtType.Cycle:
                    return AddCycle(owner);
                case WordSmartArtType.BasicProcess:
                    return AddBasicProcess(owner);
                case WordSmartArtType.CustomSmartArt1:
                    return AddCustom1(owner);
                case WordSmartArtType.CustomSmartArt2:
                    return AddCustom2(owner);
                case WordSmartArtType.Hierarchy:
                    return AddHierarchy(owner);
                case WordSmartArtType.PictureOrgChart:
                    return AddPictureOrgChart(owner);
                case WordSmartArtType.ContinuousBlockProcess:
                    return AddContinuousBlockProcess(owner);
                default:
                    return AddBasicProcess(owner);
            }
        }

        private static (string relLayout, string relColors, string relStyle, string relData) AddBasicProcess(OpenXmlPart owner) {
            var layoutPart = owner.AddNewPart<DiagramLayoutDefinitionPart>();
            SmartArtBasicProcessLayout.PopulateLayout(layoutPart);
            var colorsPart = owner.AddNewPart<DiagramColorsPart>();
            SmartArtCommonColors.PopulateColors(colorsPart);
            var stylePart = owner.AddNewPart<DiagramStylePart>();
            SmartArtCommonStyle.PopulateStyle(stylePart);
            var dataPart = owner.AddNewPart<DiagramDataPart>();
            SmartArtBasicProcessData.PopulateData(dataPart);

            return (
                owner.GetIdOfPart(layoutPart)!,
                owner.GetIdOfPart(colorsPart)!,
                owner.GetIdOfPart(stylePart)!,
                owner.GetIdOfPart(dataPart)!
            );
        }

        private static (string relLayout, string relColors, string relStyle, string relData) AddCycle(OpenXmlPart owner) {
            var layoutPart = owner.AddNewPart<DiagramLayoutDefinitionPart>();
            SmartArtCycleLayout.PopulateLayout(layoutPart);
            var colorsPart = owner.AddNewPart<DiagramColorsPart>();
            SmartArtCommonColors.PopulateColors(colorsPart);
            var stylePart = owner.AddNewPart<DiagramStylePart>();
            SmartArtCommonStyle.PopulateStyle(stylePart);
            var dataPart = owner.AddNewPart<DiagramDataPart>();
            SmartArtCycleData.PopulateData(dataPart);

            return (
                owner.GetIdOfPart(layoutPart)!,
                owner.GetIdOfPart(colorsPart)!,
                owner.GetIdOfPart(stylePart)!,
                owner.GetIdOfPart(dataPart)!
            );
        }

        private static (string relLayout, string relColors, string relStyle, string relData) AddHierarchy(OpenXmlPart owner) {
            var layoutPart = owner.AddNewPart<DiagramLayoutDefinitionPart>();
            SmartArtHierarchyLayout.PopulateLayout(layoutPart);
            var colorsPart = owner.AddNewPart<DiagramColorsPart>();
            SmartArtCommonColors.PopulateColors(colorsPart);
            var stylePart = owner.AddNewPart<DiagramStylePart>();
            SmartArtCommonStyle.PopulateStyle(stylePart);
            var dataPart = owner.AddNewPart<DiagramDataPart>();
            SmartArtHierarchyData.PopulateData(dataPart);

            return (
                owner.GetIdOfPart(layoutPart)!,
                owner.GetIdOfPart(colorsPart)!,
                owner.GetIdOfPart(stylePart)!,
                owner.GetIdOfPart(dataPart)!
            );
        }

        private static (string relLayout, string relColors, string relStyle, string relData) AddPictureOrgChart(OpenXmlPart owner) {
            var layoutPart = owner.AddNewPart<DiagramLayoutDefinitionPart>();
            SmartArtPictureOrgChartLayout.PopulateLayout(layoutPart);
            var colorsPart = owner.AddNewPart<DiagramColorsPart>();
            SmartArtCommonColors.PopulateColors(colorsPart);
            var stylePart = owner.AddNewPart<DiagramStylePart>();
            SmartArtCommonStyle.PopulateStyle(stylePart);
            var dataPart = owner.AddNewPart<DiagramDataPart>();
            SmartArtPictureOrgChartData.PopulateData(dataPart);

            return (
                owner.GetIdOfPart(layoutPart)!,
                owner.GetIdOfPart(colorsPart)!,
                owner.GetIdOfPart(stylePart)!,
                owner.GetIdOfPart(dataPart)!
            );
        }

        private static (string relLayout, string relColors, string relStyle, string relData) AddContinuousBlockProcess(OpenXmlPart owner) {
            var layoutPart = owner.AddNewPart<DiagramLayoutDefinitionPart>();
            SmartArtContinuousBlockProcessLayout.PopulateLayout(layoutPart);
            var colorsPart = owner.AddNewPart<DiagramColorsPart>();
            SmartArtCommonColors.PopulateColors(colorsPart);
            var stylePart = owner.AddNewPart<DiagramStylePart>();
            SmartArtCommonStyle.PopulateStyle(stylePart);
            var dataPart = owner.AddNewPart<DiagramDataPart>();
            SmartArtContinuousBlockProcessData.PopulateData(dataPart);

            return (
                owner.GetIdOfPart(layoutPart)!,
                owner.GetIdOfPart(colorsPart)!,
                owner.GetIdOfPart(stylePart)!,
                owner.GetIdOfPart(dataPart)!
            );
        }

        private static (string relLayout, string relColors, string relStyle, string relData) AddCustom1(OpenXmlPart owner) {
            var layoutPart = owner.AddNewPart<DiagramLayoutDefinitionPart>();
            SmartArt.Templates.SmartArtCustom1.PopulateLayout(layoutPart);
            var colorsPart = owner.AddNewPart<DiagramColorsPart>();
            SmartArt.Templates.SmartArtCustom1.PopulateColors(colorsPart);
            var stylePart = owner.AddNewPart<DiagramStylePart>();
            SmartArt.Templates.SmartArtCustom1.PopulateStyle(stylePart);

            // Optional persisted layout for exact positioning
            var persistPart = owner.AddNewPart<DiagramPersistLayoutPart>();
            SmartArt.Templates.SmartArtCustom1.PopulatePersistLayout(persistPart);
            var persistRel = owner.GetIdOfPart(persistPart)!;

            var dataPart = owner.AddNewPart<DiagramDataPart>();
            SmartArt.Templates.SmartArtCustom1.PopulateData(dataPart, persistRel);

            return (
                owner.GetIdOfPart(layoutPart)!,
                owner.GetIdOfPart(colorsPart)!,
                owner.GetIdOfPart(stylePart)!,
                owner.GetIdOfPart(dataPart)!
            );
        }

        private static (string relLayout, string relColors, string relStyle, string relData) AddCustom2(OpenXmlPart owner) {
            var layoutPart = owner.AddNewPart<DiagramLayoutDefinitionPart>();
            SmartArt.Templates.SmartArtCustom2.PopulateLayout(layoutPart);
            var colorsPart = owner.AddNewPart<DiagramColorsPart>();
            SmartArt.Templates.SmartArtCustom2.PopulateColors(colorsPart);
            var stylePart = owner.AddNewPart<DiagramStylePart>();
            SmartArt.Templates.SmartArtCustom2.PopulateStyle(stylePart);

            var persistPart = owner.AddNewPart<DiagramPersistLayoutPart>();
            SmartArt.Templates.SmartArtCustom2.PopulatePersistLayout(persistPart);
            var persistRel = owner.GetIdOfPart(persistPart)!;

            var dataPart = owner.AddNewPart<DiagramDataPart>();
            SmartArt.Templates.SmartArtCustom2.PopulateData(dataPart, persistRel);

            return (
                owner.GetIdOfPart(layoutPart)!,
                owner.GetIdOfPart(colorsPart)!,
                owner.GetIdOfPart(stylePart)!,
                owner.GetIdOfPart(dataPart)!
            );
        }
    }
}
