using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderTextFieldBlock(TextFieldBlock block, double containerX, double containerWidth) {
            double spacingBefore = ResolveTopLevelSpacingBefore(block.SpacingBefore);
            double needed = spacingBefore + block.Height + block.SpacingAfter;
            EnsureFixedFlowBlockFits("Text field", block.Width, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }

            if (spacingBefore > 0) {
                y -= spacingBefore;
            }

            double x = GetAlignedObjectX(containerX, containerWidth, block.Width, block.Align);
            currentPage!.FormFields.Add(new FormFieldAnnotation {
                X1 = x,
                Y1 = y - block.Height,
                X2 = x + block.Width,
                Y2 = y,
                Kind = FormFieldAnnotationKind.Text,
                Name = block.Name,
                Value = block.Value,
                FontSize = block.FontSize,
                Style = block.Style
            });
            pageDirty = true;
            y -= block.Height + block.SpacingAfter;
        }

        private void RenderCheckBoxBlock(CheckBoxBlock block, double containerX, double containerWidth) {
            double spacingBefore = ResolveTopLevelSpacingBefore(block.SpacingBefore);
            double needed = spacingBefore + block.Size + block.SpacingAfter;
            EnsureFixedFlowBlockFits("Check box", block.Size, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }

            if (spacingBefore > 0) {
                y -= spacingBefore;
            }

            double x = GetAlignedObjectX(containerX, containerWidth, block.Size, block.Align);
            currentPage!.FormFields.Add(new FormFieldAnnotation {
                X1 = x,
                Y1 = y - block.Size,
                X2 = x + block.Size,
                Y2 = y,
                Kind = FormFieldAnnotationKind.CheckBox,
                Name = block.Name,
                Value = block.IsChecked ? block.CheckedValueName : "Off",
                IsChecked = block.IsChecked,
                CheckedValueName = block.CheckedValueName,
                Style = block.Style
            });
            pageDirty = true;
            y -= block.Size + block.SpacingAfter;
        }

        private void RenderChoiceFieldBlock(ChoiceFieldBlock block, double containerX, double containerWidth) {
            double spacingBefore = ResolveTopLevelSpacingBefore(block.SpacingBefore);
            double needed = spacingBefore + block.Height + block.SpacingAfter;
            EnsureFixedFlowBlockFits("Choice field", block.Width, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }

            if (spacingBefore > 0) {
                y -= spacingBefore;
            }

            double x = GetAlignedObjectX(containerX, containerWidth, block.Width, block.Align);
            currentPage!.FormFields.Add(new FormFieldAnnotation {
                X1 = x,
                Y1 = y - block.Height,
                X2 = x + block.Width,
                Y2 = y,
                Kind = FormFieldAnnotationKind.Choice,
                Name = block.Name,
                Value = block.Value,
                Values = block.Values,
                FontSize = block.FontSize,
                Options = block.Options,
                IsComboBox = block.IsComboBox,
                AllowsMultipleSelection = block.AllowsMultipleSelection,
                Style = block.Style
            });
            pageDirty = true;
            y -= block.Height + block.SpacingAfter;
        }

        private void RenderRadioButtonGroupBlock(RadioButtonGroupBlock block, double containerX, double containerWidth) {
            double spacingBefore = ResolveTopLevelSpacingBefore(block.SpacingBefore);
            double height = block.Height;
            double needed = spacingBefore + height + block.SpacingAfter;
            double groupWidth = GetRadioButtonGroupWidth(block);
            EnsureFixedFlowBlockFits("Radio button group", groupWidth, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) {
                NewPage();
                spacingBefore = 0D;
            }

            if (spacingBefore > 0) {
                y -= spacingBefore;
            }

            double x = GetAlignedObjectX(containerX, containerWidth, groupWidth, block.Align);
            currentPage!.FormFields.Add(new FormFieldAnnotation {
                X1 = x,
                Y1 = y - height,
                X2 = x + block.Size,
                Y2 = y,
                Kind = FormFieldAnnotationKind.RadioButtonGroup,
                Name = block.Name,
                Value = block.Value,
                Options = block.Options,
                ButtonSize = block.Size,
                ButtonGap = block.Gap,
                Style = block.Style
            });
            RenderRadioButtonLabels(block, x, y);
            pageDirty = true;
            y -= height + block.SpacingAfter;
        }

        private static string GetFormFieldBlockName(IPdfBlock block) {
            if (block is TextFieldBlock) {
                return "Text field";
            }

            if (block is CheckBoxBlock) {
                return "Check box";
            }

            if (block is RadioButtonGroupBlock) {
                return "Radio button group";
            }

            return "Choice field";
        }

        private double GetFormFieldWidth(IPdfBlock block) {
            if (block is TextFieldBlock textField) {
                return textField.Width;
            }

            if (block is CheckBoxBlock checkBox) {
                return checkBox.Size;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                return GetRadioButtonGroupWidth(radioButtonGroup);
            }

            return ((ChoiceFieldBlock)block).Width;
        }

        private static double GetFormFieldHeight(IPdfBlock block) {
            if (block is TextFieldBlock textField) {
                return textField.Height;
            }

            if (block is CheckBoxBlock checkBox) {
                return checkBox.Size;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                return radioButtonGroup.Height;
            }

            return ((ChoiceFieldBlock)block).Height;
        }

        private static double GetFormFieldSpacingBefore(IPdfBlock block) {
            if (block is TextFieldBlock textField) {
                return textField.SpacingBefore;
            }

            if (block is CheckBoxBlock checkBox) {
                return checkBox.SpacingBefore;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                return radioButtonGroup.SpacingBefore;
            }

            return ((ChoiceFieldBlock)block).SpacingBefore;
        }

        private static double GetFormFieldSpacingAfter(IPdfBlock block) {
            if (block is TextFieldBlock textField) {
                return textField.SpacingAfter;
            }

            if (block is CheckBoxBlock checkBox) {
                return checkBox.SpacingAfter;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                return radioButtonGroup.SpacingAfter;
            }

            return ((ChoiceFieldBlock)block).SpacingAfter;
        }

        private static PdfAlign GetFormFieldAlign(IPdfBlock block) {
            if (block is TextFieldBlock textField) {
                return textField.Align;
            }

            if (block is CheckBoxBlock checkBox) {
                return checkBox.Align;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                return radioButtonGroup.Align;
            }

            return ((ChoiceFieldBlock)block).Align;
        }

        private void AddFormFieldAnnotation(IPdfBlock block, double x, double topY) {
            if (block is TextFieldBlock textField) {
                currentPage!.FormFields.Add(new FormFieldAnnotation {
                    X1 = x,
                    Y1 = topY - textField.Height,
                    X2 = x + textField.Width,
                    Y2 = topY,
                    Kind = FormFieldAnnotationKind.Text,
                    Name = textField.Name,
                    Value = textField.Value,
                    FontSize = textField.FontSize,
                    Style = textField.Style
                });
                return;
            }

            if (block is CheckBoxBlock checkBox) {
                currentPage!.FormFields.Add(new FormFieldAnnotation {
                    X1 = x,
                    Y1 = topY - checkBox.Size,
                    X2 = x + checkBox.Size,
                    Y2 = topY,
                    Kind = FormFieldAnnotationKind.CheckBox,
                    Name = checkBox.Name,
                    Value = checkBox.IsChecked ? checkBox.CheckedValueName : "Off",
                    IsChecked = checkBox.IsChecked,
                    CheckedValueName = checkBox.CheckedValueName,
                    Style = checkBox.Style
                });
                return;
            }

            if (block is RadioButtonGroupBlock radioButtonGroup) {
                currentPage!.FormFields.Add(new FormFieldAnnotation {
                    X1 = x,
                    Y1 = topY - radioButtonGroup.Height,
                    X2 = x + radioButtonGroup.Size,
                    Y2 = topY,
                    Kind = FormFieldAnnotationKind.RadioButtonGroup,
                    Name = radioButtonGroup.Name,
                    Value = radioButtonGroup.Value,
                    Options = radioButtonGroup.Options,
                    ButtonSize = radioButtonGroup.Size,
                    ButtonGap = radioButtonGroup.Gap,
                    Style = radioButtonGroup.Style
                });
                RenderRadioButtonLabels(radioButtonGroup, x, topY);
                return;
            }

            ChoiceFieldBlock choice = (ChoiceFieldBlock)block;
            currentPage!.FormFields.Add(new FormFieldAnnotation {
                X1 = x,
                Y1 = topY - choice.Height,
                X2 = x + choice.Width,
                Y2 = topY,
                Kind = FormFieldAnnotationKind.Choice,
                Name = choice.Name,
                Value = choice.Value,
                Values = choice.Values,
                FontSize = choice.FontSize,
                Options = choice.Options,
                IsComboBox = choice.IsComboBox,
                AllowsMultipleSelection = choice.AllowsMultipleSelection,
                Style = choice.Style
            });
        }

        private double GetRadioButtonGroupLabelFontSize(RadioButtonGroupBlock block) =>
            System.Math.Min(System.Math.Max(8D, currentOpts.DefaultFontSize), System.Math.Max(8D, block.Size));

        private static double GetRadioButtonGroupLabelGap(RadioButtonGroupBlock block) =>
            System.Math.Max(4D, block.Size * 0.4D);

        private double GetRadioButtonGroupWidth(RadioButtonGroupBlock block) {
            PdfStandardFont font = ChooseNormal(currentOpts.DefaultFont);
            double fontSize = GetRadioButtonGroupLabelFontSize(block);
            double labelWidth = block.Options.Max(option => EstimateSimpleTextWidthForOptions(option, font, fontSize, currentOpts));
            return block.Size + GetRadioButtonGroupLabelGap(block) + labelWidth;
        }

        private void RenderRadioButtonLabels(RadioButtonGroupBlock block, double x, double topY) {
            PdfStandardFont font = ChooseNormal(currentOpts.DefaultFont);
            string fontResource = GetStandardFontResourceName(font, font);
            double fontSize = GetRadioButtonGroupLabelFontSize(block);
            double labelX = x + block.Size + GetRadioButtonGroupLabelGap(block);
            double ascender = GetAscenderForOptions(font, fontSize, currentOpts);
            double descender = GetDescenderForOptions(font, fontSize, currentOpts);
            double labelBaselineOffset = (block.Size - ascender - descender) / 2D + descender;

            for (int i = 0; i < block.Options.Count; i++) {
                double optionTop = topY - i * (block.Size + block.Gap);
                double baseline = optionTop - block.Size + labelBaselineOffset;
                AppendPageText(sb, block.Options[i], font, fontResource, fontSize, block.Style.TextColor, labelX, baseline, currentOpts);
            }
        }

    }
}

