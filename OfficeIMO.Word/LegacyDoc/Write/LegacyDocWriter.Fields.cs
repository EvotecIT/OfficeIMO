using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const string PageFieldInstruction = " PAGE   \\* MERGEFORMAT ";

        private static void AppendSupportedPageNumberField(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableFormatting formatting) {
            AppendFormattedText(text, runs, LegacyDocField.Begin.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
            AppendFormattedText(text, runs, PageFieldInstruction, LegacyDocWritableFormatting.Plain);
            AppendFormattedText(text, runs, LegacyDocField.Separator.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
            AppendFormattedText(text, runs, "1", formatting);
            AppendFormattedText(text, runs, LegacyDocField.End.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
        }

        private static void AppendSupportedPageNumberFieldFromSimpleField(StringBuilder text, List<LegacyDocWritableRun> runs, SimpleField field, LegacyDocWritableFormatting inheritedFormatting) {
            if (!IsSupportedPageNumberFieldInstruction(field.Instruction?.Value)) {
                throw new NotSupportedException("Native DOC saving currently supports only PAGE simple fields. Other field types are not supported yet.");
            }

            LegacyDocWritableFormatting formatting = ReadSimpleFieldResultFormatting(field)
                .WithInheritedFormatting(inheritedFormatting);
            AppendSupportedPageNumberField(text, runs, formatting);
        }

        private static bool IsSupportedPageNumberFieldInstruction(string? instruction) {
            if (string.IsNullOrWhiteSpace(instruction)) {
                return false;
            }

            string trimmed = instruction!.Trim();
            if (!trimmed.StartsWith("PAGE", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            return trimmed.Length == "PAGE".Length
                || char.IsWhiteSpace(trimmed["PAGE".Length]);
        }

        private static LegacyDocWritableFormatting ReadSimpleFieldResultFormatting(SimpleField field) {
            LegacyDocWritableFormatting? formatting = null;
            foreach (Run run in field.Elements<Run>()) {
                EnsureSimplePageFieldResultRun(run);
                LegacyDocWritableFormatting runFormatting = ReadSupportedRunFormatting(run.RunProperties);
                formatting ??= runFormatting;
                if (!formatting.Value.Equals(runFormatting)) {
                    throw new NotSupportedException("Native DOC saving supports PAGE simple fields only when their display runs use one formatting set.");
                }
            }

            return formatting ?? LegacyDocWritableFormatting.Plain;
        }

        private static void EnsureSimplePageFieldResultRun(Run run) {
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                    case LastRenderedPageBreak:
                    case Text:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports PAGE simple fields only when their display result contains text runs. Unsupported PAGE field result element: {child.LocalName}.");
                }
            }
        }
    }
}
