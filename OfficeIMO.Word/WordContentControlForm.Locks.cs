using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Checks the same mapped controls used by form filling without applying values.
        /// Locks on ancestors protect their entire content; replacement operations also
        /// must preserve locked controls nested inside the content they remove.
        /// </summary>
        private IEnumerable<WordContentControlFormIssue> GetFormControlLockIssues(
            IReadOnlyDictionary<string, object?> values, WordContentControlFormKey keyMode) {
            foreach (var control in EnumerateFormLockTargets()) {
                if (!TryGetFormValueByKeys(values, GetFormKeys(keyMode, control.Tag, control.Alias), out _, out string? key)
                    || control.Element == null) {
                    continue;
                }

                bool contentLocked = IsFormContentLocked(control.Element)
                    || control.Element.Ancestors().Any(IsFormContentLocked);
                bool nestedControlLocked = control.ReplacesContent
                    && control.Element.Descendants().Any(IsFormControlLocked);
                if (contentLocked || nestedControlLocked) {
                    yield return new WordContentControlFormIssue(
                        WordContentControlFormIssueKind.LockedControl,
                        key,
                        control.ControlType,
                        $"The {control.ControlType} content control '{key}' cannot be filled because "
                            + (contentLocked ? "its content is locked." : "replacing its content would remove a locked nested control."));
                }
            }
        }

        private IEnumerable<(SdtElement? Element, string? Tag, string? Alias, string ControlType, bool ReplacesContent)> EnumerateFormLockTargets() {
            foreach (WordCheckBox control in CheckBoxes)
                yield return (control._sdtRun, control.Tag, control.Alias, "Checkbox", false);
            foreach (WordDatePicker control in DatePickers)
                yield return (control._sdtRun, control.Tag, control.Alias, "Date picker", false);
            foreach (WordDropDownList control in DropDownLists)
                yield return (control._sdtRun, control.Tag, control.Alias, "Dropdown list", false);
            foreach (WordComboBox control in ComboBoxes)
                yield return (control._sdtRun, control.Tag, control.Alias, "Combo box", false);
            foreach (WordPictureControl control in PictureControls)
                yield return (control._sdtRun, control.Tag, control.Alias, "Picture control", true);
            foreach (WordRepeatingSection control in RepeatingSections)
                yield return (control._sdtRun, control.Tag, control.Alias, "Repeating section", true);

            HashSet<SdtElement> specializedElements = GetSpecializedStructuredDocumentTagElements();
            foreach (WordStructuredDocumentTag control in StructuredDocumentTags) {
                if (!IsSpecializedStructuredDocumentTag(control, specializedElements)) {
                    yield return (control.SdtElement, control.Tag, control.Alias, "Structured document tag", false);
                }
            }
        }

        private static bool IsFormContentLocked(OpenXmlElement element) {
            string? value = GetFormControlLockValue(element);
            return value == "contentLocked" || value == "sdtContentLocked";
        }

        private static bool IsFormControlLocked(OpenXmlElement element) {
            string? value = GetFormControlLockValue(element);
            return value == "sdtLocked" || value == "contentLocked" || value == "sdtContentLocked";
        }

        // Children of an SDK-unknown repeatingSectionItem reload as unknown elements,
        // even when they use the standard WordprocessingML names. Inspect their qualified
        // names and attributes so protection survives saving and reopening the package.
        private static string? GetFormControlLockValue(OpenXmlElement element) {
            const string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            if (element.LocalName != "sdt" || element.NamespaceUri != wordNamespace) return null;
            OpenXmlElement? properties = element.ChildElements.FirstOrDefault(child => child.LocalName == "sdtPr" && child.NamespaceUri == wordNamespace);
            OpenXmlElement? locking = properties?.ChildElements.FirstOrDefault(child => child.LocalName == "lock" && child.NamespaceUri == wordNamespace);
            return locking?.GetAttribute("val", wordNamespace).Value;
        }
    }
}
