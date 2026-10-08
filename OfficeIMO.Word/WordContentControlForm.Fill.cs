using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Fills supported content controls from a form map keyed by tag or alias.
        /// </summary>
        /// <param name="values">Values to apply.</param>
        /// <param name="keyMode">Controls which metadata field is used as the map key.</param>
        /// <returns>The number of controls updated.</returns>
        /// <remarks>
        /// Content locks are checked for every matched control before any values are applied.
        /// A lock that only prevents deleting the control still permits filling its content.
        /// </remarks>
        /// <exception cref="InvalidOperationException">A supplied value targets locked content or would remove a locked nested control.</exception>
        public int FillContentControlValues(IReadOnlyDictionary<string, object?> values, WordContentControlFormKey keyMode = WordContentControlFormKey.TagThenAlias) {
            if (values == null) throw new ArgumentNullException(nameof(values));

            List<WordContentControlFormIssue> lockIssues = GetFormControlLockIssues(values, keyMode).ToList();
            if (lockIssues.Count > 0) {
                throw new InvalidOperationException(string.Join(Environment.NewLine, lockIssues.Select(issue => issue.Message)));
            }

            int updated = 0;
            foreach (WordCheckBox checkBox in CheckBoxes) {
                if (TryGetFormValue(values, keyMode, checkBox.Tag, checkBox.Alias, out object? value)
                    && TryConvertFormBoolean(value, out bool boolValue)) {
                    checkBox.IsChecked = boolValue;
                    updated++;
                }
            }

            foreach (WordDatePicker datePicker in DatePickers) {
                if (TryGetFormValue(values, keyMode, datePicker.Tag, datePicker.Alias, out object? value)
                    && TryConvertFormDate(value, out DateTime? dateValue)) {
                    datePicker.Date = dateValue;
                    updated++;
                }
            }

            foreach (WordDropDownList dropDownList in DropDownLists) {
                if (TryGetFormValue(values, keyMode, dropDownList.Tag, dropDownList.Alias, out object? value)) {
                    dropDownList.SelectedValue = ConvertFormValueToString(value);
                    updated++;
                }
            }

            foreach (WordComboBox comboBox in ComboBoxes) {
                if (TryGetFormValue(values, keyMode, comboBox.Tag, comboBox.Alias, out object? value)) {
                    comboBox.SelectedValue = ConvertFormValueToString(value);
                    updated++;
                }
            }

            foreach (WordPictureControl pictureControl in PictureControls) {
                if (TryGetFormValue(values, keyMode, pictureControl.Tag, pictureControl.Alias, out object? value)
                    && TryApplyPictureFormValue(pictureControl, value)) {
                    updated++;
                }
            }

            foreach (WordRepeatingSection repeatingSection in RepeatingSections) {
                if (TryGetFormValue(values, keyMode, repeatingSection.Tag, repeatingSection.Alias, out object? value)
                    && TryConvertRepeatingSectionValue(value, out IReadOnlyList<string> itemValues)) {
                    repeatingSection.SetTextItems(itemValues);
                    updated++;
                }
            }

            HashSet<SdtElement> specializedElements = GetSpecializedStructuredDocumentTagElements();

            foreach (WordStructuredDocumentTag structuredDocumentTag in StructuredDocumentTags) {
                if (IsSpecializedStructuredDocumentTag(structuredDocumentTag, specializedElements)) {
                    continue;
                }

                if (TryGetFormValue(values, keyMode, structuredDocumentTag.Tag, structuredDocumentTag.Alias, out object? value)) {
                    structuredDocumentTag.Text = ConvertFormValueToString(value) ?? string.Empty;
                    updated++;
                }
            }

            return updated;
        }
    }
}
