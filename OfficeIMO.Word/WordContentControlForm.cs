using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;

namespace OfficeIMO.Word {
    /// <summary>
    /// Selects which content-control metadata should be used as the form field key.
    /// </summary>
    public enum WordContentControlFormKey {
        /// <summary>Use the content-control tag.</summary>
        Tag,
        /// <summary>Use the content-control alias.</summary>
        Alias,
        /// <summary>Use the tag when present, otherwise the alias.</summary>
        TagThenAlias,
        /// <summary>Use the alias when present, otherwise the tag.</summary>
        AliasThenTag
    }

    /// <summary>
    /// Describes a content-control form-map validation issue.
    /// </summary>
    public enum WordContentControlFormIssueKind {
        /// <summary>A content control has no tag or alias that can be used as a form key.</summary>
        UnmappedControl,
        /// <summary>More than one content control maps to the same form key.</summary>
        DuplicateKey,
        /// <summary>A mapped content control has no supplied value.</summary>
        MissingValue,
        /// <summary>A supplied value does not match any mapped content control.</summary>
        UnusedValue,
        /// <summary>A checkbox value cannot be converted to a Boolean value.</summary>
        InvalidBoolean,
        /// <summary>A date picker value cannot be converted to a date.</summary>
        InvalidDate,
        /// <summary>A dropdown or combobox value is not one of the configured choices.</summary>
        InvalidChoice,
        /// <summary>A picture-control value cannot be converted to image content.</summary>
        InvalidImage,
        /// <summary>A repeating-section value cannot be converted to one or more item values.</summary>
        InvalidRepeatingSection
    }

    /// <summary>
    /// Represents a picture content-control value that can be extracted from or supplied to a form map.
    /// </summary>
    public sealed class WordContentControlPictureValue {
        private WordContentControlPictureValue(string? filePath, byte[]? bytes, string fileName, Uri? externalUri, string? relationshipId) {
            FilePath = filePath;
            Bytes = bytes;
            FileName = string.IsNullOrWhiteSpace(fileName) ? "image.bin" : fileName;
            ExternalUri = externalUri;
            RelationshipId = relationshipId;
        }

        /// <summary>Optional source file path for a replacement image.</summary>
        public string? FilePath { get; }

        /// <summary>Embedded image bytes, when the value was extracted from package content or supplied from memory.</summary>
        public byte[]? Bytes { get; }

        /// <summary>File name used to infer image type for byte-backed replacements.</summary>
        public string FileName { get; }

        /// <summary>External image URI when the extracted picture is linked instead of embedded.</summary>
        public Uri? ExternalUri { get; }

        /// <summary>Relationship id for an extracted external image, when available.</summary>
        public string? RelationshipId { get; }

        /// <summary>True when the value describes an externally linked image.</summary>
        public bool IsExternal => ExternalUri != null;

        /// <summary>Creates a picture value from a local image path.</summary>
        public static WordContentControlPictureValue FromFile(string filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
            return new WordContentControlPictureValue(filePath, null, Path.GetFileName(filePath), null, null);
        }

        /// <summary>Creates a picture value from image bytes.</summary>
        public static WordContentControlPictureValue FromBytes(byte[] bytes, string fileName) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentNullException(nameof(fileName));
            return new WordContentControlPictureValue(null, bytes.ToArray(), fileName, null, null);
        }

        internal static WordContentControlPictureValue FromExternalImage(Uri? externalUri, string? fileName, string? relationshipId) {
            return new WordContentControlPictureValue(null, null, fileName ?? "external-image", externalUri, relationshipId);
        }
    }

    /// <summary>
    /// Represents one content-control form-map validation issue.
    /// </summary>
    public sealed class WordContentControlFormIssue {
        internal WordContentControlFormIssue(WordContentControlFormIssueKind kind, string? key, string controlType, string message) {
            Kind = kind;
            Key = key;
            ControlType = controlType;
            Message = message;
        }

        /// <summary>Issue category.</summary>
        public WordContentControlFormIssueKind Kind { get; }

        /// <summary>Form-map key related to the issue, when one is available.</summary>
        public string? Key { get; }

        /// <summary>Content-control type related to the issue.</summary>
        public string ControlType { get; }

        /// <summary>Human-readable issue text.</summary>
        public string Message { get; }
    }

    /// <summary>
    /// Describes the content-control form keys and validation issues found for a supplied form map.
    /// </summary>
    public sealed partial class WordContentControlFormValidationResult {
        internal WordContentControlFormValidationResult(IEnumerable<string> expectedKeys, IEnumerable<string> suppliedKeys, IEnumerable<WordContentControlFormIssue> issues) {
            ExpectedKeys = expectedKeys.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(key => key, StringComparer.OrdinalIgnoreCase).ToList();
            SuppliedKeys = suppliedKeys.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(key => key, StringComparer.OrdinalIgnoreCase).ToList();
            Issues = issues.ToList();
        }

        /// <summary>Unique content-control keys found in the document.</summary>
        public IReadOnlyList<string> ExpectedKeys { get; }

        /// <summary>Unique keys supplied by the caller.</summary>
        public IReadOnlyList<string> SuppliedKeys { get; }

        /// <summary>Validation issues found before filling content controls.</summary>
        public IReadOnlyList<WordContentControlFormIssue> Issues { get; }

        /// <summary>True when the supplied values can be applied without validation issues.</summary>
        public bool IsValid => Issues.Count == 0;

        /// <summary>
        /// Throws when validation issues were found, otherwise returns this result.
        /// </summary>
        public WordContentControlFormValidationResult EnsureValid() {
            if (!IsValid) {
                throw new InvalidOperationException(string.Join(Environment.NewLine, Issues.Select(issue => issue.Message)));
            }

            return this;
        }
    }

    public partial class WordDocument {
        /// <summary>
        /// Extracts values from supported content controls into a form map keyed by tag or alias.
        /// </summary>
        /// <param name="keyMode">Controls which metadata field is used as the map key.</param>
        /// <returns>Extracted form values.</returns>
        public Dictionary<string, object?> ExtractContentControlValues(WordContentControlFormKey keyMode = WordContentControlFormKey.TagThenAlias) {
            var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

            foreach (WordCheckBox checkBox in CheckBoxes) {
                AddFormValue(values, keyMode, checkBox.Tag, checkBox.Alias, checkBox.IsChecked);
            }

            foreach (WordDatePicker datePicker in DatePickers) {
                AddFormValue(values, keyMode, datePicker.Tag, datePicker.Alias, datePicker.Date);
            }

            foreach (WordDropDownList dropDownList in DropDownLists) {
                AddFormValue(values, keyMode, dropDownList.Tag, dropDownList.Alias, dropDownList.SelectedValue);
            }

            foreach (WordComboBox comboBox in ComboBoxes) {
                AddFormValue(values, keyMode, comboBox.Tag, comboBox.Alias, comboBox.SelectedValue);
            }

            foreach (WordPictureControl pictureControl in PictureControls) {
                AddFormValue(values, keyMode, pictureControl.Tag, pictureControl.Alias, pictureControl.ExtractValue());
            }

            foreach (WordRepeatingSection repeatingSection in RepeatingSections) {
                AddFormValue(values, keyMode, repeatingSection.Tag, repeatingSection.Alias, repeatingSection.ExtractValue());
            }

            foreach (WordStructuredDocumentTag structuredDocumentTag in StructuredDocumentTags) {
                if (IsSpecializedStructuredDocumentTag(structuredDocumentTag)) {
                    continue;
                }

                AddFormValue(values, keyMode, structuredDocumentTag.Tag, structuredDocumentTag.Alias, structuredDocumentTag.Text);
            }

            return values;
        }

        /// <summary>
        /// Fills supported content controls from a form map keyed by tag or alias.
        /// </summary>
        /// <param name="values">Values to apply.</param>
        /// <param name="keyMode">Controls which metadata field is used as the map key.</param>
        /// <returns>The number of controls updated.</returns>
        public int FillContentControlValues(IReadOnlyDictionary<string, object?> values, WordContentControlFormKey keyMode = WordContentControlFormKey.TagThenAlias) {
            if (values == null) throw new ArgumentNullException(nameof(values));

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

        /// <summary>
        /// Validates a content-control form map before values are applied to the document.
        /// </summary>
        /// <param name="values">Values to validate.</param>
        /// <param name="keyMode">Controls which metadata field is used as the map key.</param>
        /// <param name="requireAllControls">When true, report mapped controls that have no supplied value.</param>
        /// <param name="allowUnusedValues">When false, report supplied values that do not match a mapped control.</param>
        /// <returns>Validation result containing expected keys, supplied keys, and issues.</returns>
        public WordContentControlFormValidationResult ValidateContentControlValues(
            IReadOnlyDictionary<string, object?> values,
            WordContentControlFormKey keyMode = WordContentControlFormKey.TagThenAlias,
            bool requireAllControls = true,
            bool allowUnusedValues = false) {
            if (values == null) throw new ArgumentNullException(nameof(values));

            var issues = new List<WordContentControlFormIssue>();
            var expectedKeys = new List<string>();
            var matchedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var keyOwners = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (WordCheckBox checkBox in CheckBoxes) {
                ValidateFormControl(
                    issues,
                    expectedKeys,
                    matchedKeys,
                    keyOwners,
                    values,
                    keyMode,
                    checkBox.Tag,
                    checkBox.Alias,
                    "Checkbox",
                    requireAllControls,
                    value => TryConvertFormBoolean(value, out _),
                    WordContentControlFormIssueKind.InvalidBoolean,
                    "cannot be converted to a Boolean value.");
            }

            foreach (WordDatePicker datePicker in DatePickers) {
                ValidateFormControl(
                    issues,
                    expectedKeys,
                    matchedKeys,
                    keyOwners,
                    values,
                    keyMode,
                    datePicker.Tag,
                    datePicker.Alias,
                    "Date picker",
                    requireAllControls,
                    value => TryConvertFormDate(value, out _),
                    WordContentControlFormIssueKind.InvalidDate,
                    "cannot be converted to a date.");
            }

            foreach (WordDropDownList dropDownList in DropDownLists) {
                ValidateFormControl(
                    issues,
                    expectedKeys,
                    matchedKeys,
                    keyOwners,
                    values,
                    keyMode,
                    dropDownList.Tag,
                    dropDownList.Alias,
                    "Dropdown list",
                    requireAllControls,
                    value => IsAllowedChoice(EnumerateDropDownListChoices(dropDownList), value),
                    WordContentControlFormIssueKind.InvalidChoice,
                    "is not one of the configured dropdown choices.");
            }

            foreach (WordComboBox comboBox in ComboBoxes) {
                ValidateFormControl(
                    issues,
                    expectedKeys,
                    matchedKeys,
                    keyOwners,
                    values,
                    keyMode,
                    comboBox.Tag,
                    comboBox.Alias,
                    "Combo box",
                    requireAllControls,
                    value => IsAllowedChoice(EnumerateComboBoxChoices(comboBox), value),
                    WordContentControlFormIssueKind.InvalidChoice,
                    "is not one of the configured combo box choices.");
            }

            foreach (WordPictureControl pictureControl in PictureControls) {
                ValidateFormControl(
                    issues,
                    expectedKeys,
                    matchedKeys,
                    keyOwners,
                    values,
                    keyMode,
                    pictureControl.Tag,
                    pictureControl.Alias,
                    "Picture control",
                    requireAllControls,
                    IsValidPictureFormValue,
                    WordContentControlFormIssueKind.InvalidImage,
                    "cannot be converted to image content.");
            }

            foreach (WordRepeatingSection repeatingSection in RepeatingSections) {
                ValidateFormControl(
                    issues,
                    expectedKeys,
                    matchedKeys,
                    keyOwners,
                    values,
                    keyMode,
                    repeatingSection.Tag,
                    repeatingSection.Alias,
                    "Repeating section",
                    requireAllControls,
                    value => TryConvertRepeatingSectionValue(value, out _),
                    WordContentControlFormIssueKind.InvalidRepeatingSection,
                    "cannot be converted to repeating-section item values.");
            }

            HashSet<SdtElement> specializedElements = GetSpecializedStructuredDocumentTagElements();

            foreach (WordStructuredDocumentTag structuredDocumentTag in StructuredDocumentTags) {
                if (IsSpecializedStructuredDocumentTag(structuredDocumentTag, specializedElements)) {
                    continue;
                }

                var keys = GetFormKeys(keyMode, structuredDocumentTag.Tag, structuredDocumentTag.Alias)
                    .ToList();
                if (keys.Count == 0) {
                    if (requireAllControls) {
                        issues.Add(new WordContentControlFormIssue(
                            WordContentControlFormIssueKind.UnmappedControl,
                            string.Empty,
                            "Structured document tag",
                            "Structured document tag has no tag or alias that can be used as a form key."));
                    }

                    continue;
                }

                ValidateFormControl(
                    issues,
                    expectedKeys,
                    matchedKeys,
                    keyOwners,
                    values,
                    keys,
                    "Structured document tag",
                    requireAllControls,
                    value => true,
                    WordContentControlFormIssueKind.InvalidChoice,
                    string.Empty);
            }

            if (!allowUnusedValues) {
                foreach (string suppliedKey in values.Keys) {
                    if (!matchedKeys.Contains(suppliedKey)) {
                        issues.Add(new WordContentControlFormIssue(
                            WordContentControlFormIssueKind.UnusedValue,
                            suppliedKey,
                            "Form map",
                            $"The supplied form value '{suppliedKey}' does not match any content-control key."));
                    }
                }
            }

            return new WordContentControlFormValidationResult(expectedKeys, values.Keys, issues);
        }

        private static void AddFormValue(IDictionary<string, object?> values, WordContentControlFormKey keyMode, string? tag, string? alias, object? value) {
            foreach (string key in GetFormKeys(keyMode, tag, alias)) {
                if (!values.ContainsKey(key)) {
                    values.Add(key, value);
                    return;
                }
            }
        }

        private bool IsSpecializedStructuredDocumentTag(WordStructuredDocumentTag structuredDocumentTag) {
            return IsSpecializedStructuredDocumentTag(structuredDocumentTag, GetSpecializedStructuredDocumentTagElements());
        }

        private HashSet<SdtElement> GetSpecializedStructuredDocumentTagElements() {
            var elements = new HashSet<SdtElement>();
            foreach (WordCheckBox checkBox in CheckBoxes) elements.Add(checkBox._sdtRun);
            foreach (WordDatePicker datePicker in DatePickers) elements.Add(datePicker._sdtRun);
            foreach (WordDropDownList dropDownList in DropDownLists) elements.Add(dropDownList._sdtRun);
            foreach (WordComboBox comboBox in ComboBoxes) elements.Add(comboBox._sdtRun);
            foreach (WordPictureControl pictureControl in PictureControls) elements.Add(pictureControl._sdtRun);
            foreach (WordRepeatingSection repeatingSection in RepeatingSections) elements.Add(repeatingSection._sdtRun);
            return elements;
        }

        private static bool IsSpecializedStructuredDocumentTag(WordStructuredDocumentTag structuredDocumentTag, ISet<SdtElement> specializedElements) {
            SdtElement? element = structuredDocumentTag.SdtElement;
            return element != null && specializedElements.Contains(element);
        }

        private static bool TryGetFormValue(
            IReadOnlyDictionary<string, object?> values,
            WordContentControlFormKey keyMode,
            string? tag,
            string? alias,
            out object? value,
            ISet<string>? excludedKeys = null) {
            foreach (string key in GetFormKeys(keyMode, tag, alias)) {
                if (excludedKeys != null && excludedKeys.Contains(key)) {
                    continue;
                }

                if (TryGetFormValueByKey(values, key, out value)) {
                    return true;
                }
            }

            value = null;
            return false;
        }

        private static void ValidateFormControl(
            ICollection<WordContentControlFormIssue> issues,
            ICollection<string> expectedKeys,
            ISet<string> matchedKeys,
            IDictionary<string, string> keyOwners,
            IReadOnlyDictionary<string, object?> values,
            WordContentControlFormKey keyMode,
            string? tag,
            string? alias,
            string controlType,
            bool requireAllControls,
            Func<object?, bool> isValueValid,
            WordContentControlFormIssueKind invalidKind,
            string invalidMessage) {
            ValidateFormControl(
                issues,
                expectedKeys,
                matchedKeys,
                keyOwners,
                values,
                GetFormKeys(keyMode, tag, alias).ToList(),
                controlType,
                requireAllControls,
                isValueValid,
                invalidKind,
                invalidMessage);
        }

        private static void ValidateFormControl(
            ICollection<WordContentControlFormIssue> issues,
            ICollection<string> expectedKeys,
            ISet<string> matchedKeys,
            IDictionary<string, string> keyOwners,
            IReadOnlyDictionary<string, object?> values,
            IReadOnlyList<string> keys,
            string controlType,
            bool requireAllControls,
            Func<object?, bool> isValueValid,
            WordContentControlFormIssueKind invalidKind,
            string invalidMessage) {
            var bindableKeys = keys
                .Where(key => !string.IsNullOrWhiteSpace(key))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (bindableKeys.Count == 0) {
                issues.Add(new WordContentControlFormIssue(
                    WordContentControlFormIssueKind.UnmappedControl,
                    null,
                    controlType,
                    $"A {controlType} content control does not have a tag or alias for form-map binding."));
                return;
            }

            string primaryKey = bindableKeys[0];
            expectedKeys.Add(primaryKey);
            foreach (string key in bindableKeys) {
                if (keyOwners.TryGetValue(key, out string? existingOwner)) {
                    issues.Add(new WordContentControlFormIssue(
                        WordContentControlFormIssueKind.DuplicateKey,
                        key,
                        controlType,
                        $"The form key '{key}' is used by multiple content controls ({existingOwner}, {controlType})."));
                } else {
                    keyOwners.Add(key, controlType);
                }
            }

            if (TryGetFormValueByKeys(values, bindableKeys, out object? value, out string? matchedKey)) {
                matchedKeys.Add(matchedKey!);
                if (!isValueValid(value)) {
                    issues.Add(new WordContentControlFormIssue(
                        invalidKind,
                        matchedKey,
                        controlType,
                        $"The supplied value for '{matchedKey}' {invalidMessage}"));
                }

                return;
            }

            if (requireAllControls) {
                issues.Add(new WordContentControlFormIssue(
                    WordContentControlFormIssueKind.MissingValue,
                    primaryKey,
                    controlType,
                    $"The {controlType} content control '{primaryKey}' has no supplied form value."));
            }
        }

        private static bool TryGetFormValueByKeys(IReadOnlyDictionary<string, object?> values, IEnumerable<string> keys, out object? value, out string? matchedKey) {
            foreach (string key in keys) {
                if (TryGetFormValueByKey(values, key, out value)) {
                    matchedKey = key;
                    return true;
                }
            }

            value = null;
            matchedKey = null;
            return false;
        }

        private static bool TryGetFormValueByKey(IReadOnlyDictionary<string, object?> values, string key, out object? value) {
            if (values.TryGetValue(key, out value)) {
                return true;
            }

            foreach (KeyValuePair<string, object?> pair in values) {
                if (string.Equals(pair.Key, key, StringComparison.OrdinalIgnoreCase)) {
                    value = pair.Value;
                    return true;
                }
            }

            return false;
        }

        private static IEnumerable<string> GetFormKeys(WordContentControlFormKey keyMode, string? tag, string? alias) {
            if (keyMode == WordContentControlFormKey.Tag || keyMode == WordContentControlFormKey.TagThenAlias) {
                if (!string.IsNullOrWhiteSpace(tag)) yield return tag!;
            }

            if (keyMode == WordContentControlFormKey.Alias || keyMode == WordContentControlFormKey.TagThenAlias || keyMode == WordContentControlFormKey.AliasThenTag) {
                if (!string.IsNullOrWhiteSpace(alias)) yield return alias!;
            }

            if (keyMode == WordContentControlFormKey.AliasThenTag) {
                if (!string.IsNullOrWhiteSpace(tag)) yield return tag!;
            }
        }

        private static bool TryConvertFormBoolean(object? value, out bool result) {
            if (value is bool boolValue) {
                result = boolValue;
                return true;
            }

            string? text = ConvertFormValueToString(value);
            if (string.IsNullOrWhiteSpace(text)) {
                result = false;
                return false;
            }

            if (bool.TryParse(text, out result)) {
                return true;
            }

            if (string.Equals(text, "yes", StringComparison.OrdinalIgnoreCase) || text == "1") {
                result = true;
                return true;
            }

            if (string.Equals(text, "no", StringComparison.OrdinalIgnoreCase) || text == "0") {
                result = false;
                return true;
            }

            result = false;
            return false;
        }

        private static bool TryConvertFormDate(object? value, out DateTime? result) {
            if (value == null) {
                result = null;
                return true;
            }

            if (value is DateTime dateTime) {
                result = dateTime;
                return true;
            }

            if (value is DateTimeOffset dateTimeOffset) {
                result = dateTimeOffset.DateTime;
                return true;
            }

            string? text = ConvertFormValueToString(value);
            if (string.IsNullOrWhiteSpace(text)) {
                result = null;
                return true;
            }

            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime invariantDate)
                || DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out invariantDate)) {
                result = invariantDate;
                return true;
            }

            result = null;
            return false;
        }

        private static string? ConvertFormValueToString(object? value) {
            if (value == null) {
                return null;
            }

            return value is IFormattable formattable
                ? formattable.ToString(null, CultureInfo.InvariantCulture)
                : value.ToString();
        }

        private static bool IsAllowedChoice(IEnumerable<string> allowedValues, object? value) {
            string? text = ConvertFormValueToString(value);
            return string.IsNullOrEmpty(text)
                || allowedValues.Any(allowedValue => string.Equals(allowedValue, text, StringComparison.OrdinalIgnoreCase));
        }

        private static IEnumerable<string> EnumerateDropDownListChoices(WordDropDownList dropDownList) {
            SdtContentDropDownList? list = dropDownList._sdtRun.SdtProperties?.Elements<SdtContentDropDownList>().FirstOrDefault();
            return list == null
                ? dropDownList.Items
                : EnumerateListItemChoices(list.Elements<ListItem>());
        }

        private static IEnumerable<string> EnumerateComboBoxChoices(WordComboBox comboBox) {
            SdtContentComboBox? list = comboBox._sdtRun.SdtProperties?.Elements<SdtContentComboBox>().FirstOrDefault();
            return list == null
                ? comboBox.Items
                : EnumerateListItemChoices(list.Elements<ListItem>());
        }

        private static IEnumerable<string> EnumerateListItemChoices(IEnumerable<ListItem> items) {
            foreach (ListItem item in items) {
                string? value = item.Value?.Value;
                if (!string.IsNullOrEmpty(value)) {
                    yield return value!;
                }

                string? displayText = item.DisplayText?.Value;
                if (!string.IsNullOrEmpty(displayText) && !string.Equals(displayText, value, StringComparison.OrdinalIgnoreCase)) {
                    yield return displayText!;
                }
            }
        }

        private static bool IsValidPictureFormValue(object? value) {
            if (value is WordContentControlPictureValue pictureValue) {
                if (!string.IsNullOrWhiteSpace(pictureValue.FilePath)) {
                    return File.Exists(pictureValue.FilePath);
                }

                return pictureValue.Bytes != null
                    && pictureValue.Bytes.Length > 0
                    && !string.IsNullOrWhiteSpace(pictureValue.FileName)
                    && Path.HasExtension(pictureValue.FileName);
            }

            return false;
        }

        private static bool TryApplyPictureFormValue(WordPictureControl pictureControl, object? value) {
            if (value is WordContentControlPictureValue pictureValue) {
                string? pictureFilePath = pictureValue.FilePath;
                if (!string.IsNullOrWhiteSpace(pictureFilePath) && File.Exists(pictureFilePath)) {
                    pictureControl.SetImage(pictureFilePath!);
                    return true;
                }

                if (pictureValue.Bytes != null
                    && pictureValue.Bytes.Length > 0
                    && !string.IsNullOrWhiteSpace(pictureValue.FileName)) {
                    using var stream = new MemoryStream(pictureValue.Bytes);
                    pictureControl.SetImage(stream, pictureValue.FileName);
                    return true;
                }
            }

            return false;
        }

        private static bool TryConvertRepeatingSectionValue(object? value, out IReadOnlyList<string> items) {
            if (value == null) {
                items = Array.Empty<string>();
                return true;
            }

            if (value is string text) {
                items = new[] { text };
                return true;
            }

            if (value is IEnumerable<string> stringValues) {
                items = stringValues.Select(item => item ?? string.Empty).ToList();
                return true;
            }

            if (value is IEnumerable<object?> objectValues) {
                items = objectValues.Select(item => ConvertFormValueToString(item) ?? string.Empty).ToList();
                return true;
            }

            items = Array.Empty<string>();
            return false;
        }
    }
}
