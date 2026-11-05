using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Models.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Gets typed native content controls filtered by family, alias, tag, rendered text or checkbox state.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeWordContentControl", DefaultParameterSetName = "Document")]
[OutputType(typeof(WordStructuredDocumentTag), typeof(WordCheckBox), typeof(WordComboBox), typeof(WordDatePicker),
    typeof(WordDropDownList), typeof(WordPictureControl), typeof(WordRepeatingSection))]
public sealed class GetOfficeWordContentControlCommand : PSCmdlet {
    /// <summary>Word files with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Literal Word files, including FileInfo.FullName.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Caller-owned document; omission reads the enclosing Word document.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public WordDocument? Document { get; set; }
    /// <summary>Control families; omission includes all seven native families.</summary>
    [Parameter]
    public WordContentControlType[]? Type { get; set; }
    /// <summary>Wildcard control aliases.</summary>
    [Parameter, SupportsWildcards]
    public string[]? Alias { get; set; }
    /// <summary>Wildcard control tags.</summary>
    [Parameter, SupportsWildcards]
    public string[]? Tag { get; set; }
    /// <summary>Wildcard native rendered text.</summary>
    [Parameter, SupportsWildcards]
    public string[]? Text { get; set; }
    /// <summary>Checked or unchecked; implies the checkbox family.</summary>
    [Parameter]
    public WordCheckBoxState? State { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var selected = new HashSet<WordContentControlType>(Type ?? (WordContentControlType[])Enum.GetValues(typeof(WordContentControlType)));
        if (State.HasValue) {
            if (Type != null && (selected.Count != 1 || !selected.Contains(WordContentControlType.CheckBox)))
                throw new PSArgumentException("State selects CheckBox controls; omit Type or select only CheckBox.");
            selected.Clear(); selected.Add(WordContentControlType.CheckBox);
        }
        var alias = WordFilterHelpers.BuildPatterns(Alias);
        var tag = WordFilterHelpers.BuildPatterns(Tag);
        var text = WordFilterHelpers.BuildPatterns(Text);
        void Write<T>(IEnumerable<T> controls, Func<T, string?> getAlias, Func<T, string?> getTag, Func<T, string?> getText) {
            foreach (var control in controls) {
                if (!WordFilterHelpers.Matches(getAlias(control), alias) || !WordFilterHelpers.Matches(getTag(control), tag) ||
                    !WordFilterHelpers.Matches(getText(control), text)) continue;
                if (State.HasValue && control is WordCheckBox checkbox && checkbox.IsChecked != (State == WordCheckBoxState.Checked)) continue;
                WriteObject(control);
            }
        }
        WordQuerySource.Visit(this, ParameterSetName, Path, LiteralPath, Document, document => {
            if (selected.Contains(WordContentControlType.Text)) Write(document.TextContentControls, control => control.Alias, control => control.Tag, control => control.Text);
            if (selected.Contains(WordContentControlType.CheckBox)) Write(document.CheckBoxes, control => control.Alias, control => control.Tag, control => control.Text);
            if (selected.Contains(WordContentControlType.ComboBox)) Write(document.ComboBoxes, control => control.Alias, control => control.Tag, control => control.Text);
            if (selected.Contains(WordContentControlType.DatePicker)) Write(document.DatePickers, control => control.Alias, control => control.Tag, control => control.Text);
            if (selected.Contains(WordContentControlType.DropDownList)) Write(document.DropDownLists, control => control.Alias, control => control.Tag, control => control.Text);
            if (selected.Contains(WordContentControlType.Picture)) Write(document.PictureControls, control => control.Alias, control => control.Tag, control => control.Text);
            if (selected.Contains(WordContentControlType.RepeatingSection)) Write(document.RepeatingSections, control => control.Alias, control => control.Tag, control => control.Text);
        });
    }
}
