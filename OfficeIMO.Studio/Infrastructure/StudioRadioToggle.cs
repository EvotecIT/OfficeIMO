using Avalonia.Controls.Primitives;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>
/// A toggle that represents one choice of a view-model owned selection (mode, tool, navigation area).
/// Clicking the active choice keeps it checked; only the bound state moves the selection.
/// </summary>
public sealed class StudioRadioToggle : ToggleButton {
    protected override Type StyleKeyOverride => typeof(ToggleButton);

    protected override void Toggle() {
        if (IsChecked != true) base.Toggle();
    }
}
