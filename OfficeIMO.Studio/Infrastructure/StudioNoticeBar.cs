using System.Windows.Input;
using Avalonia;
using Avalonia.Automation;
using Avalonia.Automation.Peers;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>How urgent a notice is; it chooses the icon, accent and how assistive technology announces it.</summary>
public enum StudioNoticeSeverity {
    Information,
    Success,
    Warning,
    Error
}

/// <summary>
/// One full-width notice for state that stays until it is resolved: a failure, a restriction or a storage problem.
/// Every notice states what happened in plain text, may offer up to two next actions, and can be dismissed.
/// Finished operations use the bottom toast instead; field validation stays next to the field.
/// </summary>
public class StudioNoticeBar : TemplatedControl {
    public static readonly StyledProperty<StudioNoticeSeverity> SeverityProperty =
        AvaloniaProperty.Register<StudioNoticeBar, StudioNoticeSeverity>(nameof(Severity));

    public static readonly StyledProperty<string?> MessageProperty =
        AvaloniaProperty.Register<StudioNoticeBar, string?>(nameof(Message));

    public static readonly StyledProperty<string?> ActionTextProperty =
        AvaloniaProperty.Register<StudioNoticeBar, string?>(nameof(ActionText));

    public static readonly StyledProperty<ICommand?> ActionCommandProperty =
        AvaloniaProperty.Register<StudioNoticeBar, ICommand?>(nameof(ActionCommand));

    public static readonly StyledProperty<bool> IsActionVisibleProperty =
        AvaloniaProperty.Register<StudioNoticeBar, bool>(nameof(IsActionVisible), true);

    public static readonly StyledProperty<string?> SecondaryActionTextProperty =
        AvaloniaProperty.Register<StudioNoticeBar, string?>(nameof(SecondaryActionText));

    public static readonly StyledProperty<ICommand?> SecondaryActionCommandProperty =
        AvaloniaProperty.Register<StudioNoticeBar, ICommand?>(nameof(SecondaryActionCommand));

    public static readonly StyledProperty<bool> IsSecondaryActionVisibleProperty =
        AvaloniaProperty.Register<StudioNoticeBar, bool>(nameof(IsSecondaryActionVisible), true);

    public static readonly StyledProperty<ICommand?> DismissCommandProperty =
        AvaloniaProperty.Register<StudioNoticeBar, ICommand?>(nameof(DismissCommand));

    public static readonly DirectProperty<StudioNoticeBar, bool> HasActionProperty =
        AvaloniaProperty.RegisterDirect<StudioNoticeBar, bool>(nameof(HasAction), bar => bar.HasAction);

    public static readonly DirectProperty<StudioNoticeBar, bool> HasSecondaryActionProperty =
        AvaloniaProperty.RegisterDirect<StudioNoticeBar, bool>(nameof(HasSecondaryAction), bar => bar.HasSecondaryAction);

    public static readonly DirectProperty<StudioNoticeBar, bool> IsDismissibleProperty =
        AvaloniaProperty.RegisterDirect<StudioNoticeBar, bool>(nameof(IsDismissible), bar => bar.IsDismissible);

    private bool _hasAction;
    private bool _hasSecondaryAction;
    private bool _isDismissible;

    public StudioNoticeBar() => UpdateSeverity();

    public StudioNoticeSeverity Severity { get => GetValue(SeverityProperty); set => SetValue(SeverityProperty, value); }
    public string? Message { get => GetValue(MessageProperty); set => SetValue(MessageProperty, value); }
    public string? ActionText { get => GetValue(ActionTextProperty); set => SetValue(ActionTextProperty, value); }
    public ICommand? ActionCommand { get => GetValue(ActionCommandProperty); set => SetValue(ActionCommandProperty, value); }
    public bool IsActionVisible { get => GetValue(IsActionVisibleProperty); set => SetValue(IsActionVisibleProperty, value); }
    public string? SecondaryActionText { get => GetValue(SecondaryActionTextProperty); set => SetValue(SecondaryActionTextProperty, value); }
    public ICommand? SecondaryActionCommand { get => GetValue(SecondaryActionCommandProperty); set => SetValue(SecondaryActionCommandProperty, value); }
    public bool IsSecondaryActionVisible { get => GetValue(IsSecondaryActionVisibleProperty); set => SetValue(IsSecondaryActionVisibleProperty, value); }
    public ICommand? DismissCommand { get => GetValue(DismissCommandProperty); set => SetValue(DismissCommandProperty, value); }

    public bool HasAction { get => _hasAction; private set => SetAndRaise(HasActionProperty, ref _hasAction, value); }
    public bool HasSecondaryAction { get => _hasSecondaryAction; private set => SetAndRaise(HasSecondaryActionProperty, ref _hasSecondaryAction, value); }
    public bool IsDismissible { get => _isDismissible; private set => SetAndRaise(IsDismissibleProperty, ref _isDismissible, value); }

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change) {
        base.OnPropertyChanged(change);
        if (change.Property == SeverityProperty) UpdateSeverity();
        else if (change.Property == ActionTextProperty || change.Property == ActionCommandProperty || change.Property == IsActionVisibleProperty)
            HasAction = IsActionVisible && ActionCommand is not null && !string.IsNullOrWhiteSpace(ActionText);
        else if (change.Property == SecondaryActionTextProperty || change.Property == SecondaryActionCommandProperty || change.Property == IsSecondaryActionVisibleProperty)
            HasSecondaryAction = IsSecondaryActionVisible && SecondaryActionCommand is not null && !string.IsNullOrWhiteSpace(SecondaryActionText);
        else if (change.Property == DismissCommandProperty)
            IsDismissible = DismissCommand is not null;
        else if (change.Property == MessageProperty)
            AutomationProperties.SetName(this, Message);
    }

    private void UpdateSeverity() {
        PseudoClasses.Set(":information", Severity == StudioNoticeSeverity.Information);
        PseudoClasses.Set(":success", Severity == StudioNoticeSeverity.Success);
        PseudoClasses.Set(":warning", Severity == StudioNoticeSeverity.Warning);
        PseudoClasses.Set(":error", Severity == StudioNoticeSeverity.Error);
        // Failures interrupt a screen reader; restrictions and confirmations wait for a pause.
        AutomationProperties.SetLiveSetting(this, Severity == StudioNoticeSeverity.Error ? AutomationLiveSetting.Assertive : AutomationLiveSetting.Polite);
    }

    protected override AutomationPeer OnCreateAutomationPeer() => new StudioNoticeBarAutomationPeer(this);

    private sealed class StudioNoticeBarAutomationPeer(StudioNoticeBar owner) : ControlAutomationPeer(owner) {
        protected override AutomationControlType GetAutomationControlTypeCore() => AutomationControlType.StatusBar;
        protected override bool IsContentElementCore() => true;
        protected override bool IsControlElementCore() => true;
    }
}

/// <summary>Localized text used by the notice bar template.</summary>
public static class StudioNoticeBarText {
    public static string Dismiss => Localization.StudioLocalization.Current.Get("MainWindow.Dismiss");
}
