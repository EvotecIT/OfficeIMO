using AngleSharp.Dom;
using AngleSharp.Dom.Events;
using AngleSharp.Html.Dom;
using AngleSharp.Io;
using Jint;
using Jint.Native;
using Jint.Native.Object;
using Jint.Runtime.Descriptors;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed partial class RuntimeAutomation {
    private readonly Engine _engine;
    private readonly HtmlScriptRequest _options;
    private readonly HashSet<IHtmlFormElement> _resetting = new(ReferenceEqualityComparer.Instance);

    private HtmlAutomationResult ApplyFormDefault(IHtmlFormElement form, IHtmlElement submitter, string type) {
        RuntimeFormActionOutcome outcome = type == "reset"
            ? ResetForm(form)
            : SubmitForm(form, submitter, dispatchSubmitEvent: true, validate: true);
        return outcome.Status == HtmlAutomationStatus.Success
            ? Success(submitter)
            : Failure(outcome.Status, outcome.Message!, 1, Inspect(submitter));
    }

    internal RuntimeFormActionOutcome ResetForm(IHtmlFormElement form) {
        if (!_resetting.Add(form)) return RuntimeFormActionOutcome.Success;
        try {
            var reset = new Event("reset", bubbles: true, cancelable: true);
            form.Dispatch(reset);
            if (!reset.IsDefaultPrevented) form.Reset();
            return RuntimeFormActionOutcome.Success;
        } finally {
            _resetting.Remove(form);
        }
    }

    internal RuntimeFormActionOutcome SubmitForm(
        IHtmlFormElement form,
        IHtmlElement? submitter,
        bool dispatchSubmitEvent,
        bool validate) {
        if (!RuntimeFocusController.IsConnected(form)) return RuntimeFormActionOutcome.Success;
        if (submitter != null && (!HtmlFormControlSemantics.IsSubmitter(submitter)
            || !ReferenceEquals(HtmlFormControlSemantics.ResolveFormOwner(submitter), form)))
            return RuntimeFormActionOutcome.Invalid("The submitter must be a submit button owned by this form.");
        if (validate && !form.NoValidate && submitter?.HasAttribute("formnovalidate") != true && !form.CheckValidity())
            return RuntimeFormActionOutcome.Success;
        if (dispatchSubmitEvent) {
            var submit = new Event("submit", bubbles: true, cancelable: true);
            ObjectInstance wrapped = JsValue.FromObject(_engine, submit).AsObject();
            wrapped.FastSetProperty("submitter", new PropertyDescriptor(
                submitter == null ? JsValue.Null : JsValue.FromObject(_engine, submitter),
                writable: false,
                enumerable: true,
                configurable: false));
            form.Dispatch(submit);
            if (submit.IsDefaultPrevented || !RuntimeFocusController.IsConnected(form)) return RuntimeFormActionOutcome.Success;
        }

        string target = submitter?.GetAttribute("formtarget") ?? form.Target ?? string.Empty;
        if (target.Length > 0 && target.ToLowerInvariant() is not ("_self" or "_top" or "_parent"))
            return RuntimeFormActionOutcome.Unsupported("Additional browsing contexts are outside this form profile.");
        string method = HtmlFormControlSemantics.GetEffectiveFormMethod(submitter?.GetAttribute("formmethod") ?? form.Method);
        if (method != "get")
            return RuntimeFormActionOutcome.Unsupported("WebApplicationV1 currently supports only GET form navigation.");

        DocumentRequest? request = CreateGetSubmission(form, submitter, validate);
        if (request == null) return RuntimeFormActionOutcome.Rejected("Form validation or document policy rejected submission.");
        if (request.Method != AngleSharp.Io.HttpMethod.Get)
            return RuntimeFormActionOutcome.Unsupported("WebApplicationV1 currently supports only GET form navigation.");
        try {
            _history.NavigateFragment(request.Target.Href);
            return RuntimeFormActionOutcome.Success;
        } catch (Jint.Runtime.JavaScriptException error) {
            return RuntimeFormActionOutcome.Unsupported(error.Message);
        }
    }

    private DocumentRequest? CreateGetSubmission(IHtmlFormElement form, IHtmlElement? submitter, bool validate) {
        CancellationToken token = _viewport.CurrentCommandToken;
        var live = (IHtmlDocument)_document;
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxInputCharacters = _options.MaxInputCharacters;
        limits.MaxHtmlNodes = _options.MaxNodes;
        limits.MaxHtmlDepth = _options.MaxDepth;
        HtmlConversionInputGuard.ValidateDocument(live, limits, token);

        IElement[] sourceElements = live.QuerySelectorAll("*").ToArray();
        int formIndex = Array.IndexOf(sourceElements, form);
        int submitterIndex = submitter == null ? -1 : Array.IndexOf(sourceElements, submitter);
        if (formIndex < 0 || submitter != null && submitterIndex < 0) return null;

        IHtmlDocument clone = live.Clone(deep: true) as IHtmlDocument
            ?? throw new HtmlScriptRuntimeException("The DOM provider could not clone the form document.");
        IElement[] clonedElements = clone.QuerySelectorAll("*").ToArray();
        CopyFormState(sourceElements, clonedElements, token);
        if (formIndex >= clonedElements.Length || clonedElements[formIndex] is not IHtmlFormElement clonedForm) return null;
        IHtmlElement? clonedSubmitter = submitterIndex < 0 || submitterIndex >= clonedElements.Length
            ? null
            : clonedElements[submitterIndex] as IHtmlElement;
        if (submitter != null && clonedSubmitter == null) return null;

        string action = submitter?.HasAttribute("formaction") == true
            ? submitter.GetAttribute("formaction") ?? string.Empty
            : form.GetAttribute("action") ?? _document.Url;
        var resolvedAction = action.Length == 0
            ? new Url(_document.Url)
            : new Url(new Url(RuntimeDocumentUrls.Base(_document)), action);
        if (resolvedAction.IsInvalid) return null;
        clonedForm.Action = resolvedAction.Href;
        clonedForm.Method = "get";
        return clonedSubmitter == null
            ? validate ? clonedForm.GetSubmission(clonedForm) : clonedForm.GetSubmission()
            : clonedForm.GetSubmission(clonedSubmitter);
    }

    private static void CopyFormState(IReadOnlyList<IElement> source, IReadOnlyList<IElement> target, CancellationToken token) {
        int count = Math.Min(source.Count, target.Count);
        for (int index = 0; index < count; index++) {
            token.ThrowIfCancellationRequested();
            if (source[index] is IHtmlInputElement sourceInput && target[index] is IHtmlInputElement targetInput) {
                targetInput.Value = sourceInput.Value;
                targetInput.IsChecked = sourceInput.IsChecked;
                targetInput.IsIndeterminate = sourceInput.IsIndeterminate;
            } else if (source[index] is IHtmlTextAreaElement sourceArea && target[index] is IHtmlTextAreaElement targetArea) {
                targetArea.Value = sourceArea.Value;
            } else if (source[index] is IHtmlOptionElement sourceOption && target[index] is IHtmlOptionElement targetOption) {
                targetOption.IsSelected = sourceOption.IsSelected;
            }
        }
    }
}

internal readonly record struct RuntimeFormActionOutcome(HtmlAutomationStatus Status, string? Message) {
    internal static RuntimeFormActionOutcome Success => new(HtmlAutomationStatus.Success, null);
    internal static RuntimeFormActionOutcome Rejected(string message) => new(HtmlAutomationStatus.Rejected, message);
    internal static RuntimeFormActionOutcome Invalid(string message) => new(HtmlAutomationStatus.InvalidValue, message);
    internal static RuntimeFormActionOutcome Unsupported(string message) => new(HtmlAutomationStatus.Unsupported, message);
}
