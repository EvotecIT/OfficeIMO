using System.Text.Json;
using HtmlTinkerX;
using Microsoft.Playwright;
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Runtime;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

// This comparison-only adapter proves that the public contract is provider-neutral.
// It deliberately stays outside the normal solution and package dependency graph.
internal sealed class ChromiumConformanceRuntimeHost(IHtmlParserProvider parser) : IHtmlRuntimeHost, IAsyncDisposable {
    private readonly List<ChromiumConformanceContext> _contexts = new();
    private int _disposed;

    public HtmlRuntimeProviderDescriptor Descriptor { get; } = new(
        "chromium.playwright-comparison",
        typeof(IPlaywright).Assembly.GetName().Version?.ToString() ?? "unknown",
        new[] { HtmlRuntimeProfile.WebApplicationV1 },
        new[] { HtmlRuntimeCapabilityIds.PersistentPage, HtmlRuntimeCapabilityIds.IsolatedContext,
            HtmlRuntimeCapabilityIds.Navigation, HtmlRuntimeCapabilityIds.ScriptEvaluation,
            HtmlRuntimeCapabilityIds.SemanticObservation, HtmlRuntimeCapabilityIds.VisualObservation,
            HtmlRuntimeCapabilityIds.RevisionBoundReferences, HtmlRuntimeCapabilityIds.StructuredActions },
        int.MaxValue, 1);

    public Task<IHtmlRuntimeContext> CreateContextAsync(HtmlRuntimeContextOptions? options = null, CancellationToken cancellationToken = default) {
        ObjectDisposedException.ThrowIf(Volatile.Read(ref _disposed) != 0, this);
        cancellationToken.ThrowIfCancellationRequested();
        var context = new ChromiumConformanceContext(this, parser, (options ?? new()).Snapshot());
        lock (_contexts) _contexts.Add(context);
        return Task.FromResult<IHtmlRuntimeContext>(context);
    }

    public async ValueTask DisposeAsync() {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        ChromiumConformanceContext[] contexts;
        lock (_contexts) { contexts = _contexts.ToArray(); _contexts.Clear(); }
        foreach (ChromiumConformanceContext context in contexts) await context.DisposeAsync();
    }
}

internal sealed class ChromiumConformanceContext(
    ChromiumConformanceRuntimeHost host,
    IHtmlParserProvider parser,
    HtmlRuntimeContextOptions options) : IHtmlRuntimeContext {
    private ChromiumConformancePage? _page;
    private int _disposed;

    public string Id { get; } = options.Id ?? Guid.NewGuid().ToString("N");
    public HtmlRuntimeProviderDescriptor Provider => host.Descriptor;
    public IReadOnlyList<IHtmlRuntimePage> Pages => _page == null ? Array.Empty<IHtmlRuntimePage>() : new IHtmlRuntimePage[] { _page };

    public async Task<IHtmlRuntimePage> OpenPageAsync(HtmlScriptRequest request, CancellationToken cancellationToken = default) {
        ObjectDisposedException.ThrowIf(Volatile.Read(ref _disposed) != 0, this);
        if (_page != null) throw new InvalidOperationException("This comparison adapter supports one page per context.");
        HtmlScriptRequest snapshot = (request ?? throw new ArgumentNullException(nameof(request))).Snapshot();
        if (snapshot.Profile != HtmlRuntimeProfile.WebApplicationV1)
            throw new NotSupportedException("The Chromium comparison adapter qualifies WebApplicationV1.");
        if (snapshot.ResourcePolicy.AllowNetwork)
            throw new NotSupportedException("The Chromium conformance adapter accepts supplied offline resources only.");
        HtmlBrowserSession session = await HtmlBrowser.OpenSessionAsync("about:blank", cancellationToken: cancellationToken).ConfigureAwait(false);
        var page = new ChromiumConformancePage(session, parser, Provider, Id, snapshot);
        _page = page;
        try { await page.InitializeAsync(cancellationToken).ConfigureAwait(false); return page; }
        catch { await page.DisposeAsync(); _page = null; throw; }
    }

    public async ValueTask DisposeAsync() {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        if (_page != null) await _page.DisposeAsync();
        _page = null;
    }
}

internal sealed class ChromiumConformancePage : IHtmlRuntimePage {
    private const long NavigationRevisionSize = 1_000_000_000L;
    private readonly HtmlBrowserSession _session;
    private readonly IHtmlParserProvider _parser;
    private readonly HtmlScriptRequest _options;
    private readonly Dictionary<string, HtmlRuntimeResource> _resources;
    private long _navigationRevision;
    private int _disposed;

    internal ChromiumConformancePage(HtmlBrowserSession session, IHtmlParserProvider parser,
        HtmlRuntimeProviderDescriptor provider, string contextId, HtmlScriptRequest options) {
        _session = session;
        _parser = parser;
        Provider = provider;
        ContextId = contextId;
        Id = Guid.NewGuid().ToString("N");
        _options = options;
        _resources = options.Resources.ToDictionary(resource => Key(resource.Url), StringComparer.Ordinal);
        _resources[Key(options.DocumentUrl)] = HtmlRuntimeResource.FromText(options.DocumentUrl, options.Html, "text/html; charset=utf-8");
    }

    public string Id { get; }
    public string ContextId { get; }
    public HtmlRuntimeProviderDescriptor Provider { get; }

    internal async Task InitializeAsync(CancellationToken token) {
        await _session.Page.SetViewportSizeAsync((int)_options.ViewportWidth, (int)_options.ViewportHeight).WaitAsync(token).ConfigureAwait(false);
        await _session.Page.AddInitScriptAsync("""
            (() => {
              window.__officeimoRevision = 1;
              new MutationObserver(() => window.__officeimoRevision++).observe(document, {subtree:true,childList:true,attributes:true,characterData:true});
            })();
            """).WaitAsync(token).ConfigureAwait(false);
        await _session.Page.RouteAsync("**/*", RouteAsync).WaitAsync(token).ConfigureAwait(false);
        await NavigateCoreAsync(_options.DocumentUrl, replaceHistoryEntry: false, token).ConfigureAwait(false);
        foreach (string script in _options.Scripts)
            await _session.Page.EvaluateAsync(script).WaitAsync(token).ConfigureAwait(false);
        await BumpAsync(token).ConfigureAwait(false);
    }

    private Task RouteAsync(IRoute route) {
        if (_resources.TryGetValue(Key(new Uri(route.Request.Url)), out HtmlRuntimeResource? resource))
            return route.FulfillAsync(new RouteFulfillOptions {
                BodyBytes = resource.Content,
                ContentType = resource.ContentType,
                Status = resource.StatusCode,
                Headers = resource.Headers
            });
        return route.AbortAsync("blockedbyclient");
    }

    public Task NavigateAsync(Uri url, bool replaceHistoryEntry = false, CancellationToken cancellationToken = default) =>
        NavigateCoreAsync(HtmlRuntimeResourcePolicy.ValidateUrl(url), replaceHistoryEntry, cancellationToken);

    private async Task NavigateCoreAsync(Uri url, bool replaceHistoryEntry, CancellationToken token) {
        if (replaceHistoryEntry) {
            Task loaded = _session.Page.WaitForURLAsync(url.AbsoluteUri, new PageWaitForURLOptions { WaitUntil = WaitUntilState.Load });
            await _session.Page.EvaluateAsync("url => window.location.replace(url)", url.AbsoluteUri).WaitAsync(token).ConfigureAwait(false);
            await loaded.WaitAsync(token).ConfigureAwait(false);
        } else {
            await _session.Page.GotoAsync(url.AbsoluteUri, new PageGotoOptions { WaitUntil = WaitUntilState.Load }).WaitAsync(token).ConfigureAwait(false);
        }
        Interlocked.Increment(ref _navigationRevision);
    }

    public async Task ReloadAsync(CancellationToken cancellationToken = default) {
        await _session.Page.ReloadAsync(new PageReloadOptions { WaitUntil = WaitUntilState.Load }).WaitAsync(cancellationToken).ConfigureAwait(false);
        Interlocked.Increment(ref _navigationRevision);
    }

    public async Task ExecuteAsync(string script, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(script);
        if (script.Length > _options.MaxInputCharacters) throw new ArgumentException("The script exceeds MaxInputCharacters.", nameof(script));
        await _session.Page.EvaluateAsync(script).WaitAsync(cancellationToken).ConfigureAwait(false);
        await BumpAsync(cancellationToken).ConfigureAwait(false);
    }

    public async Task<JsonElement> EvaluateAsync(string expression, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(expression);
        JsonElement value = await _session.Page.EvaluateAsync<JsonElement>("value => eval(value)", expression).WaitAsync(cancellationToken).ConfigureAwait(false);
        await BumpAsync(cancellationToken).ConfigureAwait(false);
        return value;
    }

    public async Task WaitForAsync(string expression, CancellationToken cancellationToken = default) {
        DateTimeOffset deadline = DateTimeOffset.UtcNow + _options.Timeout;
        while (!await _session.Page.EvaluateAsync<bool>("value => Boolean(eval(value))", expression).WaitAsync(cancellationToken).ConfigureAwait(false)) {
            if (DateTimeOffset.UtcNow >= deadline) throw new TimeoutException("The runtime command exceeded its deadline.");
            await Task.Delay(_options.PollInterval, cancellationToken).ConfigureAwait(false);
        }
    }

    public async Task<HtmlScriptCapture> CaptureAsync(string? readyExpression = null, CancellationToken cancellationToken = default) {
        await WaitForAsync(readyExpression ?? _options.ReadyExpression, cancellationToken).ConfigureAwait(false);
        string html = await _session.Page.ContentAsync().WaitAsync(cancellationToken).ConfigureAwait(false);
        HtmlDocument document = _parser.Parse(html, new HtmlParseOptions {
            MaxInputCharacters = _options.MaxOutputCharacters,
            MaxNodes = _options.MaxNodes,
            MaxDepth = _options.MaxDepth
        }, cancellationToken);
        return new HtmlScriptCapture(document, Provider.Id + "/" + Provider.Version, new Uri(_session.Page.Url), _options.Resources);
    }

    public async Task<HtmlAutomationResult> AutomateAsync(HtmlAutomationRequest request, CancellationToken cancellationToken = default) {
        HtmlAutomationRequest input = (request ?? throw new ArgumentNullException(nameof(request))).Snapshot(_options.MaxInputCharacters);
        long revision = await CurrentRevisionAsync(cancellationToken).ConfigureAwait(false);
        ILocator locator;
        if (input.Reference != null) {
            if (input.Reference.PageId != Id || input.Reference.Revision != revision)
                return Failure(HtmlAutomationStatus.Stale, "The observed element reference is stale.", revision);
            locator = _session.Page.Locator("*").Nth(input.Reference.ElementIndex);
            var identity = await locator.EvaluateAsync<ElementIdentity>("element => ({ name: element.localName, id: element.id || '' })").WaitAsync(cancellationToken).ConfigureAwait(false);
            if (identity.Name != input.Reference.ElementName || identity.Id != input.Reference.ElementId)
                return Failure(HtmlAutomationStatus.Stale, "The observed element reference was replaced.", revision);
        } else if (input.Query!.Kind == HtmlLocatorKind.Css && input.Query.Scope == null) {
            locator = _session.Page.Locator(input.Query.Value);
            if (input.Query.Index is int index) locator = locator.Nth(index);
        } else return Failure(HtmlAutomationStatus.Unsupported, "This comparison adapter qualifies CSS locators and observed references.", revision);

        if (input.Action == HtmlAutomationAction.Wait)
            return await WaitForLocatorAsync(input, locator, revision, cancellationToken).ConfigureAwait(false);
        int count = await locator.CountAsync().WaitAsync(cancellationToken).ConfigureAwait(false);
        if (input.Action == HtmlAutomationAction.Count) return new HtmlAutomationResult { MatchCount = count, PageRevision = revision };
        if (count == 0) return Failure(HtmlAutomationStatus.NotFound, "The locator matched no element.", revision);
        if (count != 1) return Failure(HtmlAutomationStatus.Ambiguous, "The operation requires one element.", revision, count);
        HtmlRuntimeElementState state = await InspectAsync(locator, cancellationToken).ConfigureAwait(false);
        if (input.Action == HtmlAutomationAction.Inspect) return Success(state, revision);
        try {
            switch (input.Action) {
                case HtmlAutomationAction.Click: await locator.ClickAsync().WaitAsync(cancellationToken).ConfigureAwait(false); break;
                case HtmlAutomationAction.Hover: await locator.HoverAsync().WaitAsync(cancellationToken).ConfigureAwait(false); break;
                case HtmlAutomationAction.Press: await locator.PressAsync(KeyWithModifiers(input.Value!, input.Modifiers)).WaitAsync(cancellationToken).ConfigureAwait(false); break;
                case HtmlAutomationAction.Fill: await locator.FillAsync(input.Value!).WaitAsync(cancellationToken).ConfigureAwait(false); break;
                case HtmlAutomationAction.SetChecked: await locator.SetCheckedAsync(input.Checked!.Value).WaitAsync(cancellationToken).ConfigureAwait(false); break;
                case HtmlAutomationAction.SelectOptions: await locator.SelectOptionAsync(input.Values).WaitAsync(cancellationToken).ConfigureAwait(false); break;
                case HtmlAutomationAction.Focus: await locator.FocusAsync().WaitAsync(cancellationToken).ConfigureAwait(false); break;
                case HtmlAutomationAction.Blur: await locator.BlurAsync().WaitAsync(cancellationToken).ConfigureAwait(false); break;
                case HtmlAutomationAction.ScrollIntoView: await locator.ScrollIntoViewIfNeededAsync().WaitAsync(cancellationToken).ConfigureAwait(false); break;
                case HtmlAutomationAction.SetSelection:
                    await locator.EvaluateAsync("(element, range) => element.setSelectionRange(range.start, range.end)",
                        new { start = input.SelectionStart, end = input.SelectionEnd }).WaitAsync(cancellationToken).ConfigureAwait(false); break;
                default: return Failure(HtmlAutomationStatus.Unsupported, "Unsupported comparison action.", revision);
            }
        } catch (PlaywrightException error) { return Failure(HtmlAutomationStatus.Rejected, error.Message, revision, 1, state); }
        await BumpAsync(cancellationToken).ConfigureAwait(false);
        revision = await CurrentRevisionAsync(cancellationToken).ConfigureAwait(false);
        return Success(await InspectAsync(locator, cancellationToken).ConfigureAwait(false), revision);
    }

    public async Task<HtmlPageObservation> ObserveAsync(HtmlPageObservationRequest? request = null, CancellationToken cancellationToken = default) {
        HtmlPageObservationRequest input = (request ?? new()).Snapshot(_options.MaxOutputCharacters);
        if (input.IncludeScreenshotReference) throw new NotSupportedException("The comparison adapter does not retain screenshot artifacts.");
        long revision = await CurrentRevisionAsync(cancellationToken).ConfigureAwait(false);
        BrowserObservation observed = await _session.Page.EvaluateAsync<BrowserObservation>(ObservationScript, new {
            pageId = Id,
            revision,
            mode = (int)input.Mode,
            input.MaxElements,
            input.MaxTextCharacters,
            input.IncludeHidden,
            input.ActionableOnly
        }).WaitAsync(cancellationToken).ConfigureAwait(false);
        return observed.ToPublic(Provider.Id, ContextId, Id, revision, input.Mode);
    }

    public HtmlRuntimeTrace GetTrace() => new() { ProviderId = Provider.Id, ContextId = ContextId, PageId = Id };

    private async Task<HtmlRuntimeElementState> InspectAsync(ILocator locator, CancellationToken token) {
        BrowserElement state = await locator.EvaluateAsync<BrowserElement>(InspectScript).WaitAsync(token).ConfigureAwait(false);
        return state.ToState();
    }

    private async Task<HtmlAutomationResult> WaitForLocatorAsync(HtmlAutomationRequest input, ILocator locator,
        long revision, CancellationToken token) {
        DateTimeOffset deadline = DateTimeOffset.UtcNow + _options.Timeout;
        while (true) {
            int count = await locator.CountAsync().WaitAsync(token).ConfigureAwait(false);
            revision = await CurrentRevisionAsync(token).ConfigureAwait(false);
            if (count == 0 && input.WaitState is HtmlLocatorWaitState.Detached or HtmlLocatorWaitState.Hidden)
                return new HtmlAutomationResult { PageRevision = revision };
            HtmlAutomationResult result;
            if (count == 0) result = Failure(HtmlAutomationStatus.NotFound, "The locator matched no element.", revision);
            else if (input.WaitState == HtmlLocatorWaitState.Detached)
                result = Failure(HtmlAutomationStatus.NotReady, "The locator still matches attached elements.", revision, count);
            else if (count != 1) return Failure(HtmlAutomationStatus.Ambiguous, "The operation requires one element.", revision, count);
            else {
                HtmlRuntimeElementState state = await InspectAsync(locator, token).ConfigureAwait(false);
                bool ready = input.WaitState switch {
                    HtmlLocatorWaitState.Attached => true,
                    HtmlLocatorWaitState.Enabled => !state.IsDisabled,
                    HtmlLocatorWaitState.Disabled => state.IsDisabled,
                    HtmlLocatorWaitState.Editable => state.IsEditable,
                    HtmlLocatorWaitState.Focused => state.IsFocused,
                    HtmlLocatorWaitState.Visible => state.IsVisible == true,
                    HtmlLocatorWaitState.Hidden => state.IsVisible == false,
                    HtmlLocatorWaitState.InViewport => state.IsInViewport == true,
                    HtmlLocatorWaitState.Value => await locator.InputValueAsync().WaitAsync(token).ConfigureAwait(false) == input.Value,
                    HtmlLocatorWaitState.Text => state.Text == Normalize(input.Value),
                    HtmlLocatorWaitState.Checked => state.IsChecked == input.Checked,
                    _ => false
                };
                result = ready ? Success(state, revision)
                    : Failure(HtmlAutomationStatus.NotReady, "The requested state has not been reached.", revision, 1, state);
            }
            if (result.Status == HtmlAutomationStatus.Success || !input.WaitForReady || input.Reference != null)
                return result;
            if (DateTimeOffset.UtcNow >= deadline) throw new TimeoutException("The runtime command exceeded its deadline.");
            await Task.Delay(_options.PollInterval, token).ConfigureAwait(false);
        }
    }

    private async Task<long> CurrentRevisionAsync(CancellationToken token) {
        long local = await _session.Page.EvaluateAsync<long>("() => window.__officeimoRevision || 1").WaitAsync(token).ConfigureAwait(false);
        return Interlocked.Read(ref _navigationRevision) * NavigationRevisionSize + local;
    }

    private Task BumpAsync(CancellationToken token) => _session.Page.EvaluateAsync("() => window.__officeimoRevision++").WaitAsync(token);
    private static string Key(Uri url) => new UriBuilder(url) { Fragment = string.Empty }.Uri.AbsoluteUri;
    private static string Normalize(string? value) => string.Join(" ", (value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    private static string KeyWithModifiers(string key, HtmlKeyboardModifiers modifiers) {
        var parts = new List<string>(5);
        if ((modifiers & HtmlKeyboardModifiers.Alt) != 0) parts.Add("Alt");
        if ((modifiers & HtmlKeyboardModifiers.Control) != 0) parts.Add("Control");
        if ((modifiers & HtmlKeyboardModifiers.Meta) != 0) parts.Add("Meta");
        if ((modifiers & HtmlKeyboardModifiers.Shift) != 0) parts.Add("Shift");
        parts.Add(key);
        return string.Join("+", parts);
    }
    private static HtmlAutomationResult Success(HtmlRuntimeElementState state, long revision) => new() { MatchCount = 1, Element = state, PageRevision = revision };
    private static HtmlAutomationResult Failure(HtmlAutomationStatus status, string message, long revision, int count = 0, HtmlRuntimeElementState? state = null) =>
        new() { Status = status, Message = message, MatchCount = count, Element = state, PageRevision = revision };

    public async ValueTask DisposeAsync() {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        await HtmlBrowser.CloseSessionAsync(_session).ConfigureAwait(false);
    }

    private const string InspectScript = """
        element => {
          const rect=element.getBoundingClientRect(), style=getComputedStyle(element), visible=!!(rect.width&&rect.height)&&style.display!=='none'&&style.visibility!=='hidden';
          const sensitive=element instanceof HTMLInputElement&&element.type==='password';
          const value=!sensitive&&('value' in element)?String(element.value):null;
          const text=(element.textContent||'').replace(/\s+/g,' ').trim();
          const name=(element.getAttribute('aria-label')||element.alt||element.title||((['BUTTON','A','H1','H2','H3','H4','H5','H6','OPTION'].includes(element.tagName))?text:'')).replace(/\s+/g,' ').trim();
          return {elementName:element.localName,id:element.id||'',accessibleName:name,text,value,
            selectionStart:!sensitive&&Number.isInteger(element.selectionStart)?element.selectionStart:null,selectionEnd:!sensitive&&Number.isInteger(element.selectionEnd)?element.selectionEnd:null,
            selectedValues:element.selectedOptions?Array.from(element.selectedOptions,x=>x.value):[],isChecked:('checked' in element)?!!element.checked:null,
            isIndeterminate:!!element.indeterminate,isDisabled:!!element.disabled,isReadOnly:!!element.readOnly,
            isEditable:['INPUT','TEXTAREA'].includes(element.tagName)&&!element.disabled&&!element.readOnly,
            isHiddenByMarkup:element.hidden||!!element.closest('[hidden],[inert]'),isFocused:document.activeElement===element,isConnected:element.isConnected,
            isVisible:visible,isInViewport:visible&&rect.left<innerWidth&&rect.top<innerHeight&&rect.right>0&&rect.bottom>0,
            acceptsPointerEvents:style.pointerEvents!=='none',receivesPointerAtCenter:true,
            boundingBox:visible?{x:rect.x,y:rect.y,width:rect.width,height:rect.height}:null,scrollX,scrollY};
        }
        """;

    private const string ObservationScript = """
        args => {
          const all=Array.from(document.querySelectorAll('*')), indexes=new Map(all.map((e,i)=>[e,i])), included=new Map(), elements=[];
          let chars=0,truncated=false;
          const norm=v=>(v||'').replace(/\s+/g,' ').trim();
          const role=e=>{const r=(e.getAttribute('role')||'').split(/\s+/)[0];if(r)return r;const n=e.localName;if(n==='a'&&e.hasAttribute('href'))return'link';if(n==='button')return'button';if(/^h[1-6]$/.test(n))return'heading';if(n==='input')return e.type==='checkbox'?'checkbox':e.type==='radio'?'radio':'textbox';if(n==='textarea')return'textbox';if(n==='select')return'combobox';if(n==='option')return'option';return''};
          for(let index=0;index<all.length;index++){
            const e=all[index],rect=e.getBoundingClientRect(),style=getComputedStyle(e),hidden=e.hidden||!!e.closest('[hidden],[inert]'),visible=!hidden&&!!(rect.width&&rect.height)&&style.display!=='none'&&style.visibility!=='hidden';
            const actionable=['a','button','input','select','textarea','summary'].includes(e.localName)||e.hasAttribute('tabindex')||['button','link','checkbox','radio','textbox','combobox'].includes(e.getAttribute('role'));
            if(!args.IncludeHidden&&(!visible||hidden))continue;if(args.ActionableOnly&&!actionable)continue;if(elements.length>=args.MaxElements){truncated=true;break;}
            let text=args.mode===1?'':norm(e.textContent),name=args.mode===1?'':norm(e.getAttribute('aria-label')||e.alt||e.title||(['button','a','option','h1','h2','h3','h4','h5','h6'].includes(e.localName)?text:''));
            const room=Math.max(0,args.MaxTextCharacters-chars);if(name.length+text.length>room){const combined=(name+text).slice(0,room);name=combined.slice(0,Math.min(name.length,combined.length));text=combined.slice(name.length);truncated=true;}chars+=name.length+text.length;
            let parent=e.parentElement,parentObserved=null;while(parent){const pi=indexes.get(parent);if(included.has(pi)){parentObserved=included.get(pi);break;}parent=parent.parentElement;}included.set(index,elements.length);
            elements.push({reference:{pageId:args.pageId,revision:args.revision,elementIndex:index,elementName:e.localName,elementId:e.id||''},depth:(()=>{let d=1,p=e.parentElement;while(p){d++;p=p.parentElement;}return d;})(),parentElementIndex:parentObserved,elementName:e.localName,role:args.mode===1?'':role(e),accessibleName:name,text,
              value:args.mode===1||e instanceof HTMLInputElement&&e.type==='password'?null:(('value'in e)?String(e.value):null),selectedValues:args.mode===1?[]:(e.selectedOptions?Array.from(e.selectedOptions,x=>x.value):[]),selectionStart:args.mode===1||e instanceof HTMLInputElement&&e.type==='password'?null:(Number.isInteger(e.selectionStart)?e.selectionStart:null),selectionEnd:args.mode===1||e instanceof HTMLInputElement&&e.type==='password'?null:(Number.isInteger(e.selectionEnd)?e.selectionEnd:null),isChecked:args.mode===1?null:(('checked'in e)?!!e.checked:null),isDisabled:!!e.disabled,isEditable:['INPUT','TEXTAREA'].includes(e.tagName)&&!e.disabled&&!e.readOnly,isFocused:document.activeElement===e,isVisible:args.mode===0?null:visible,isInViewport:args.mode===0?null:(visible&&rect.left<innerWidth&&rect.top<innerHeight&&rect.right>0&&rect.bottom>0),isActionable:actionable&&!e.disabled&&!hidden&&visible,boundingBox:args.mode===0||!visible?null:{x:rect.x,y:rect.y,width:rect.width,height:rect.height}});
          }
          return {url:location.href,title:document.title,viewportWidth:innerWidth,viewportHeight:innerHeight,scrollX,scrollY,documentWidth:document.documentElement.scrollWidth,documentHeight:document.documentElement.scrollHeight,isTruncated:truncated,elements};
        }
        """;

    private sealed class ElementIdentity { public string Name { get; set; } = ""; public string Id { get; set; } = ""; }
    private sealed class BrowserObservation {
        public string Url { get; set; } = "https://officeimo.invalid/"; public string Title { get; set; } = "";
        public double ViewportWidth { get; set; } public double ViewportHeight { get; set; } public double ScrollX { get; set; } public double ScrollY { get; set; }
        public double DocumentWidth { get; set; } public double DocumentHeight { get; set; } public bool IsTruncated { get; set; }
        public BrowserObservedElement[] Elements { get; set; } = Array.Empty<BrowserObservedElement>();
        public HtmlPageObservation ToPublic(string providerId,string contextId,string pageId,long revision,HtmlPageObservationMode mode)=>new(){ProviderId=providerId,ContextId=contextId,PageId=pageId,Revision=revision,Url=new Uri(Url),Title=Title,Mode=mode,ViewportWidth=ViewportWidth,ViewportHeight=ViewportHeight,ScrollX=ScrollX,ScrollY=ScrollY,DocumentWidth=DocumentWidth,DocumentHeight=DocumentHeight,IsTruncated=IsTruncated,Elements=Elements.Select(x=>x.ToPublic()).ToArray()};
    }
    private sealed class BrowserObservedElement : BrowserElement {
        public HtmlObservedElementReference Reference { get; set; } = null!; public int Depth { get; set; } public int? ParentElementIndex { get; set; } public string Role { get; set; } = ""; public bool IsActionable { get; set; }
        public HtmlObservedElement ToPublic()=>new(){Reference=Reference,Depth=Depth,ParentElementIndex=ParentElementIndex,ElementName=ElementName,Role=Role,AccessibleName=AccessibleName,Text=Text,Value=Value,SelectedValues=SelectedValues,SelectionStart=SelectionStart,SelectionEnd=SelectionEnd,IsChecked=IsChecked,IsDisabled=IsDisabled,IsEditable=IsEditable,IsFocused=IsFocused,IsVisible=IsVisible,IsInViewport=IsInViewport,IsActionable=IsActionable,BoundingBox=BoundingBox};
    }
    private class BrowserElement {
        public string ElementName { get; set; }=""; public string Id { get; set; }=""; public string AccessibleName { get; set; }=""; public string Text { get; set; }=""; public string? Value { get; set; }
        public int? SelectionStart { get; set; } public int? SelectionEnd { get; set; } public string[] SelectedValues { get; set; }=Array.Empty<string>(); public bool? IsChecked { get; set; } public bool IsIndeterminate { get; set; } public bool IsDisabled { get; set; } public bool IsReadOnly { get; set; } public bool IsEditable { get; set; } public bool IsHiddenByMarkup { get; set; } public bool IsFocused { get; set; } public bool IsConnected { get; set; } public bool? IsVisible { get; set; } public bool? IsInViewport { get; set; } public bool? AcceptsPointerEvents { get; set; } public bool? ReceivesPointerAtCenter { get; set; } public HtmlRuntimeRect? BoundingBox { get; set; } public double ScrollX { get; set; } public double ScrollY { get; set; }
        public HtmlRuntimeElementState ToState()=>new(){ElementName=ElementName,Id=Id,AccessibleName=AccessibleName,Text=Text,Value=Value,SelectionStart=SelectionStart,SelectionEnd=SelectionEnd,SelectedValues=SelectedValues,IsChecked=IsChecked,IsIndeterminate=IsIndeterminate,IsDisabled=IsDisabled,IsReadOnly=IsReadOnly,IsEditable=IsEditable,IsHiddenByMarkup=IsHiddenByMarkup,IsFocused=IsFocused,IsConnected=IsConnected,IsVisible=IsVisible,IsInViewport=IsInViewport,AcceptsPointerEvents=AcceptsPointerEvents,ReceivesPointerAtCenter=ReceivesPointerAtCenter,BoundingBox=BoundingBox,ScrollX=ScrollX,ScrollY=ScrollY};
    }
}
