namespace AngleSharp.Js.Dom
{
    using AngleSharp.Dom;
    using AngleSharp.Html.Dom;
    using AngleSharp.Scripting;
    using Jint.Native;
    using Jint.Native.Function;
    using System;
    using System.Collections.Generic;

    sealed class EventAttributeObserver : IAttributeObserver
    {
        private readonly Dictionary<String, Action<IElement, String>> _observers;
        private readonly JsScriptingService _service;

        public EventAttributeObserver(JsScriptingService service)
        {
            _service = service;
            _observers = new Dictionary<String, Action<IElement, String>>();
            RegisterEventCallback<IHtmlBodyElement>(EventNames.AfterPrint);
            RegisterEventCallback<IHtmlBodyElement>(EventNames.BeforePrint);
            RegisterEventCallback<IHtmlBodyElement>(EventNames.Unloading);
            RegisterEventCallback<IHtmlBodyElement>(EventNames.HashChange);
            RegisterEventCallback<IHtmlBodyElement>(EventNames.Message);
            RegisterEventCallback<IHtmlBodyElement>(EventNames.Offline);
            RegisterEventCallback<IHtmlBodyElement>(EventNames.Online);
            RegisterEventCallback<IHtmlBodyElement>(EventNames.PageHide);
            RegisterEventCallback<IHtmlBodyElement>(EventNames.PageShow);
            RegisterEventCallback<IHtmlBodyElement>(EventNames.PopState);
            RegisterEventCallback<IHtmlBodyElement>(EventNames.Storage);
            RegisterEventCallback<IHtmlBodyElement>(EventNames.Unload);
            RegisterEventCallback<IHtmlElement>(EventNames.Load);
            RegisterEventCallback<IHtmlElement>(EventNames.Abort);
            RegisterEventCallback<IHtmlElement>(EventNames.Blur);
            RegisterEventCallback<IHtmlElement>(EventNames.Cancel);
            RegisterEventCallback<IHtmlElement>(EventNames.CanPlay);
            RegisterEventCallback<IHtmlElement>(EventNames.CanPlayThrough);
            RegisterEventCallback<IHtmlElement>(EventNames.Change);
            RegisterEventCallback<IHtmlElement>(EventNames.Click);
            RegisterEventCallback<IHtmlElement>(EventNames.CueChange);
            RegisterEventCallback<IHtmlElement>(EventNames.DblClick);
            RegisterEventCallback<IHtmlElement>(EventNames.Drag);
            RegisterEventCallback<IHtmlElement>(EventNames.DragEnd);
            RegisterEventCallback<IHtmlElement>(EventNames.DragEnter);
            RegisterEventCallback<IHtmlElement>(EventNames.DragExit);
            RegisterEventCallback<IHtmlElement>(EventNames.DragLeave);
            RegisterEventCallback<IHtmlElement>(EventNames.DragOver);
            RegisterEventCallback<IHtmlElement>(EventNames.DragStart);
            RegisterEventCallback<IHtmlElement>(EventNames.Drop);
            RegisterEventCallback<IHtmlElement>(EventNames.DurationChange);
            RegisterEventCallback<IHtmlElement>(EventNames.Emptied);
            RegisterEventCallback<IHtmlElement>(EventNames.Ended);
            RegisterEventCallback<IHtmlElement>(EventNames.Error);
            RegisterEventCallback<IHtmlElement>(EventNames.Focus);
            RegisterEventCallback<IHtmlElement>(EventNames.Input);
            RegisterEventCallback<IHtmlElement>(EventNames.Invalid);
            RegisterEventCallback<IHtmlElement>(EventNames.Keydown);
            RegisterEventCallback<IHtmlElement>(EventNames.Keypress);
            RegisterEventCallback<IHtmlElement>(EventNames.Keyup);
            RegisterEventCallback<IHtmlElement>(EventNames.LoadedData);
            RegisterEventCallback<IHtmlElement>(EventNames.LoadedMetaData);
            RegisterEventCallback<IHtmlElement>(EventNames.LoadStart);
            RegisterEventCallback<IHtmlElement>(EventNames.Mousedown);
            RegisterEventCallback<IHtmlElement>(EventNames.Mouseup);
            RegisterEventCallback<IHtmlElement>(EventNames.Mouseenter);
            RegisterEventCallback<IHtmlElement>(EventNames.Mouseleave);
            RegisterEventCallback<IHtmlElement>(EventNames.Mouseover);
            RegisterEventCallback<IHtmlElement>(EventNames.Mousemove);
            RegisterEventCallback<IHtmlElement>(EventNames.Wheel);
            RegisterEventCallback<IHtmlElement>(EventNames.Pause);
            RegisterEventCallback<IHtmlElement>(EventNames.Play);
            RegisterEventCallback<IHtmlElement>(EventNames.Playing);
            RegisterEventCallback<IHtmlElement>(EventNames.Progress);
            RegisterEventCallback<IHtmlElement>(EventNames.RateChange);
            RegisterEventCallback<IHtmlElement>(EventNames.Reset);
            RegisterEventCallback<IHtmlElement>(EventNames.Resize);
            RegisterEventCallback<IHtmlElement>(EventNames.Scroll);
            RegisterEventCallback<IHtmlElement>(EventNames.Seeked);
            RegisterEventCallback<IHtmlElement>(EventNames.Seeking);
            RegisterEventCallback<IHtmlElement>(EventNames.Select);
            RegisterEventCallback<IHtmlElement>(EventNames.Show);
            RegisterEventCallback<IHtmlElement>(EventNames.Stalled);
            RegisterEventCallback<IHtmlElement>(EventNames.Submit);
            RegisterEventCallback<IHtmlElement>(EventNames.Suspend);
            RegisterEventCallback<IHtmlElement>(EventNames.TimeUpdate);
            RegisterEventCallback<IHtmlElement>(EventNames.Toggle);
            RegisterEventCallback<IHtmlElement>(EventNames.VolumeChange);
            RegisterEventCallback<IHtmlElement>(EventNames.Waiting);
        }

        private void RegisterEventCallback<TElement>(String eventName)
            where TElement : IElement
        {
            _observers.Add("on" + eventName, (sender, value) =>
            {
                if (sender is TElement element)
                {
                    var document = element.Owner;
                    var engine = _service.GetOrCreateInstance(document);
                    var jint = engine.Jint;
                    var instance = jint.Intrinsics.Function.Construct(new JsValue[] { "event", value }, JsValue.Undefined);

                    if (instance is Function functor)
                    {
                        element.AddEventListener(eventName, functor.ToListener(engine));
                    }
                }
            });
        }

        void IAttributeObserver.NotifyChange(IElement host, String name, String value)
        {
            if (_observers.TryGetValue(name, out var observer))
            {
                observer.Invoke(host, value);
            }
        }
    }
}
