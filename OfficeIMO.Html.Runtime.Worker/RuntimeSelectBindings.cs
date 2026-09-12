using AngleSharp.Html.Dom;
using Jint;
using Jint.Native;
using Jint.Runtime;
using Jint.Runtime.Descriptors;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

// Keep browser property writes on native option state so automation and capture observe
// the same values. The retained select interface exposes selectedIndex as read-only and
// its value setter compares without case sensitivity and selects every duplicate value.
internal static class RuntimeSelectBindings {
    internal static void Install(Engine engine) {
        var selectPrototype = engine.Global.Get("HTMLSelectElement").AsObject().Get("prototype").AsObject();
        selectPrototype.FastSetProperty("item", new PropertyDescriptor(new ClrFunction(engine, "item", (receiver, args) => {
            var options = Select(receiver).Options;
            uint index = TypeConverter.ToUint32(args.ElementAtOrDefault(0) ?? JsValue.Undefined);
            return index < options.Length ? JsValue.FromObject(engine, options.GetOptionAt((int)index)) : JsValue.Null;
        }), true, true, true));
        selectPrototype.FastSetProperty("namedItem", new PropertyDescriptor(new ClrFunction(engine, "namedItem", (receiver, args) => {
            var options = Select(receiver).Options;
            string name = TypeConverter.ToString(args.ElementAtOrDefault(0) ?? JsValue.Undefined);
            return name.Length == 0 ? JsValue.Null : JsValue.FromObject(engine,
                options.FirstOrDefault(option => option.Id == name || option.GetAttribute("name") == name));
        }), true, true, true));
        selectPrototype.FastSetProperty("value", new GetSetPropertyDescriptor(
            new ClrFunction(engine, "get value", (receiver, _) => Select(receiver).Value ?? string.Empty),
            new ClrFunction(engine, "set value", (receiver, args) => {
                IHtmlSelectElement select = Select(receiver);
                string value = TypeConverter.ToString(args.ElementAtOrDefault(0) ?? JsValue.Undefined);
                bool found = false;
                foreach (IHtmlOptionElement option in select.Options) {
                    bool selected = !found && string.Equals(option.Value, value, StringComparison.Ordinal);
                    option.IsSelected = selected;
                    found |= selected;
                }
                return JsValue.Undefined;
            }), true, true));
        selectPrototype.FastSetProperty("selectedIndex", IndexAccessors(engine, receiver => Select(receiver).Options));
        var optionsPrototype = engine.Global.Get("HTMLOptionsCollection").AsObject().Get("prototype").AsObject();
        optionsPrototype.FastSetProperty("item", new PropertyDescriptor(new ClrFunction(engine, "item", (receiver, args) => {
            var options = Options(receiver);
            uint index = TypeConverter.ToUint32(args.ElementAtOrDefault(0) ?? JsValue.Undefined);
            return index < options.Length ? JsValue.FromObject(engine, options.GetOptionAt((int)index)) : JsValue.Null;
        }), true, true, true));
        optionsPrototype.FastSetProperty("namedItem", new PropertyDescriptor(new ClrFunction(engine, "namedItem", (receiver, args) => {
            var options = Options(receiver);
            string name = TypeConverter.ToString(args.ElementAtOrDefault(0) ?? JsValue.Undefined);
            return name.Length == 0 ? JsValue.Null : JsValue.FromObject(engine,
                options.FirstOrDefault(option => option.Id == name || option.GetAttribute("name") == name));
        }), true, true, true));
        optionsPrototype.FastSetProperty("selectedIndex", IndexAccessors(engine, receiver =>
            Options(receiver)));
    }

    private static GetSetPropertyDescriptor IndexAccessors(Engine engine, Func<JsValue, IHtmlOptionsCollection> options) => new(
        new ClrFunction(engine, "get selectedIndex", (receiver, _) => {
            int index = 0;
            foreach (var option in options(receiver)) {
                if (option.IsSelected) return index;
                index++;
            }
            return -1;
        }),
        new ClrFunction(engine, "set selectedIndex", (receiver, args) => {
            var collection = options(receiver);
            int selectedIndex = TypeConverter.ToInt32(args.ElementAtOrDefault(0) ?? JsValue.Undefined);
            int index = 0;
            foreach (var option in collection) option.IsSelected = index++ == selectedIndex;
            return JsValue.Undefined;
        }), true, true);

    private static IHtmlSelectElement Select(JsValue receiver) => receiver.ToObject() as IHtmlSelectElement
        ?? throw new ArgumentException("The receiver must be a select element.");
    private static IHtmlOptionsCollection Options(JsValue receiver) => receiver.ToObject() as IHtmlOptionsCollection
        ?? throw new ArgumentException("The receiver must be an options collection.");
}
