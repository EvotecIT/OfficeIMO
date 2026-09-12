namespace AngleSharp.Js
{
    using AngleSharp.Attributes;
    using Jint;
    using Jint.Native;
    using Jint.Native.Function;
    using Jint.Runtime;
    using Jint.Runtime.Interop;
    using System.Linq;
    using System;
    using System.Reflection;

    static class JsValueExtensions
    {
        private static Object AsComplex(this JsValue value, Type targetType, EngineInstance engine)
        {
            var obj = value.FromJsValue();
            var sourceType = obj.GetType();

            if (sourceType == targetType || sourceType.GetTypeInfo().IsSubclassOf(targetType) || targetType.GetTypeInfo().IsAssignableFrom(sourceType.GetTypeInfo()))
            {
                return obj;
            }
            else if (targetType.GetTypeInfo().IsSubclassOf(typeof(Delegate)))
            {
                var f = obj as Function;

                if (f == null && obj is String b)
                {
                    var e = engine.Jint;
                    var p = new[] { JsValue.FromObjectWithType(e, b, typeof(String)) };
                    f = new ClrFunction(e, "AsComplex", (_this, args) => e.Intrinsics.Eval.Call(_this, p));
                }

                if (f != null)
                {
                    return targetType.ToDelegate(f, engine);
                }
            }

            var method = sourceType.PrepareConvert(targetType) ??
                throw new JavaScriptException("[Internal] Could not find corresponding cast target.");

            return method.Invoke(obj, null);
        }

        public static Object FromJsValue(this JsValue val)
        {
            switch (val.Type)
            {
                case Types.Boolean:
                    return val.AsBoolean();
                case Types.Number:
                    return val.AsNumber();
                case Types.String:
                    return val.AsString();
                case Types.Object:
                    var obj = val.AsObject();
                    var node = obj as IDomProxy;
                    return node != null ? node.Value : obj;
                case Types.Undefined:
                    return JsValue.Undefined.ToString();
                case Types.Null:
                    return null;
            }

            return val.ToObject();
        }

        public static Object As(this JsValue value, Type targetType, EngineInstance engine)
        {
            if (value != JsValue.Null)
            {
                if (targetType == typeof(Int32))
                {
                    return TypeConverter.ToInt32(value);
                }
                else if (targetType == typeof(Nullable<Int32>))
                {
                    if (value.IsUndefined())
                    {
                        return null;
                    }
                    else
                    {
                        return TypeConverter.ToInt32(value);
                    }
                }
                else if (targetType == typeof(Double))
                {
                    return TypeConverter.ToNumber(value);
                }
                else if (targetType == typeof(String))
                {
                    return value.IsPrimitive() ? TypeConverter.ToString(value) : value.ToString();
                }
                else if (targetType == typeof(Boolean))
                {
                    return TypeConverter.ToBoolean(value);
                }
                else if (targetType == typeof(UInt32))
                {
                    return TypeConverter.ToUint32(value);
                }
                else if (targetType == typeof(UInt16))
                {
                    return TypeConverter.ToUint16(value);
                }
                else if (targetType.GetTypeInfo().IsEnum)
                {
                    if (value.IsString())
                    {
                        var literal = TypeConverter.ToString(value);
                        var member = targetType
                            .GetTypeInfo()
                            .DeclaredFields
                            .Where(m => m.IsLiteral)
                            .FirstOrDefault(m => m.GetCustomAttributes<DomNameAttribute>().FirstOrDefault()?.OfficialName == literal || m.Name == literal);

                        if (member != null)
                        {
                            return Enum.Parse(targetType, member.Name);
                        }
                    }

                    var underlyingType = Enum.GetUnderlyingType(targetType);
                    var raw = Convert.ChangeType(TypeConverter.ToNumber(value), underlyingType);
                    return Enum.ToObject(targetType, raw);
                }
                else
                {
                    return value.AsComplex(targetType, engine);
                }
            }

            return null;
        }
    }
}
