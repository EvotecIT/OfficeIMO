namespace AngleSharp.Js.Cache
{
    using AngleSharp.Attributes;
    using AngleSharp.Text;
    using System;
    using System.Collections.Concurrent;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Reflection;

    /// <summary>
    /// Decides whether a DOM type is an indexed collection - a node list, an HTML collection, a
    /// token list, a named node map - and resolves the two members that make it one.
    /// </summary>
    /// <remarks>
    /// This is the WebIDL shape: an indexed property getter plus a "length". Both are required,
    /// because the engine derives the whole array-like property model from them and an object
    /// offering only one of the two would be a collection whose end nothing can find. The answer
    /// depends on the type alone, so the cache is process-wide - the reflection runs once per
    /// type however many documents are opened.
    /// </remarks>
    static class CollectionTypeCache
    {
        private static readonly ConcurrentDictionary<Type, IndexedCollection> _collections = new();

        public static IndexedCollection GetIndexedCollection(this Type type) =>
            _collections.GetOrAdd(type, static current => Resolve(current));

        private static IndexedCollection Resolve(Type type)
        {
            MethodInfo indexer = null;
            MethodInfo length = null;

            foreach (var current in type.GetTypeTree())
            {
                foreach (var property in current.GetTypeInfo().DeclaredProperties)
                {
                    var getter = property.GetMethod;

                    if (getter == null)
                    {
                        continue;
                    }

                    var indexParameters = property.GetIndexParameters();

                    if (indexParameters.Length == 1)
                    {
                        //  The same test SetIndexer applies, so a type is a collection here
                        //  exactly when its prototype would have offered a numeric indexer.
                        if (indexer == null &&
                            indexParameters[0].ParameterType == typeof(Int32) &&
                            IsIndexedAccessor(property))
                        {
                            indexer = getter;
                        }
                    }
                    else if (length == null && indexParameters.Length == 0 && IsLength(property))
                    {
                        length = getter;
                    }
                }
            }

            return indexer != null && length != null ? IndexedCollection.TryCreate(type, indexer, length) : null;
        }

        private static Boolean IsIndexedAccessor(PropertyInfo property)
        {
            var accessor = property.GetCustomAttribute<DomAccessorAttribute>()?.Type;

            if (accessor.HasValue && (accessor.Value & (Accessors.Getter | Accessors.Setter)) != 0)
            {
                return true;
            }

            return property.GetCustomAttributes<DomNameAttribute>().Any(m => m.OfficialName.Is("item"));
        }

        private static Boolean IsLength(PropertyInfo property) =>
            property.PropertyType == typeof(Int32) &&
            property.GetCustomAttributes<DomNameAttribute>().Any(m => m.OfficialName.Is("length"));
    }

    /// <summary>
    /// The indexed getter and the length of a collection type, compiled once.
    /// </summary>
    /// <remarks>
    /// These two run on every single indexed read, which is the one place in this binding where
    /// reflection is genuinely too slow: the engine reads an element without materialising a
    /// descriptor or a key, so a MethodInfo.Invoke and the object[] it needs would be all that is
    /// left of the cost. They are compiled instead, against the interface that declares them, so
    /// the call is a virtual dispatch that lands on whatever concrete class the target happens to
    /// be - the case a reflective invoke of an explicitly re-implemented member gets wrong.
    /// </remarks>
    sealed class IndexedCollection
    {
        private readonly Func<Object, Int32> _length;
        private readonly Func<Object, Int32, Object> _item;

        private IndexedCollection(Func<Object, Int32> length, Func<Object, Int32, Object> item)
        {
            _length = length;
            _item = item;
        }

        public static IndexedCollection TryCreate(Type type, MethodInfo indexer, MethodInfo length)
        {
            //  What carries the DOM attributes and what can actually be called are not always the
            //  same member: an interface may re-implement one of its own bases explicitly, as
            //  IHtmlCollection<T> does with "T IReadOnlyList<T>.this[Int32]", and that member is
            //  private and abstract. The declaration is what identifies the collection, the
            //  public method of the same name on the concrete class is what runs.
            var resolvedIndexer = ResolvePublic(type, indexer);
            var resolvedLength = ResolvePublic(type, length);

            if (resolvedIndexer == null || resolvedLength == null)
            {
                return null;
            }

            try
            {
                return new IndexedCollection(
                    CompileLength(type, resolvedLength),
                    CompileItem(type, resolvedIndexer));
            }
            catch (Exception)
            {
                //  A member that cannot be expressed as an expression tree leaves the type on the
                //  ordinary node proxy, which reaches the same accessors reflectively.
                return null;
            }
        }

        private static MethodInfo ResolvePublic(Type type, MethodInfo declared)
        {
            if (declared.IsPublic && declared.DeclaringType.GetTypeInfo().IsAssignableFrom(type.GetTypeInfo()))
            {
                return declared;
            }

            var name = declared.Name;
            var simpleName = name.Substring(name.LastIndexOf('.') + 1);
            var parameters = declared.GetParameters();
            var parameterTypes = new Type[parameters.Length];

            for (var i = 0; i < parameters.Length; i++)
            {
                parameterTypes[i] = parameters[i].ParameterType;
            }

            var resolved = type.GetRuntimeMethod(simpleName, parameterTypes);
            return resolved != null && resolved.IsPublic ? resolved : null;
        }

        public Int32 GetLength(Object target) => _length.Invoke(target);

        public Object GetItem(Object target, Int32 index) => _item.Invoke(target, index);

        private static Func<Object, Int32> CompileLength(Type type, MethodInfo length)
        {
            var target = Expression.Parameter(typeof(Object), "target");
            var call = Expression.Call(Expression.Convert(target, type), length);
            return Expression.Lambda<Func<Object, Int32>>(call, target).Compile();
        }

        private static Func<Object, Int32, Object> CompileItem(Type type, MethodInfo indexer)
        {
            var target = Expression.Parameter(typeof(Object), "target");
            var index = Expression.Parameter(typeof(Int32), "index");
            var call = Expression.Call(Expression.Convert(target, type), indexer, index);
            return Expression.Lambda<Func<Object, Int32, Object>>(
                Expression.Convert(call, typeof(Object)), target, index).Compile();
        }
    }
}
