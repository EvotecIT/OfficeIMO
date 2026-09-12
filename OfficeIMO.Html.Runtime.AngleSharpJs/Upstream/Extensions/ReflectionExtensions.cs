namespace AngleSharp.Js
{
    using AngleSharp.Attributes;
    using AngleSharp.Text;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Reflection;

    static class ReflectionExtensions
    {
        public static String[] GetParameterNames(this MethodInfo method) =>
            method?.GetParameters().Select(m => m.Name).ToArray();

        public static Assembly GetAssembly(this Type type) => type.GetTypeInfo().Assembly;

        public static Object GetDefaultValue(this Type type) =>
            type.GetTypeInfo().IsValueType ? Activator.CreateInstance(type) : null;

        public static IEnumerable<Type> GetExtensionTypes(this IEnumerable<Assembly> libs, String name) => libs
            .SelectMany(m => m.ExportedTypes)
            .Where(m => m.GetTypeInfo().GetCustomAttributes<DomExposedAttribute>().Any(n => n.Target.Is(name)))
            .ToArray();

        public static IEnumerable<Type> GetTypeTree(this Type root)
        {
            var type = root;
            var types = new List<Type>(type.GetTypeInfo().ImplementedInterfaces);

            do
            {
                types.Add(type);
                type = type.GetTypeInfo().BaseType;
            }
            while (type != null);

            return types;
        }

        public static MethodInfo PrepareConvert(this Type fromType, Type toType)
        {
            try
            {
                // Throws an exception if there is no conversion from fromType to toType
                var exp = Expression.Convert(Expression.Parameter(fromType, null), toType);
                return exp.Method;
            }
            catch
            {
                return null;
            }
        }

        public static String GetOfficialName(this MemberInfo member)
        {
            var names = member.GetCustomAttributes<DomNameAttribute>();
            var officialNameAttribute = names.FirstOrDefault();
            return officialNameAttribute?.OfficialName ?? member.Name;
        }

        public static String GetOfficialName(this Type currentType, Type baseType)
        {
            return currentType.GetOfficialNames(baseType).FirstOrDefault();
        }

        public static String[] GetOfficialNames(this Type currentType, Type baseType)
        {
            var ti = currentType.GetTypeInfo();
            var names = ti.GetCustomAttributes<DomNameAttribute>(true)
                .Select(m => m.OfficialName)
                .Where(m => m != null)
                .Distinct(StringComparer.Ordinal)
                .ToArray();

            if (names.Length > 0)
            {
                return names;
            }

            var interfaces = ti.ImplementedInterfaces;

            if (baseType != null)
            {
                var bi = baseType.GetTypeInfo();
                var exclude = bi.ImplementedInterfaces;
                interfaces = interfaces.Except(exclude);
            }

            foreach (var impl in interfaces)
            {
                names = impl.GetTypeInfo().GetCustomAttributes<DomNameAttribute>(false)
                    .Select(m => m.OfficialName)
                    .Where(m => m != null)
                    .Distinct(StringComparer.Ordinal)
                    .ToArray();

                if (names.Length > 0)
                {
                    return names;
                }
            }

            return Array.Empty<String>();
        }

        public static String GetOfficialName(this Enum value)
        {
            var enumType = value.GetType();
            var member = enumType.GetMember(value.ToString()).FirstOrDefault();

            // if the enum value does not have a DomNameAttribute, calling member.GetOfficialName() would return the value name
            // to allow previous behaviour to be preserved, if the DomNameAttribute is not present then null will be returned
            var names = member.GetCustomAttributes<DomNameAttribute>();
            var officialNameAttribute = names.FirstOrDefault();
            return officialNameAttribute?.OfficialName;
        }

        public static PropertyInfo GetInheritedProperty(this Type type, String propertyName, BindingFlags bindingAttr = BindingFlags.Public | BindingFlags.Instance)
        {
            if (type.GetTypeInfo().IsInterface)
            {
                return type.GetInterfaces()
                    .Union(new[] { type })
                    .Select(i => i.GetProperty(propertyName, bindingAttr))
                    .Where(propertyInfo => propertyInfo != null)
                    .FirstOrDefault();
            }

            return type.GetProperty(propertyName, bindingAttr);
        }

        public static IEnumerable<PropertyInfo> GetInheritedProperties(this Type type, BindingFlags bindingAttr = BindingFlags.Public | BindingFlags.Instance)
        {
            if (type.GetTypeInfo().IsInterface)
            {
                return type.GetInterfaces()
                    .Union(new[] { type })
                    .SelectMany(i => i.GetProperties(bindingAttr))
                    .Distinct();
            }

            return type.GetProperties(bindingAttr);
        }
    }
}
