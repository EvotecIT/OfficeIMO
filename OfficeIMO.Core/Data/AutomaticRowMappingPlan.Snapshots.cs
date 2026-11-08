using System;
using System.Data.Common;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;

namespace OfficeIMO.Data;

/// <summary>Owns captured values until the existing bounded snapshot worker returns its batch.</summary>
internal interface IRowMappingSnapshotColumns {
    void Capture(DbDataReader reader, int row);
    void Clear();
}

internal interface IRowMappingSnapshotColumns<T> : IRowMappingSnapshotColumns {
    T Map(int row);
}

internal sealed partial class AutomaticRowMappingPlan<T> {
    /// <summary>Uses typed owned columns only for the existing compiled, mutable-class mapping path.</summary>
    internal Func<int, IRowMappingSnapshotColumns>? CreateSnapshotColumnsFactory(
        DbDataReader reader,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        DataMappingErrorValuePolicy errorValuePolicy) {
        if (_constructor is not null || typeof(T).IsValueType || _bindings.Length == 0
            || reader is not IDataReaderFastMappingValues && reader is not IDataReaderMappingRowSource) return null;
        var factories = new Func<int, SnapshotColumn>[_bindings.Length];
        bool hasTypedColumn = false;
        for (int index = 0; index < _bindings.Length; index++) {
            MappingBinding binding = _bindings[index];
            Func<int, SnapshotColumn>? factory = CreateTypedSnapshotColumnFactory(binding);
            hasTypedColumn |= factory is not null;
            factories[index] = factory ?? (capacity => new SnapshotColumn(binding, capacity));
        }
        return hasTypedColumn
            ? capacity => new SnapshotColumns(factories, capacity, culture, dateTimeFormats, errorValuePolicy)
            : null;
    }

    private static Func<int, SnapshotColumn>? CreateTypedSnapshotColumnFactory(MappingBinding binding) {
        if (!binding.Property.SupportsTypedSnapshot) return null;
        Type type = binding.Property.ValueType;
        if (type == typeof(int)) return CreateTypedSnapshotColumnFactory<int>(binding, static (reader, ordinal) => reader.GetInt32(ordinal));
        if (type == typeof(long)) return CreateTypedSnapshotColumnFactory<long>(binding, static (reader, ordinal) => reader.GetInt64(ordinal));
        if (type == typeof(double)) return CreateTypedSnapshotColumnFactory<double>(binding, static (reader, ordinal) => reader.GetDouble(ordinal));
        if (type == typeof(float)) return CreateTypedSnapshotColumnFactory<float>(binding, static (reader, ordinal) => reader.GetFloat(ordinal));
        if (type == typeof(decimal)) return CreateTypedSnapshotColumnFactory<decimal>(binding, static (reader, ordinal) => reader.GetDecimal(ordinal));
        if (type == typeof(DateTime)) return CreateTypedSnapshotColumnFactory<DateTime>(binding, static (reader, ordinal) => reader.GetDateTime(ordinal));
        if (type == typeof(string)) return CreateTypedSnapshotColumnFactory<string>(binding, static (reader, ordinal) => reader.GetString(ordinal));
        if (type == typeof(bool)) return CreateTypedSnapshotColumnFactory<bool>(binding, static (reader, ordinal) => reader.GetBoolean(ordinal));
        if (type == typeof(short)) return CreateTypedSnapshotColumnFactory<short>(binding, static (reader, ordinal) => reader.GetInt16(ordinal));
        if (type == typeof(byte)) return CreateTypedSnapshotColumnFactory<byte>(binding, static (reader, ordinal) => reader.GetByte(ordinal));
        if (type == typeof(char)) return CreateTypedSnapshotColumnFactory<char>(binding, static (reader, ordinal) => reader.GetChar(ordinal));
        if (type == typeof(Guid)) return CreateTypedSnapshotColumnFactory<Guid>(binding, static (reader, ordinal) => reader.GetGuid(ordinal));
        return null;
    }

    private static Func<int, SnapshotColumn>? CreateTypedSnapshotColumnFactory<TValue>(
        MappingBinding binding,
        Func<DbDataReader, int, TValue> read) {
        try {
            var assign = (Action<T, TValue>)binding.Property.Property.SetMethod!.CreateDelegate(typeof(Action<T, TValue>));
            return capacity => new TypedSnapshotColumn<TValue>(binding, capacity, read, assign);
        } catch (ArgumentException) {
        } catch (MemberAccessException) {
        } catch (NotSupportedException) {
        }
        return null;
    }

    private sealed class SnapshotColumns : IRowMappingSnapshotColumns<T> {
        private readonly SnapshotColumn[] _columns;
        private readonly CultureInfo _culture;
        private readonly IReadOnlyList<string>? _dateTimeFormats;
        private readonly DataMappingErrorValuePolicy _errorValuePolicy;
        private readonly int _maximumCapacity;
        private int _capacity;

        internal SnapshotColumns(Func<int, SnapshotColumn>[] factories, int capacity,
            CultureInfo culture, IReadOnlyList<string>? dateTimeFormats,
            DataMappingErrorValuePolicy errorValuePolicy) {
            _maximumCapacity = capacity;
            _capacity = Math.Min(capacity, 128);
            _columns = new SnapshotColumn[factories.Length];
            for (int index = 0; index < factories.Length; index++) _columns[index] = factories[index](_capacity);
            _culture = culture;
            _dateTimeFormats = dateTimeFormats;
            _errorValuePolicy = errorValuePolicy;
        }

        public void Capture(DbDataReader reader, int row) {
            if (row == _capacity) {
                int next = _capacity <= _maximumCapacity / 2 ? _capacity * 2 : _maximumCapacity;
                foreach (SnapshotColumn column in _columns) column.Resize(next);
                _capacity = next;
            }
            DbDataReader fastReader = (reader as IDataReaderMappingRowSource)?.CurrentRowReader ?? reader;
            foreach (SnapshotColumn column in _columns) column.Capture(reader, fastReader, row);
            // Keep the original post-capture cloning point for providers that reuse
            // one borrowed byte/character buffer across multiple fields.
            foreach (SnapshotColumn column in _columns) column.CloneBorrowedValue(row);
        }

        public T Map(int row) {
            T instance = Activator.CreateInstance<T>();
            foreach (SnapshotColumn column in _columns)
                instance = column.Apply(instance, row, _culture, _dateTimeFormats, _errorValuePolicy);
            return instance;
        }

        public void Clear() {
            foreach (SnapshotColumn column in _columns) column.Clear();
        }
    }

    private class SnapshotColumn {
        protected readonly MappingBinding Binding;
        protected object?[]? Fallback;
        protected int Capacity;

        internal SnapshotColumn(MappingBinding binding, int capacity) {
            Binding = binding;
            Capacity = capacity;
        }

        internal virtual void Capture(DbDataReader reader, DbDataReader fastReader, int row) =>
            (Fallback ??= new object?[Capacity])[row] = Binding.Property.ReadReaderValue(reader, fastReader, Binding.ColumnIndex);

        internal void CloneBorrowedValue(int row) {
            if (Fallback is null) return;
            if (Fallback[row] is byte[] bytes) Fallback[row] = (byte[])bytes.Clone();
            else if (Fallback[row] is char[] characters) Fallback[row] = (char[])characters.Clone();
        }

        internal virtual T Apply(T instance, int row, CultureInfo culture,
            IReadOnlyList<string>? dateTimeFormats, DataMappingErrorValuePolicy errorValuePolicy) =>
            Binding.Apply(instance, Fallback is null ? null : Fallback[row], culture, dateTimeFormats, null, errorValuePolicy);

        internal virtual void Resize(int capacity) {
            Capacity = capacity;
            if (Fallback is not null) Array.Resize(ref Fallback, capacity);
        }

        internal virtual void Clear() {
            if (Fallback is not null) Array.Clear(Fallback, 0, Fallback.Length);
        }
    }

    private sealed class TypedSnapshotColumn<TValue> : SnapshotColumn {
        private TValue[]? _values;
        private bool[] _typed;
        private readonly Func<DbDataReader, int, TValue> _read;
        private readonly Action<T, TValue> _assign;

        internal TypedSnapshotColumn(MappingBinding binding, int capacity,
            Func<DbDataReader, int, TValue> read, Action<T, TValue> assign) : base(binding, capacity) {
            _typed = new bool[capacity];
            _read = read;
            _assign = assign;
        }

        internal override void Capture(DbDataReader reader, DbDataReader fastReader, int row) {
            int ordinal = Binding.ColumnIndex;
            Type valueType = Binding.Property.ValueType;
            bool nonNull = DataReaderTypedMappingCompatibility.CanUseNonNullTypedGetter(fastReader, ordinal, valueType);
            if (!nonNull && reader.IsDBNull(ordinal)) return;
            if (nonNull || DataReaderTypedMappingCompatibility.CanUseTypedGetter(fastReader, ordinal, valueType)) {
                try {
                    (_values ??= new TValue[Capacity])[row] = _read(fastReader, ordinal);
                    _typed[row] = true;
                    return;
                } catch (InvalidCastException) {
                } catch (FormatException) {
                } catch (OverflowException) {
                } catch (NotSupportedException) {
                } catch (NotImplementedException) {
                }
            }
            // A failed typed getter is not invoked again by the fallback.
            object? value = DataReaderMappingValue.Read(reader, ordinal, valueType);
            (Fallback ??= new object?[Capacity])[row] = ReferenceEquals(value, DBNull.Value) ? null : value;
        }

        internal override T Apply(T instance, int row, CultureInfo culture,
            IReadOnlyList<string>? dateTimeFormats, DataMappingErrorValuePolicy errorValuePolicy) {
            if (!_typed[row]) return base.Apply(instance, row, culture, dateTimeFormats, errorValuePolicy);
            _assign(instance, _values![row]);
            return instance;
        }

        internal override void Resize(int capacity) {
            base.Resize(capacity);
            if (_values is not null) Array.Resize(ref _values, capacity);
            Array.Resize(ref _typed, capacity);
        }

        internal override void Clear() {
            base.Clear();
            if (_values is not null) Array.Clear(_values, 0, _values.Length);
            Array.Clear(_typed, 0, _typed.Length);
        }
    }
}
