((brand, maximumBytes) => {
    "use strict";
    const apply = Reflect.apply, keys = Object.keys, define = Object.defineProperty;
    const descriptor = Object.getOwnPropertyDescriptor, prototype = Object.getPrototypeOf;
    const ArrayCtor = Array, MapCtor = Map, SetCtor = Set, DateCtor = Date, RegExpCtor = RegExp;
    const BufferCtor = ArrayBuffer, ViewCtor = DataView, BytesCtor = Uint8Array;
    const mapEntries = Map.prototype.entries, mapSet = Map.prototype.set, setValues = Set.prototype.values, setAdd = Set.prototype.add;
    const mapHas = Map.prototype.has, mapGet = Map.prototype.get, from = Array.from, hasOwn = Object.hasOwn;
    const mapSize = descriptor(Map.prototype,'size').get, setSize = descriptor(Set.prototype,'size').get;
    const dateValue = Date.prototype.getTime, regexpSource = descriptor(RegExp.prototype, 'source').get;
    const regexpFlags = ['hasIndices','global','ignoreCase','multiline','dotAll','unicode','unicodeSets','sticky']
        .map((name,i) => [descriptor(RegExp.prototype,name)?.get, 'dgimsuvy'[i]]);
    const bufferLength = descriptor(ArrayBuffer.prototype,'byteLength').get;
    const bufferResizable = descriptor(ArrayBuffer.prototype,'resizable')?.get;
    const typedPrototype = prototype(Uint8Array.prototype);
    const typed = Object.fromEntries(['buffer','byteOffset','length'].map(name => [name,descriptor(typedPrototype,name).get]));
    const typedName = descriptor(typedPrototype,Symbol.toStringTag).get;
    const views = Object.fromEntries(['Int8Array','Uint8Array','Uint8ClampedArray','Int16Array','Uint16Array','Int32Array','Uint32Array','Float16Array','Float32Array','Float64Array','BigInt64Array','BigUint64Array']
        .filter(name => typeof globalThis[name] === 'function').map(name => [name,globalThis[name]]));
    const dataView = Object.fromEntries(['buffer','byteOffset','byteLength'].map(name => [name,descriptor(DataView.prototype,name).get]));
    const isView = ArrayBuffer.isView;
    const boxes = [Boolean,Number,String,BigInt].map(type => type.prototype.valueOf);
    const objectBox = Object;
    const errors = {Error,EvalError,RangeError,ReferenceError,SyntaxError,TypeError,URIError};
    function fail(message) { const error=new errors.Error(message);error.name='DataCloneError';throw error; }
    return function snapshot(input) {
        const seen = new MapCtor();
        let bytes = 0;
        function charge(count) { bytes += count; if (bytes > maximumBytes) fail('History state exceeds its byte budget'); }
        function copy(value, depth) {
            if (depth > 256) fail('History state exceeds its depth budget');
            if (typeof value === 'symbol' || typeof value === 'function') fail('The value cannot be cloned');
            if (value === null || typeof value !== 'object') {
                charge(typeof value === 'string' ? value.length * 2 : typeof value === 'bigint' ? value.toString().length * 2 : 16);
                return value;
            }
            if (apply(mapHas,seen,[value])) return apply(mapGet,seen,[value]);
            charge(64);
            let result, kind = brand(value);
            if (kind === 'buffer') {
                const length = apply(bufferLength,value,[]);
                if (bufferResizable && apply(bufferResizable,value,[])) fail('Resizable buffers are not supported in history state');
                charge(length);
                try {
                    const original = new BytesCtor(value);
                    result = new BufferCtor(length);
                    const target = new BytesCtor(result);
                    for(let i=0;i<length;i++) target[i]=original[i];
                } catch (_) { fail('Detached buffers cannot be cloned'); }
            } else if (isView(value)) {
                const name = apply(typedName,value,[]);
                const getters = name ? typed : dataView;
                let buffer, offset, length;
                try {
                    buffer = apply(getters.buffer,value,[]);
                    offset = apply(getters.byteOffset,value,[]);
                    length = apply(name ? getters.length : getters.byteLength,value,[]);
                } catch (_) { fail('Detached or out-of-bounds views cannot be cloned'); }
                result = new (name ? views[name] : ViewCtor)(copy(buffer,depth+1),offset,length);
            } else if (kind === 'date') result = new DateCtor(apply(dateValue,value,[]));
            else if (kind === 'regexp') {
                const source = apply(regexpSource,value,[]);
                charge(source.length*2);
                result = new RegExpCtor(source,regexpFlags.filter(([getter])=>getter && apply(getter,value,[])).map(([,flag])=>flag).join(''));
            } else if (kind === 'map') result = new MapCtor();
            else if (kind === 'set') result = new SetCtor();
            else if (kind === 'array') {const length=descriptor(value,'length').value;charge(length*8);result = new ArrayCtor(length);}
            else if (kind === 'object') result = {};
            else if (kind === 'error') {
                const name = value.name;
                const message = descriptor(value,'message');
                const text = message && 'value' in message ? String(message.value) : undefined;
                charge(text?.length*2 || 0);
                result = new (typeof name==='string' && hasOwn(errors,name) ? errors[name] : errors.Error)(text);
            } else {
                for (const getter of boxes) {
                    try { const primitive = apply(getter,value,[]); result = objectBox(copy(primitive,depth+1)); break; }
                    catch (error) { if (error?.name === 'DataCloneError') throw error; }
                }
                if (!result) fail('This object type cannot be cloned');
            }
            apply(mapSet,seen,[value,result]);
            if (kind === 'map') {
                charge(apply(mapSize,value,[])*32);
                const entries = from(apply(mapEntries,value,[]));
                for(const [key,item] of entries) apply(mapSet,result,[copy(key,depth+1),copy(item,depth+1)]);
            } else if (kind === 'set') {
                charge(apply(setSize,value,[])*16);
                const entries = from(apply(setValues,value,[]));
                for(const item of entries) apply(setAdd,result,[copy(item,depth+1)]);
            } else if (kind === 'array' || kind === 'object') {
                const names = keys(value);
                charge(names.length*16);
                for(const name of names) {
                    const property = descriptor(value,name);
                    if (!property?.enumerable) continue;
                    charge(name.length*2);
                    define(result,name,{value:copy(value[name],depth+1),enumerable:true,writable:true,configurable:true});
                }
            } else if (kind === 'error') {
                const cause = descriptor(value,'cause');
                if (cause && 'value' in cause) define(result,'cause',{value:copy(cause.value,depth+1),writable:true,configurable:true});
            }
            return result;
        }
        return {value:copy(input,0),bytes};
    };
})
