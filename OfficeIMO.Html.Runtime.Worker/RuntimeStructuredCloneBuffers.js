(isLengthTracking => {
    "use strict";
    const apply = Reflect.apply, descriptor = Object.getOwnPropertyDescriptor, define = Object.defineProperty;
    const BufferCtor = ArrayBuffer, BytesCtor = Uint8Array, ViewCtor = DataView;
    const isView = ArrayBuffer.isView, ArrayCtor = Array;
    const resizable = descriptor(BufferCtor.prototype, 'resizable').get;
    const maximum = descriptor(BufferCtor.prototype, 'maxByteLength').get;
    const typedPrototype = Object.getPrototypeOf(BytesCtor.prototype);
    const typed = Object.fromEntries(['buffer','byteOffset','length'].map(name => [name,descriptor(typedPrototype,name).get]));
    const typedName = descriptor(typedPrototype, Symbol.toStringTag).get;
    const typedValues = typedPrototype.values;
    const dataView = Object.fromEntries(['buffer','byteOffset','byteLength'].map(name => [name,descriptor(ViewCtor.prototype,name).get]));
    const views = Object.fromEntries(['Int8Array','Uint8Array','Uint8ClampedArray','Int16Array','Uint16Array','Int32Array','Uint32Array','Float16Array','Float32Array','Float64Array','BigInt64Array','BigUint64Array']
        .filter(name => typeof globalThis[name] === 'function').map(name => [name,globalThis[name]]));
    function readBuffer(value) {
        // Constructing the view validates detachment without reading authored properties.
        const bytes = new BytesCtor(value);
        return {bytes, length:apply(typed.length,bytes,[]), maximum:apply(resizable,value,[]) ? apply(maximum,value,[]) : undefined};
    }
    function createBuffer(bytes, maximumLength) {
        const length = isView(bytes) ? apply(typed.length,bytes,[]) : bytes.length;
        const buffer = maximumLength === undefined ? new BufferCtor(length) : new BufferCtor(length, {maxByteLength:maximumLength});
        const target = new BytesCtor(buffer);
        for(let i=0;i<length;i++) target[i]=bytes[i];
        return buffer;
    }
    function readView(value) {
        const name = apply(typedName,value,[]), getters = name ? typed : dataView;
        // Typed-array length/offset getters return zero for out-of-bounds views. Validate first.
        if(name) apply(typedValues,value,[]);
        return {name:name || 'DataView', buffer:apply(getters.buffer,value,[]),
            offset:apply(getters.byteOffset,value,[]),
            length:apply(name ? getters.length : getters.byteLength,value,[]), tracking:isLengthTracking(value)};
    }
    function createView(fields, buffer) {
        const ctor = fields.name === 'DataView' ? ViewCtor : views[fields.name];
        return fields.tracking ? new ctor(buffer,fields.offset) : new ctor(buffer,fields.offset,fields.length);
    }
    function bytesToArray(bytes) {
        const length = apply(typed.length,bytes,[]), result = new ArrayCtor(length);
        for(let i=0;i<length;i++) define(result,i,{value:bytes[i],writable:true,enumerable:true,configurable:true});
        return result;
    }
    return {isView, readBuffer, createBuffer, readView, createView, bytesToArray};
})
