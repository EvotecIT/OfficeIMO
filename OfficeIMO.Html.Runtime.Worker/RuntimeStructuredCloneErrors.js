(() => {
    "use strict";
    const apply=Reflect.apply, descriptor=Object.getOwnPropertyDescriptor, hasOwn=Object.hasOwn, define=Object.defineProperty;
    const string=String, TypeErrorCtor=TypeError;
    const constructors={Error,EvalError,RangeError,ReferenceError,SyntaxError,TypeError,URIError};
    const stack=descriptor(Error.prototype,'stack');
    function read(value) {
        const name=value.name, message=descriptor(value,'message');
        let text;
        if(message && hasOwn(message,'value')) {
            // String(symbol) has special behavior; serialization uses ToString.
            if(typeof message.value==='symbol')throw new TypeErrorCtor('Cannot convert a Symbol to a string');
            text=string(message.value);
        }
        // Preserve an explicit string trace without invoking a stack accessor.
        // Otherwise read the interpreter's captured stack directly.
        const ownStack=descriptor(value,'stack');
        const trace=ownStack && hasOwn(ownStack,'value') && typeof ownStack.value==='string'
            ? ownStack.value : apply(stack.get,value,[]);
        return {name:typeof name==='string' && hasOwn(constructors,name)?name:'Error',
            message:text,stack:typeof trace==='string'?trace:undefined};
    }
    function create(fields) {
        const name=hasOwn(fields,'name')&&hasOwn(constructors,fields.name)?fields.name:'Error';
        const result=new constructors[name](hasOwn(fields,'message')?fields.message:undefined);
        if(hasOwn(fields,'stack'))define(result,'stack',{value:fields.stack,writable:true,configurable:true});
        return result;
    }
    return {read,create};
})()
