((brand, maximumCharacters, errorFields) => {
    "use strict";
    const apply = Reflect.apply, descriptor = Object.getOwnPropertyDescriptor, prototype = Object.getPrototypeOf;
    const mapEntries = Map.prototype.entries, mapSet = Map.prototype.set, setValues = Set.prototype.values, setAdd = Set.prototype.add;
    const dateValue = Date.prototype.getTime, regexpSource = descriptor(RegExp.prototype, 'source').get;
    const regexpFlags = ['hasIndices','global','ignoreCase','multiline','dotAll','unicode','unicodeSets','sticky']
        .map((name,i) => [descriptor(RegExp.prototype,name)?.get, 'dgimsuvy'[i]]);
    const typedPrototype = prototype(Uint8Array.prototype);
    const typed = Object.fromEntries(['buffer','byteOffset','length'].map(name => [name,descriptor(typedPrototype,name).get]));
    const typedName = descriptor(typedPrototype,Symbol.toStringTag).get;
    const dataView = Object.fromEntries(['buffer','byteOffset','byteLength'].map(name => [name,descriptor(DataView.prototype,name).get]));
    const views = Object.fromEntries(['Int8Array','Uint8Array','Uint8ClampedArray','Int16Array','Uint16Array','Int32Array','Uint32Array','Float16Array','Float32Array','Float64Array','BigInt64Array','BigUint64Array']
        .filter(name => typeof globalThis[name] === 'function').map(name => [name,globalThis[name]]));
    function fail(message) { const error=new Error(message);error.name='DataCloneError';throw error; }
    function encode(input) {
        const seen=new Map(), nodes=[];let estimate=0;
        function charge(count){estimate+=count;if(estimate>maximumCharacters)fail('The frame message exceeds its character budget');}
        function value(item,depth) {
            if(depth>256) fail('The message exceeds its depth budget');
            const type=typeof item;
            if(type==='symbol'||type==='function') fail('The value cannot be cloned');
            if(item===null||type==='boolean'){charge(8);return item;}
            if(type==='string'){charge(item.length*6+8);return item;}
            if(type==='undefined'){charge(8);return {u:1};}
            if(type==='number') {
                charge(24);
                if(Number.isNaN(item)) return {n:'nan'};
                if(item===Infinity) return {n:'inf'};
                if(item===-Infinity) return {n:'-inf'};
                if(Object.is(item,-0)) return {n:'-0'};
                return item;
            }
            if(type==='bigint'){const text=String(item);charge(text.length+16);return {bi:text};}
            if(seen.has(item)) return {r:seen.get(item)};
            charge(64);
            const index=nodes.length;seen.set(item,index);nodes.push(null);
            let node, kind=brand(item);
            if(kind==='buffer') {
                let bytes;try{const view=new Uint8Array(item);charge(view.byteLength*4);bytes=Array.from(view);}catch(error){if(error?.name==='DataCloneError')throw error;fail('Detached buffers cannot be cloned');}
                node={t:'b',v:bytes};
            } else if(ArrayBuffer.isView(item)) {
                const name=apply(typedName,item,[]), getters=name?typed:dataView;
                let buffer,offset,length;try{buffer=apply(getters.buffer,item,[]);offset=apply(getters.byteOffset,item,[]);length=apply(name?getters.length:getters.byteLength,item,[]);}catch(_){fail('Detached views cannot be cloned');}
                node={t:'v',n:name||'DataView',b:value(buffer,depth+1),o:offset,l:length};
            } else if(kind==='date') node={t:'d',v:apply(dateValue,item,[])};
            else if(kind==='regexp') node={t:'r',s:apply(regexpSource,item,[]),f:regexpFlags.filter(([getter])=>getter&&apply(getter,item,[])).map(([,flag])=>flag).join('')};
            else if(kind==='map') {const entries=[];for(const pair of apply(mapEntries,item,[])){charge(16);entries.push([value(pair[0],depth+1),value(pair[1],depth+1)]);}node={t:'m',v:entries};}
            else if(kind==='set') {const entries=[];for(const entry of apply(setValues,item,[])){charge(8);entries.push(value(entry,depth+1));}node={t:'s',v:entries};}
            else if(kind==='array'||kind==='object') {const length=kind==='array'?descriptor(item,'length').value:undefined;if(length>maximumCharacters)fail('The frame message exceeds its character budget');const properties=[];for(const name of Object.keys(item)){charge(name.length*6+16);properties.push([name,value(item[name],depth+1)]);}node={t:kind==='array'?'a':'o',l:length,p:properties};}
            else if(kind==='error') {
                const fields=errorFields.read(item), cause=descriptor(item,'cause');
                charge((fields.message?.length || 0)*6 + (fields.stack?.length || 0)*6);
                node={t:'e',v:fields,c:cause&&'value'in cause?value(cause.value,depth+1):undefined};
            } else fail('This object type cannot be cloned');
            nodes[index]=node;return {r:index};
        }
        const wire=JSON.stringify({root:value(input,0),nodes});
        if(wire.length>maximumCharacters) { seen.clear();nodes.length=0;fail('The frame message exceeds its character budget'); }
        return wire;
    }
    function decode(wire) {
        if(typeof wire!=='string'||wire.length>maximumCharacters) fail('The frame message exceeds its character budget');
        const graph=JSON.parse(wire), nodes=graph.nodes;
        if(!graph||!Array.isArray(nodes)||nodes.length>maximumCharacters/2) fail('The frame message graph is invalid');
        const values=new Array(nodes.length);
        function primitive(item) {
            if(item===null||typeof item!=='object') return item;
            if(Object.hasOwn(item,'r')) { const index=item.r;if(!Number.isInteger(index)||index<0||index>=values.length) fail('The frame message graph is invalid');return values[index]; }
            if(item.u===1) return undefined;
            if(Object.hasOwn(item,'bi')) return BigInt(item.bi);
            if(item.n==='nan') return NaN;if(item.n==='inf')return Infinity;if(item.n==='-inf')return -Infinity;if(item.n==='-0')return -0;
            fail('The frame message graph is invalid');
        }
        for(let i=0;i<nodes.length;i++) {
            const node=nodes[i];if(!node||typeof node.t!=='string')fail('The frame message graph is invalid');
            if(node.t==='b') values[i]=new Uint8Array(node.v).buffer;
            else if(node.t==='d') values[i]=new Date(node.v);
            else if(node.t==='r') values[i]=new RegExp(node.s,node.f);
            else if(node.t==='m') values[i]=new Map();
            else if(node.t==='s') values[i]=new Set();
            else if(node.t==='a') values[i]=new Array(node.l);
            else if(node.t==='o') values[i]={};
            else if(node.t==='e') values[i]=errorFields.create(node.v);
        }
        for(let i=0;i<nodes.length;i++) if(nodes[i].t==='v') {
            const node=nodes[i], buffer=primitive(node.b), ctor=node.n==='DataView'?DataView:views[node.n];
            if(!ctor) fail('The typed array is not supported');
            values[i]=new ctor(buffer,node.o,node.l);
        }
        for(let i=0;i<nodes.length;i++) {
            const node=nodes[i], target=values[i];
            if(node.t==='m') for(const pair of node.v) apply(mapSet,target,[primitive(pair[0]),primitive(pair[1])]);
            else if(node.t==='s') for(const item of node.v) apply(setAdd,target,[primitive(item)]);
            else if(node.t==='a'||node.t==='o') for(const pair of node.p) Object.defineProperty(target,pair[0],{value:primitive(pair[1]),enumerable:true,writable:true,configurable:true});
            else if(node.t==='e'&&Object.hasOwn(node,'c')) Object.defineProperty(target,'cause',{value:primitive(node.c),writable:true,configurable:true});
        }
        return primitive(graph.root);
    }
    return {encode,decode};
})
