(function (decodeForm) {
    "use strict";
    const states = new WeakMap(), linked = new WeakMap();
    function scalar(value) {
        return String(value).replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]|[\uD800-\uDFFF]/g, part => part.length === 2 ? part : "\uFFFD");
    }
    function encode(value) { return encodeURIComponent(value).replace(/[!'()~]/g, c => "%" + c.charCodeAt(0).toString(16).toUpperCase()).replace(/%20/g, "+"); }
    function parse(value) {
        if (value.startsWith("?")) value = value.slice(1);
        return value.split("&").filter(part => part !== "").map(part => {
            const index = part.indexOf("=");
            return index < 0 ? [decodeForm(part), ""] : [decodeForm(part.slice(0,index)), decodeForm(part.slice(index+1))];
        });
    }
    function get(receiver) {
        const data = states.get(receiver);
        if (!data) throw new TypeError("Illegal invocation");
        if (data.url && data.url.search !== data.last) {
            data.last = data.url.search; data.pairs = parse(data.last);
        }
        return data;
    }
    function serialize(data) { return data.pairs.map(pair => encode(pair[0]) + "=" + encode(pair[1])).join("&"); }
    function changed(data) {
        if (data.url) { data.url.search = serialize(data); data.last = data.url.search; }
    }
    function requireArgs(count, expected) { if (count < expected) throw new TypeError("Missing URLSearchParams argument"); }
    class URLSearchParams {
        constructor(init = "") {
            const data = { pairs: [], url: null, last: "" };
            states.set(this, data);
            if (init !== null && typeof init === "object") {
                if (typeof init[Symbol.iterator] === "function") {
                    for (const item of init) {
                        if (item === null || typeof item !== "object") throw new TypeError("A query entry must be a sequence");
                        const pair = Array.from(item);
                        if (pair.length !== 2) throw new TypeError("A query entry must contain two values");
                        data.pairs.push([scalar(pair[0]),scalar(pair[1])]);
                    }
                } else { for (const name of Object.keys(init)) data.pairs.push([scalar(name),scalar(init[name])]); }
            } else { data.pairs = parse(init === null ? "" : scalar(init)); }
        }
        get size() { return get(this).pairs.length; }
        append(name,value) { requireArgs(arguments.length,2); const data=get(this); data.pairs.push([scalar(name),scalar(value)]); changed(data); }
        delete(name,value) {
            requireArgs(arguments.length,1); const data=get(this); name=scalar(name);
            const specified=value!==undefined; if(specified) value=scalar(value);
            data.pairs=data.pairs.filter(pair=>pair[0]!==name || specified && pair[1]!==value); changed(data);
        }
        get(name) { requireArgs(arguments.length,1); name=scalar(name); const pair=get(this).pairs.find(pair=>pair[0]===name); return pair ? pair[1] : null; }
        getAll(name) { requireArgs(arguments.length,1); name=scalar(name); return get(this).pairs.filter(pair=>pair[0]===name).map(pair=>pair[1]); }
        has(name,value) {
            requireArgs(arguments.length,1); name=scalar(name); const specified=value!==undefined; if(specified) value=scalar(value);
            return get(this).pairs.some(pair=>pair[0]===name && (!specified || pair[1]===value));
        }
        set(name,value) {
            requireArgs(arguments.length,2); const data=get(this); name=scalar(name); value=scalar(value);
            const index=data.pairs.findIndex(pair=>pair[0]===name);
            if(index<0) data.pairs.push([name,value]);
            else { data.pairs[index]=[name,value]; data.pairs=data.pairs.filter((pair,i)=>pair[0]!==name || i===index); }
            changed(data);
        }
        sort() { const data=get(this); data.pairs.sort((a,b)=>a[0]<b[0]?-1:a[0]>b[0]?1:0); changed(data); }
        *entries() { for(let i=0;i<get(this).pairs.length;i++) yield get(this).pairs[i].slice(); }
        *keys() { for(const pair of this.entries()) yield pair[0]; }
        *values() { for(const pair of this.entries()) yield pair[1]; }
        [Symbol.iterator]() { return this.entries(); }
        forEach(callback,receiver) { for(const pair of this.entries()) callback.call(receiver,pair[1],pair[0],this); }
        toString() { return serialize(get(this)); }
    }
    function getSearchParams() {
        let params=linked.get(this);
        if(!params) {
            params=new URLSearchParams(this.search);
            const data=get(params); data.url=this; data.last=this.search; linked.set(this,params);
        }
        get(params);
        return params;
    }
    return { URLSearchParams, getSearchParams };
})
