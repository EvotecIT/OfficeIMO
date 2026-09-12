((currentUrl, parseRoute, updateRoute, enqueue, snapshot, maximumEntries, maximumBytes, maximumTasks, dispatch) => {
    "use strict";
    const URLCtor = URL, EventCtor = Event, ErrorCtor = Error, define = Object.defineProperty;
    const apply=Reflect.apply, stringIndex=String.prototype.indexOf, stringSlice=String.prototype.slice, setPrototype=Object.setPrototypeOf;
    const clone = value => snapshot(value).value;
    let entries = setPrototype([{url:currentUrl(),snapshot:{value:null,bytes:0},scroll:'auto'}],null);
    let index = 0, state = null, pending = 0;
    function fail(name,message) { const error=new ErrorCtor(message);error.name=name;throw error; }
    function check(receiver,target) { if(receiver !== target) throw new TypeError('Illegal invocation'); }
    function checkQueue() { if(pending>=maximumTasks)fail('QuotaExceededError','History task queue exceeds its budget'); }
    function queue(callback) {checkQueue();pending++;enqueue(()=>{pending--;callback()});}
    function PopStateEvent(type,options={}) {
        if(!new.target||!arguments.length)throw new TypeError('An event constructor call and type are required');
        return popEvent(text(type),options??{},options?.state ?? null);
    }
    function popEvent(type,options,state) {
        const value=new EventCtor(text(type),options??{});Object.setPrototypeOf(value,PopStateEvent.prototype);
        define(value,'state',{value:state,enumerable:true});return value;
    }
    function HashChangeEvent(type,options={}) {
        if(!new.target||!arguments.length)throw new TypeError('An event constructor call and type are required');
        const value=new EventCtor(text(type),options??{});Object.setPrototypeOf(value,HashChangeEvent.prototype);
        for(const key of ['oldURL','newURL'])define(value,key,{value:text(options?.[key] ?? ''),enumerable:true});return value;
    }
    Object.setPrototypeOf(PopStateEvent.prototype,EventCtor.prototype);
    Object.setPrototypeOf(HashChangeEvent.prototype,EventCtor.prototype);
    function event(type,properties) {
        const value = type==='popstate' ? popEvent(type,{},properties.state) : new HashChangeEvent(type,properties);
        dispatch(value);
    }
    function withoutFragment(url) { const offset=apply(stringIndex,url,['#']);return offset<0 ? url : apply(stringSlice,url,[0,offset]); }
    function fragment(url) { const offset=apply(stringIndex,url,['#']);return offset<0 ? null : apply(stringSlice,url,[offset+1]); }
    function text(value) {if(typeof value==='symbol')throw new TypeError('Cannot convert a Symbol to a string');return String(value);}
    function commit(url,saved,replace) {
        const entry={url,snapshot:saved,scroll:entries[index].scroll};
        const next=setPrototype([],null);
        for(let i=0;i<(replace?entries.length:index+1);i++)next[i]=entries[i];
        let nextIndex=replace ? index : next.length;
        next[nextIndex]=entry;
        let bytes=0;for(let i=0;i<next.length;i++)bytes+=next[i].snapshot.bytes;
        while(next.length>maximumEntries || bytes>maximumBytes) {
            const remove=nextIndex===1 ? 2 : 1;
            if(remove>=next.length)fail('QuotaExceededError','History cannot retain its first and current state within the byte budget');
            bytes-=next[remove].snapshot.bytes;
            for(let i=remove;i<next.length-1;i++)next[i]=next[i+1];next.length--;
            if(remove<nextIndex)nextIndex--;
        }
        const nextState=clone(saved.value);
        entries=next;index=nextIndex;state=nextState;updateRoute(url);
    }
    const history = {};
    define(history,Symbol.toStringTag,{value:'History'});
    define(history,'length',{get(){check(this,history);return entries.length},enumerable:true});
    define(history,'state',{get(){check(this,history);return state},enumerable:true});
    define(history,'scrollRestoration',{get(){check(this,history);return entries[index].scroll},set(value){check(this,history);value=text(value);if(value==='auto'||value==='manual')entries[index].scroll=value},enumerable:true});
    function change(receiver,data,unused,url,replace,count) {
        check(receiver,history);
        if(count<2) throw new TypeError('History state requires data and an unused title');
        text(unused);
        if(url!==undefined && url!==null) url=text(url);
        const saved=snapshot(data);
        commit(url===undefined||url===null||url==='' ? currentUrl() : parseRoute(url),saved,replace);
    }
    history.pushState=function(data,unused,url){change(this,data,unused,url,false,arguments.length)};
    history.replaceState=function(data,unused,url){change(this,data,unused,url,true,arguments.length)};
    function traverse(delta){
        if(delta===0) fail('NotSupportedError','Document reload is outside this session profile');
        queue(()=>{
            const next=index+delta;
            if(next<0||next>=entries.length)return;
            const oldUrl=currentUrl(),entry=entries[next],nextState=clone(entry.snapshot.value);
            index=next;state=nextState;updateRoute(entry.url);
            if(fragment(oldUrl)!==fragment(entry.url))queue(()=>event('hashchange',{oldURL:oldUrl,newURL:entry.url}));
            event('popstate',{state});
        });
    }
    history.go=function(delta=0){check(this,history);traverse(+delta|0)};
    history.back=function(){check(this,history);traverse(-1)};
    history.forward=function(){check(this,history);traverse(1)};
    function navigate(value,replace) {
        const url=parseRoute(text(value)),oldUrl=currentUrl();
        if(withoutFragment(url)!==withoutFragment(oldUrl) || fragment(url)===null)
            fail('NotSupportedError','Cross-document navigation is outside this session profile');
        if(url===oldUrl)return;
        checkQueue();
        commit(url,{value:null,bytes:0},replace);
        queue(()=>event('hashchange',{oldURL:oldUrl,newURL:url}));
        event('popstate',{state});
    }
    const location={};
    define(location,Symbol.toStringTag,{value:'Location'});
    for(const name of ['href','origin','protocol','host','hostname','port','pathname','search','hash']) {
        define(location,name,{enumerable:true,get(){check(this,location);return new URLCtor(currentUrl())[name]},
            ...(name==='origin'?{}:{set(value){check(this,location);value=text(value);if(name==='href'){navigate(value,false);return;}if(name==='hash'&&(value===''||value==='#')){navigate(withoutFragment(currentUrl())+'#',false);return;}const url=new URLCtor(currentUrl());url[name]=value;navigate(url.href,false)}})});
    }
    location.assign=function(url){check(this,location);if(!arguments.length)throw new TypeError('URL required');navigate(url,false)};
    location.replace=function(url){check(this,location);if(!arguments.length)throw new TypeError('URL required');navigate(url,true)};
    location.reload=function(){check(this,location);fail('NotSupportedError','Document reload is outside this session profile')};
    location.toString=function(){check(this,location);return currentUrl()};
    function History(){throw new TypeError('Illegal constructor');}
    function Location(){throw new TypeError('Illegal constructor');}
    Object.defineProperties(History.prototype,Object.getOwnPropertyDescriptors(history));
    Object.defineProperties(Location.prototype,Object.getOwnPropertyDescriptors(location));
    Object.setPrototypeOf(history,History.prototype);Object.setPrototypeOf(location,Location.prototype);
    return {history,location,navigate,History,Location,PopStateEvent,HashChangeEvent};
})
