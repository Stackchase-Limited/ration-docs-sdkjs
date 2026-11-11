/*
 * Ration Docs: separate JavaScript realm for PDF form/annotation scripts.
 *
 * PDF scripts come from the document, so they are untrusted. They run here inside
 * QuickJS (compiled to WebAssembly), not in the editor's own V8 realm: a script has
 * no way to name or reach `window`, `AscDesktopEditor`, the sdkjs objects or any
 * constructor of the editor realm, because none of them exist inside QuickJS.
 *
 * The only connection to the editor is one host function. Only JSON text crosses it:
 * primitives, arrays (copied) and numeric handles. A handle stands for an editor object
 * (the event, a field, the document) or for a function the editor allows; every
 * property read/write and call on a handle goes back to `hooks` (pdf/src/forms/formActions.js),
 * which decides what the script may touch.
 *
 * Built into pdf/src/engine/pdfscripting.js by ./build.sh (esbuild, IIFE). Third-party code:
 * QuickJS (MIT, Fabrice Bellard, Charlie Gordon) and quickjs-emscripten (MIT, Jake Teton-Landis).
 */

import { newQuickJSWASMModuleFromVariant } from "quickjs-emscripten-core";
import variant from "@jitl/quickjs-singlefile-browser-release-sync";

const MEMORY_LIMIT  = 64 * 1024 * 1024;
const STACK_LIMIT   = 128 * 1024;
const TIME_LIMIT_MS = 3000;
const MAX_DEPTH     = 32;

// Runs once inside every QuickJS context. Keeps the host channel in a closure and removes
// it from the global object; installs the API globals and the `__rdRun` entry point.
const GUEST_BOOTSTRAP = `(function() {
"use strict";
const host = globalThis.__rdHost;
delete globalThis.__rdHost;
const proxies = new Map();
const proxyIds = new WeakMap();
const expandos = new Map();
const stringify = JSON.stringify, parse = JSON.parse;

function enc(v, depth) {
    depth = depth || 0;
    if (depth > ${MAX_DEPTH}) throw new RangeError("value is nested too deeply");
    if (v === undefined) return { "u": 1 };
    if (v === null || typeof v === "boolean" || typeof v === "string") return v;
    if (typeof v === "number") return isFinite(v) ? v : { "n": String(v) };
    if (typeof v === "bigint") return Number(v);
    if (proxyIds.has(v)) return { "h": proxyIds.get(v) };
    if (Array.isArray(v)) return { "a": v.map(function(x) { return enc(x, depth + 1); }) };
    if (typeof v === "object") {
        if (v instanceof Date) return { "t": v.getTime() };
        const o = {};
        Object.keys(v).forEach(function(k) { o[k] = enc(v[k], depth + 1); });
        return { "d": o };
    }
    if (typeof v === "symbol") return String(v.description);
    return { "u": 1 };
}
function dec(v) {
    if (v === null || typeof v !== "object") return v;
    if (v["u"] === 1) return undefined;
    if ("n" in v) return Number(v["n"]);
    if ("a" in v) return v["a"].map(dec);
    if ("o" in v) return wrap(v["o"], false);
    if ("f" in v) return wrap(v["f"], true);
    return undefined;
}
function request(msg) {
    return dec(parse(host(stringify(msg))));
}
function wrap(id, isFunc) {
    let p = proxies.get(id);
    if (p) return p;
    const own = Object.create(null);
    expandos.set(id, own);
    const target = isFunc ? function() {} : {};
    p = new Proxy(target, {
        get: function(t, key) {
            if (typeof key !== "string") return undefined;
            if (key in own) return own[key];
            return request(["get", id, key]);
        },
        set: function(t, key, value) {
            if (typeof key !== "string") return false;
            // a property the API does not have stays on this side, like an expando
            if (!request(["set", id, key, enc(value)])) own[key] = value;
            return true;
        },
        has: function(t, key) {
            return typeof key === "string" && ((key in own) || request(["get", id, key]) !== undefined);
        },
        deleteProperty: function(t, key) { delete own[key]; return true; },
        apply: function(t, self, args) {
            return request(["call", id, args.map(function(a) { return enc(a); })]);
        },
        construct: function() { throw new TypeError("not a constructor"); },
        defineProperty: function() { return false; },
        setPrototypeOf: function() { return false; },
        ownKeys: function(t) { return Reflect.ownKeys(t); },
        getOwnPropertyDescriptor: function(t, key) { return Reflect.getOwnPropertyDescriptor(t, key); }
    });
    proxies.set(id, p);
    proxyIds.set(p, id);
    return p;
}
function deepFreeze(v) {
    if (v && typeof v === "object") { Object.keys(v).forEach(function(k) { deepFreeze(v[k]); }); Object.freeze(v); }
    return v;
}

globalThis.__rdInstall = function(sGlobals, sConstants, sGetters) {
    const g = parse(sGlobals), c = parse(sConstants), getters = parse(sGetters);
    Object.keys(g).forEach(function(name) { globalThis[name] = dec(g[name]); });
    Object.keys(c).forEach(function(name) {
        globalThis[name] = deepFreeze(c[name]);
    });
    // objects whose properties return a fresh copy on every read (e.g. color.red)
    Object.keys(getters).forEach(function(name) {
        const src = getters[name], obj = {};
        Object.keys(src).forEach(function(k) {
            Object.defineProperty(obj, k, { enumerable: true, get: function() { return parse(stringify(src[k])); } });
        });
        globalThis[name] = Object.freeze(obj);
    });
    delete globalThis.__rdInstall;
};
const run = function(sScript, sThis, sEvent) {
    const fn = new Function("event", sScript);
    fn.call(dec(parse(sThis)), dec(parse(sEvent)));
};
Object.defineProperty(globalThis, "__rdRun", { value: run, writable: false, configurable: false, enumerable: false });
})();`;

let modulePromise = null;
let quickjs = null;

function load() {
    if (!modulePromise) {
        modulePromise = newQuickJSWASMModuleFromVariant(variant).then(function(m) {
            quickjs = m;
            return true;
        });
    }
    return modulePromise;
}

function isReady() {
    return quickjs !== null;
}

/**
 * Creates a realm. hooks = {
 *   get(obj, key)          -> value (undefined when not allowed)
 *   set(obj, key, value)   -> true when the API handled it, false otherwise
 *   call(fn, self, args)   -> value
 *   exportValue(value)     -> true when the value may be handed to the script as a handle
 * }
 * globals: name -> host value (functions/objects become handles);
 * constants: name -> plain JSON data; getters: name -> {key: JSON data}, read as fresh copies.
 */
function createRealm(hooks, globals, constants, getters) {
    if (!quickjs)
        throw new Error("PDF scripting engine is not loaded");

    const runtime = quickjs.newRuntime();
    runtime.setMemoryLimit(MEMORY_LIMIT);
    runtime.setMaxStackSize(STACK_LIMIT);
    const vm = runtime.newContext();

    // handle tables live on the editor side only; the script sees the numbers
    const objToId = new Map();
    const idToEntry = [null];
    const methodIds = new Map();

    function handleFor(value, self, key) {
        if (typeof value === "function" && self !== undefined) {
            let byKey = methodIds.get(self);
            if (!byKey) methodIds.set(self, byKey = new Map());
            let id = byKey.get(key);
            if (id === undefined || idToEntry[id].fn !== value) {
                id = idToEntry.length;
                idToEntry.push({ fn: value, self: self });
                byKey.set(key, id);
            }
            return id;
        }
        let id = objToId.get(value);
        if (id === undefined) {
            id = idToEntry.length;
            idToEntry.push(typeof value === "function" ? { fn: value, self: undefined } : { obj: value });
            objToId.set(value, id);
        }
        return id;
    }

    function enc(v, self, key, depth) {
        depth = depth || 0;
        if (depth > MAX_DEPTH) throw new RangeError("value is nested too deeply");
        if (v === undefined) return { "u": 1 };
        if (v === null || typeof v === "boolean" || typeof v === "string") return v;
        if (typeof v === "number") return isFinite(v) ? v : { "n": String(v) };
        if (Array.isArray(v)) return { "a": v.map(function(x) { return enc(x, undefined, undefined, depth + 1); }) };
        if ((typeof v === "object" || typeof v === "function") && hooks["exportValue"](v)) {
            if (typeof v === "function") return { "f": handleFor(v, self, key) };
            return { "o": handleFor(v) };
        }
        return { "u": 1 };
    }
    function dec(v, depth) {
        depth = depth || 0;
        if (depth > MAX_DEPTH) throw new RangeError("value is nested too deeply");
        if (v === null || typeof v !== "object") return v;
        if (v["u"] === 1) return undefined;
        if ("n" in v) return Number(v["n"]);
        if ("t" in v) return new Date(Number(v["t"]));
        if ("a" in v) return v["a"].map(function(x) { return dec(x, depth + 1); });
        if ("h" in v) {
            // editor functions are never handed back to the editor as values
            const entry = idToEntry[v["h"]];
            return entry ? entry.obj : undefined;
        }
        if ("d" in v) {
            const o = {};
            Object.keys(v["d"]).forEach(function(k) {
                if (k !== "__proto__")
                    o[k] = dec(v["d"][k], depth + 1);
            });
            return o;
        }
        return undefined;
    }
    function entryOf(id) {
        const entry = typeof id === "number" ? idToEntry[id] : null;
        if (!entry) throw new TypeError("invalid handle");
        return entry;
    }

    function onRequest(sMsg) {
        const msg = JSON.parse(sMsg);
        const op = msg[0], entry = entryOf(msg[1]);
        if (op === "get") {
            if (entry.obj === undefined || typeof msg[2] !== "string") return enc(undefined);
            return enc(hooks["get"](entry.obj, msg[2]), entry.obj, msg[2]);
        }
        if (op === "set") {
            if (entry.obj === undefined || typeof msg[2] !== "string") return false;
            return !!hooks["set"](entry.obj, msg[2], dec(msg[3]));
        }
        if (op === "call") {
            if (entry.fn === undefined || !Array.isArray(msg[2])) throw new TypeError("not a function");
            return enc(hooks["call"](entry.fn, entry.self, msg[2].map(function(a) { return dec(a); })));
        }
        throw new TypeError("unknown request");
    }

    const hostFn = vm.newFunction("__rdHost", function(hMsg) {
        let sOut;
        try {
            sOut = JSON.stringify(onRequest(vm.getString(hMsg)));
        }
        catch (err) {
            return { error: vm.newError(String(err && err.message || err)) };
        }
        return vm.newString(sOut);
    });
    vm.setProp(vm.global, "__rdHost", hostFn);
    hostFn.dispose();

    // a script can start another one (setting a field value runs its calculate scripts);
    // the outermost one owns the deadline
    let evalDepth = 0;
    function evalOrThrow(code) {
        if (evalDepth === 0) {
            const deadline = Date.now() + TIME_LIMIT_MS;
            runtime.setInterruptHandler(function() { return Date.now() > deadline; });
        }
        evalDepth++;
        let message = null;
        try {
            const res = vm.evalCode(code);
            if (res.error) {
                // dump may run script getters, so the deadline still applies
                let e;
                try { e = vm.dump(res.error); } catch (err) { e = "uncaught exception"; }
                res.error.dispose();
                message = e && typeof e === "object" ? (e.name ? e.name + ": " : "") + e.message : String(e);
            }
            else {
                res.value.dispose();
            }
        }
        finally {
            if (--evalDepth === 0)
                runtime.removeInterruptHandler();
        }
        if (message !== null)
            throw new Error("PDF script: " + message);
    }
    function literal(s) {
        // JSON text is a valid string literal for QuickJS (ES2019 allows U+2028/U+2029 in strings)
        return JSON.stringify(s);
    }

    evalOrThrow(GUEST_BOOTSTRAP);
    const encGlobals = {};
    Object.keys(globals || {}).forEach(function(name) { encGlobals[name] = enc(globals[name]); });
    evalOrThrow("__rdInstall(" + literal(JSON.stringify(encGlobals)) + "," + literal(JSON.stringify(constants || {})) + "," + literal(JSON.stringify(getters || {})) + ")");

    let disposed = false;
    return {
        "run": function(sScript, oThis, oEvent) {
            if (disposed) throw new Error("PDF script realm is disposed");
            evalOrThrow("__rdRun(" + literal(String(sScript)) + "," + literal(JSON.stringify(enc(oThis))) + "," + literal(JSON.stringify(enc(oEvent))) + ")");
        },
        "dispose": function() {
            if (disposed) return;
            disposed = true;
            vm.dispose();
            runtime.dispose();
        }
    };
}

const api = { "load": load, "isReady": isReady, "createRealm": createRealm };
globalThis["AscPDFScriptSandbox"] = api;
