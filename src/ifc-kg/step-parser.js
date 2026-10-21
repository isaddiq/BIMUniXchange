/* BIMUniXchange · BIM Knowledge Graph Studio
 * step-parser.js — streaming-friendly ISO 10303-21 (STEP physical file) parser.
 *
 * Pass 1 indexes every instance (#id = TYPE(args);) by byte range without
 * parsing its arguments. Arguments are decoded lazily and cached, so geometry
 * (typically >80% of an IFC file) is never materialised.
 */
(function (KG) {
  "use strict";

  class Ref {
    constructor(id) { this.id = id; }
  }
  class Enum {
    constructor(v) { this.v = v; }
  }
  class Typed {
    constructor(type, value) { this.type = type; this.value = value; }
  }
  const DERIVED = Object.freeze({ derived: true });

  /** Decode STEP string escapes: '' , \\ , \S\x , \X\hh , \X2\hhhh…\X0\ , \X4\hhhhhhhh…\X0\ */
  function decodeString(s) {
    if (s.indexOf("\\") < 0) return s;
    let out = "";
    let i = 0;
    while (i < s.length) {
      const c = s[i];
      if (c === "\\" && s[i + 1] === "X" && s[i + 2] === "2" && s[i + 3] === "\\") {
        i += 4;
        const end = s.indexOf("\\X0\\", i);
        const hex = s.slice(i, end < 0 ? s.length : end);
        for (let k = 0; k + 4 <= hex.length; k += 4) out += String.fromCharCode(parseInt(hex.substr(k, 4), 16));
        i = end < 0 ? s.length : end + 4;
      } else if (c === "\\" && s[i + 1] === "X" && s[i + 2] === "4" && s[i + 3] === "\\") {
        i += 4;
        const end = s.indexOf("\\X0\\", i);
        const hex = s.slice(i, end < 0 ? s.length : end);
        for (let k = 0; k + 8 <= hex.length; k += 8) out += String.fromCodePoint(parseInt(hex.substr(k, 8), 16));
        i = end < 0 ? s.length : end + 4;
      } else if (c === "\\" && s[i + 1] === "X" && s[i + 2] === "\\") {
        out += String.fromCharCode(parseInt(s.substr(i + 3, 2), 16));
        i += 5;
      } else if (c === "\\" && s[i + 1] === "S" && s[i + 2] === "\\") {
        out += String.fromCharCode(s.charCodeAt(i + 3) + 128);
        i += 4;
      } else if (c === "\\" && s[i + 1] === "P" && s[i + 3] === "\\") {
        i += 4; // code page switch – ignored
      } else if (c === "\\" && s[i + 1] === "\\") {
        out += "\\";
        i += 2;
      } else {
        out += c;
        i++;
      }
    }
    return out;
  }

  /** Parse an argument list string (without outer parentheses) into JS values. */
  function parseArgs(str) {
    let i = 0;
    const n = str.length;

    function ws() {
      while (i < n) {
        const c = str.charCodeAt(i);
        if (c === 32 || c === 10 || c === 13 || c === 9) i++;
        else if (c === 47 && str.charCodeAt(i + 1) === 42) {
          const e = str.indexOf("*/", i + 2);
          i = e < 0 ? n : e + 2;
        } else break;
      }
    }

    function value() {
      ws();
      const c = str[i];
      if (c === "'") {
        i++;
        let s = "";
        let start = i;
        while (i < n) {
          if (str[i] === "'") {
            if (str[i + 1] === "'") { s += str.slice(start, i + 1); i += 2; start = i; continue; }
            break;
          }
          i++;
        }
        s += str.slice(start, i);
        i++;
        return decodeString(s);
      }
      if (c === "#") {
        i++;
        const st = i;
        while (i < n && str.charCodeAt(i) >= 48 && str.charCodeAt(i) <= 57) i++;
        return new Ref(+str.slice(st, i));
      }
      if (c === "$") { i++; return null; }
      if (c === "*") { i++; return DERIVED; }
      if (c === ".") {
        const e = str.indexOf(".", i + 1);
        const v = str.slice(i + 1, e);
        i = e + 1;
        return new Enum(v);
      }
      if (c === "(") {
        i++;
        const arr = [];
        ws();
        if (str[i] === ")") { i++; return arr; }
        while (i < n) {
          arr.push(value());
          ws();
          if (str[i] === ",") { i++; continue; }
          if (str[i] === ")") { i++; break; }
          i++; // tolerate garbage
        }
        return arr;
      }
      if (c === '"') {
        const e = str.indexOf('"', i + 1);
        const v = str.slice(i + 1, e);
        i = e + 1;
        return v; // binary – keep hex string
      }
      // Typed value: IFCLABEL('x')
      if ((c >= "A" && c <= "Z") || (c >= "a" && c <= "z")) {
        const st = i;
        while (i < n && /[A-Za-z0-9_]/.test(str[i])) i++;
        const type = str.slice(st, i).toUpperCase();
        ws();
        if (str[i] === "(") {
          i++;
          const inner = [];
          ws();
          while (i < n && str[i] !== ")") {
            inner.push(value());
            ws();
            if (str[i] === ",") i++;
            ws();
          }
          i++;
          return new Typed(type, inner.length === 1 ? inner[0] : inner);
        }
        return new Enum(type);
      }
      // Number
      const st = i;
      while (i < n && /[-+0-9.eE]/.test(str[i])) i++;
      const num = str.slice(st, i);
      if (!num) { i++; return null; }
      return parseFloat(num);
    }

    const out = [];
    ws();
    if (i >= n) return out;
    while (i < n) {
      out.push(value());
      ws();
      if (str[i] === ",") { i++; continue; }
      if (i >= n) break;
      i++;
    }
    return out;
  }

  class StepFile {
    constructor(text) {
      this.text = text;
      this.ids = [];
      this.types = [];
      this.starts = [];
      this.ends = [];
      this.indexOf = new Map(); // id → position in arrays
      this.byType = new Map(); // TYPE → [ids]
      this.cache = new Map();
      this.header = {};
      this.schema = "";
      this.errors = 0;
    }

    /** Index the file. `onProgress(fraction)` may be async-awaited for UI yielding. */
    async index(onProgress) {
      const t = this.text;
      const n = t.length;
      this.parseHeader();
      let i = t.indexOf("DATA;");
      i = i < 0 ? 0 : i + 5;
      let nextYield = i + 4_000_000;
      const typeIntern = new Map();

      while (i < n) {
        // seek '#' at statement start (skip whitespace / comments)
        let c = t.charCodeAt(i);
        if (c === 35 /* # */) {
          let j = i + 1;
          let id = 0;
          while (j < n) {
            const d = t.charCodeAt(j);
            if (d < 48 || d > 57) break;
            id = id * 10 + (d - 48);
            j++;
          }
          while (j < n && t.charCodeAt(j) <= 32) j++;
          if (t.charCodeAt(j) !== 61 /* = */) { i = j; continue; }
          j++;
          while (j < n && t.charCodeAt(j) <= 32) j++;
          const ts = j;
          while (j < n) {
            const d = t.charCodeAt(j);
            if ((d >= 65 && d <= 90) || (d >= 97 && d <= 122) || (d >= 48 && d <= 57) || d === 95) j++;
            else break;
          }
          let type = t.slice(ts, j).toUpperCase();
          while (j < n && t.charCodeAt(j) <= 32) j++;
          if (t.charCodeAt(j) !== 40 /* ( */) {
            // malformed instance – skip to the end of the statement
            const e = t.indexOf(";", j);
            this.errors++;
            i = e < 0 ? n : e + 1;
            continue;
          }
          // complex entity instance "#1=(A()B());" has no leading type name
          if (!type) type = "COMPLEX";
          const as = j + 1;
          // scan to matching close paren
          let depth = 0;
          let inStr = false;
          let k = j;
          for (; k < n; k++) {
            const d = t.charCodeAt(k);
            if (inStr) {
              if (d === 39) inStr = false; // '' escape toggles twice
              continue;
            }
            if (d === 39) inStr = true;
            else if (d === 40) depth++;
            else if (d === 41) {
              depth--;
              if (depth === 0) break;
            } else if (d === 59 && depth <= 0) break;
          }
          const ae = k;
          // advance to ';'
          let semi = t.indexOf(";", k);
          if (semi < 0) semi = n;
          if (!typeIntern.has(type)) typeIntern.set(type, type);
          type = typeIntern.get(type);
          this.indexOf.set(id, this.ids.length);
          this.ids.push(id);
          this.types.push(type);
          this.starts.push(as);
          this.ends.push(ae);
          let arr = this.byType.get(type);
          if (!arr) this.byType.set(type, (arr = []));
          arr.push(id);
          i = semi + 1;
          if (i > nextYield && onProgress) {
            nextYield = i + 4_000_000;
            await onProgress(i / n);
          }
        } else if (c === 47 && t.charCodeAt(i + 1) === 42) {
          const e = t.indexOf("*/", i + 2);
          i = e < 0 ? n : e + 2;
        } else if (c === 69 /* E */ && t.startsWith("ENDSEC", i)) {
          // There can be several DATA sections; look for the next one.
          const nd = t.indexOf("DATA;", i);
          if (nd < 0) break;
          i = nd + 5;
        } else {
          i++;
        }
      }
      if (onProgress) await onProgress(1);
      return this;
    }

    parseHeader() {
      const t = this.text;
      const hs = t.indexOf("HEADER;");
      const he = t.indexOf("ENDSEC;", hs);
      if (hs < 0 || he < 0) return;
      const head = t.slice(hs + 7, he);
      const grab = (name) => {
        const m = head.match(new RegExp(name + "\\s*\\(([\\s\\S]*?)\\);"));
        return m ? parseArgs(m[1]) : null;
      };
      const fd = grab("FILE_DESCRIPTION");
      const fn = grab("FILE_NAME");
      const fs = grab("FILE_SCHEMA");
      const flat = (v) => (Array.isArray(v) ? v.filter((x) => x).join(", ") : v || "");
      this.header = {
        description: fd ? flat(fd[0]) : "",
        implementationLevel: fd ? fd[1] : "",
        name: fn ? fn[0] : "",
        timeStamp: fn ? fn[1] : "",
        author: fn ? flat(fn[2]) : "",
        organization: fn ? flat(fn[3]) : "",
        preprocessor: fn ? fn[4] : "",
        originatingSystem: fn ? fn[5] : "",
        authorization: fn ? fn[6] : "",
        schema: fs ? flat(fs[0]) : "",
      };
      this.schema = (this.header.schema || "").toUpperCase();
      const mvd = /ViewDefinition\s*\[([^\]]*)\]/i.exec(this.header.description || "");
      this.header.mvd = mvd ? mvd[1].trim() : "";
    }

    has(id) { return this.indexOf.has(id); }
    type(id) {
      const p = this.indexOf.get(id);
      return p === undefined ? null : this.types[p];
    }
    raw(id) {
      const p = this.indexOf.get(id);
      if (p === undefined) return "";
      return this.text.slice(this.starts[p], this.ends[p]);
    }
    line(id) {
      const p = this.indexOf.get(id);
      if (p === undefined) return "";
      return `#${id}=${this.types[p]}(${this.raw(id)});`;
    }
    args(id) {
      let a = this.cache.get(id);
      if (a) return a;
      const p = this.indexOf.get(id);
      if (p === undefined) return [];
      try {
        a = parseArgs(this.text.slice(this.starts[p], this.ends[p]));
      } catch (e) {
        this.errors++;
        a = [];
      }
      this.cache.set(id, a);
      return a;
    }
    /** True when the instance is an IfcRoot subtype (first attribute is a 22-char GlobalId). */
    isRooted(id) {
      const p = this.indexOf.get(id);
      if (p === undefined) return false;
      const s = this.starts[p];
      const t = this.text;
      let k = s;
      while (t.charCodeAt(k) <= 32) k++;
      if (t.charCodeAt(k) !== 39 || t.charCodeAt(k + 23) !== 39) return false;
      for (let i = k + 1; i <= k + 22; i++) {
        const c = t.charCodeAt(i);
        const ok = (c >= 48 && c <= 57) || (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || c === 95 || c === 36;
        if (!ok) return false;
      }
      let j = k + 24;
      while (t.charCodeAt(j) <= 32) j++;
      return t.charCodeAt(j) === 44 /* , */;
    }
    ofType(type) { return this.byType.get(type) || []; }
    get size() { return this.ids.length; }
  }

  // Value helpers
  function refId(v) { return v instanceof Ref ? v.id : null; }
  function refIds(v) {
    if (v instanceof Ref) return [v.id];
    if (Array.isArray(v)) return v.filter((x) => x instanceof Ref).map((x) => x.id);
    return [];
  }
  function str(v) {
    if (v == null || v === DERIVED) return "";
    if (typeof v === "string") return v;
    if (v instanceof Typed) return str(v.value);
    if (v instanceof Enum) return v.v;
    return String(v);
  }
  /** Convert a STEP value to a plain JS value { value, datatype }. */
  function plain(v) {
    if (v == null || v === DERIVED) return { value: null, datatype: null };
    if (v instanceof Typed) {
      const inner = plain(v.value);
      return { value: inner.value, datatype: v.type };
    }
    if (v instanceof Enum) {
      if (v.v === "T") return { value: true, datatype: "BOOLEAN" };
      if (v.v === "F") return { value: false, datatype: "BOOLEAN" };
      if (v.v === "U") return { value: null, datatype: "LOGICAL" };
      return { value: v.v, datatype: "ENUM" };
    }
    if (v instanceof Ref) return { value: "#" + v.id, datatype: "REF" };
    if (Array.isArray(v)) return { value: v.map((x) => plain(x).value), datatype: "LIST" };
    return { value: v, datatype: typeof v === "number" ? "REAL" : "STRING" };
  }
  function display(v) {
    if (v == null || v === DERIVED) return "";
    if (v instanceof Ref) return "#" + v.id;
    if (v instanceof Enum) return "." + v.v + ".";
    if (v instanceof Typed) return `${KG.Schema ? KG.Schema.canonical(v.type) : v.type}(${display(v.value)})`;
    if (Array.isArray(v)) return "(" + v.map(display).join(", ") + ")";
    if (typeof v === "string") return v;
    return String(v);
  }

  KG.Step = { StepFile, Ref, Enum, Typed, DERIVED, parseArgs, decodeString, refId, refIds, str, plain, display };
})((window.KG = window.KG || {}));
