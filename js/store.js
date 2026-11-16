/* Lưu trữ cục bộ bằng IndexedDB: chương đã chỉnh sửa và ảnh tải lên. Có dự phòng trong bộ nhớ. */
(function (root) {
  'use strict';
  var DB = 'pmnm-editor', VER = 1, dbp = null, mem = { chapters: {}, images: {} }, useMem = false;

  function open() {
    if (dbp) return dbp;
    dbp = new Promise(function (res, rej) {
      try {
        var r = indexedDB.open(DB, VER);
        r.onupgradeneeded = function () { var d = r.result; d.createObjectStore('chapters'); d.createObjectStore('images'); };
        r.onsuccess = function () { res(r.result); };
        r.onerror = function () { rej(r.error); };
      } catch (e) { rej(e); }
    }).catch(function () { useMem = true; return null; });
    return dbp;
  }

  function run(store, mode, fn) {
    return open().then(function (db) {
      if (!db) return fn(null, true);
      return new Promise(function (res, rej) {
        var t = db.transaction(store, mode), s = t.objectStore(store), r = fn(s, false);
        t.oncomplete = function () { res(r && 'result' in r ? r.result : undefined); };
        t.onerror = function () { rej(t.error); };
        t.onabort = function () { rej(t.error); };
      });
    });
  }

  root.Store = {
    get: function (store, key) {
      return run(store, 'readonly', function (s, m) { return m ? { result: mem[store][key] } : s.get(key); });
    },
    put: function (store, key, val) {
      return run(store, 'readwrite', function (s, m) { if (m) { mem[store][key] = val; return null; } return s.put(val, key); });
    },
    del: function (store, key) {
      return run(store, 'readwrite', function (s, m) { if (m) { delete mem[store][key]; return null; } return s.delete(key); });
    },
    keys: function (store) {
      return run(store, 'readonly', function (s, m) { return m ? { result: Object.keys(mem[store]) } : s.getAllKeys(); });
    },
    clear: function (store) {
      return run(store, 'readwrite', function (s, m) { if (m) { mem[store] = {}; return null; } return s.clear(); });
    }
  };
})(window);
