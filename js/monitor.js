/*
 * Monitor Laporan — the Laporan Tim app's side (installed in superapp-laporan2, v59+; needs SA.photo.shrink,
 * added to js/photo.js in v59c).
 *
 * For every Security / Patrol report the phone SAVES:
 *   1. a small NOTE (no photos, no caption) is queued and posted to the Apps Script -- the board's
 *      red/green needs only this;
 *   2. then each of the report's photos, as a smaller copy (1024 px, about 100-150 KB), one at a time.
 *
 * Nothing here is in the guard's way (Billy, 2026-11-23):
 *   - Simpan -> WhatsApp works exactly as before; nothing waits for the network.
 *   - Notes always go before photos.
 *   - Photos wait PHOTO_DELAY after the last save, so the WhatsApp send has the signal to itself, and
 *     only go while the app is open on screen (a web app cannot upload once it is in the background).
 *   - Queues live in localStorage and survive closing the app; the photos themselves are read from
 *     the app's own database when their turn comes, so nothing is stored twice.
 */
var SA = window.SA || {};

(function (SA) {
  'use strict';

  var URL_ = 'https://script.google.com/macros/s/AKfycbxCeMscffzpDzhQJwnhSHw6lX_VazOT5GL2iyui2rc25N26hCKGskeJcTW6GhniE9uusQ/exec';
  var KEY = 'laporan-tim-monitor-1';          // = WRITE_KEY in the script's properties
  var TEAMS = { security: true, patrol: true };
  /* '' for Laporan Tim itself. A copy published under the same github.io account shares the phone's
     storage with it, so it needs its own prefix (Sec Prabhu Online: 'uji'); its rows then start 'uji'. */
  var PREFIX = 'uji';                         // Sec Prabhu Online (same origin as sec-prabhu)
  function keyOf(name) { return PREFIX ? PREFIX + name.charAt(0).toUpperCase() + name.slice(1) : name; }
  var NOTES = keyOf('monitorQueue');
  var PHOTOS = keyOf('monitorPhotos');
  var DEVICE = keyOf('monitorDevice');
  var BATCH = 50;
  var KEEP_NOTES = 1000;                      // the oldest give way if a phone is offline for weeks
  var KEEP_PHOTOS = 3000;
  var PHOTO_SIZE = 1024;                      // longest side of the uploaded copy
  var PHOTO_CEILING = 180 * 1024;              // bytes: SA.photo.shrink lowers the quality until it fits
  var PHOTO_DELAY = 90 * 1000;                // after a save: WhatsApp first
  var MAX_TRIES = 8;                          // the SERVER refused it 8 times: dropped (the report is not).
                                              // A network failure never counts -- weak signal only waits.

  function read(key) { try { return JSON.parse(localStorage.getItem(key)) || []; } catch (e) { return []; } }
  function write(key, q, keep) {
    try { localStorage.setItem(key, JSON.stringify(q.slice(-keep))); } catch (e) { /* full: the note is lost, the report is not */ }
    changed();
  }
  function changed() {
    try { window.dispatchEvent(new Event('monitorchange')); } catch (e) { /* old browser */ }
  }
  function device() {
    var id = null;
    try { id = localStorage.getItem(DEVICE); } catch (e) { /* private mode */ }
    if (!id) {
      id = PREFIX + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
      try { localStorage.setItem(DEVICE, id); } catch (e) { /* ignore */ }
    }
    return id;
  }

  function noteOf(r) {
    var dev = device();
    return {
      id: dev + '-' + r.id + '-' + String(r.timestamp || '').replace(/\D/g, ''),
      team: r.team, kind: r.kind,
      unit: r.team === 'patrol' ? r.patrolId : r.post,
      zone: r.zone || '', shift: r.shift, shiftDate: r.shiftDate,
      hour: r.kind === 'check' ? r.hour : '',
      point: r.kind === 'patrol' ? (r.point || r.area || '') : '',
      pass: r.pass || '',
      timestamp: r.timestamp, reporter: r.reporter || '', officers: r.officers || [],
      build: SA.BUILD || '', device: dev,
      report: reportOf(r)
    };
  }

  /* The report as the app's Excel lays it out: that team's sheet for this kind, built by the app's own
     code for this one record. The server puts it in a tab of the same name, so the Drive sheet reads
     like the export. Kinds the Excel leaves out (Mulai Shift) send no row. Photo cells are filled by the
     server with a Drive link when each photo arrives. */
  var SHEETS = { security: function () { return SA.secRecords; }, patrol: function () { return SA.patrolRecords; } };
  function reportOf(r) {
    try {
      var mod = SHEETS[r.team] && SHEETS[r.team]();
      if (!mod || !mod.sheets) return null;
      var s = mod.sheets([r]).filter(function (x) { return x.rows && x.rows.length; })[0];
      if (!s) return null;
      var headers = [], cells = [], photoCols = [];
      s.columns.forEach(function (c, i) {
        if (c.type === 'index') return;                 // "No": the sheet numbers its own rows
        headers.push(c.header);
        var v = s.rows[0].cells[i];
        cells.push(v == null ? '' : typeof v === 'number' ? v : String(v));
        if (c.type === 'photo') photoCols.push(c.header);
      });
      return { sheet: s.name, headers: headers, cells: cells, photoCols: photoCols };
    } catch (e) { return null; }                      // never stop the note for the row
  }

  function post(body) {
    var text = JSON.stringify(body);
    return fetch(URL_, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },   // keeps it a simple request: no preflight
      body: text,
      keepalive: text.length < 60000                              // keepalive bodies are capped at 64 KB
    }).catch(function (e) {
      var err = new Error('jaringan: ' + (e && e.message || e));
      err.network = true;                                         // no answer at all: try again later, never count it
      throw err;
    }).then(function (res) { return res.json(); });
  }

  /* ---------------- notes ---------------- */

  var noteBusy = false;
  function flushNotes() {
    if (!URL_ || noteBusy || navigator.onLine === false) return Promise.resolve(false);
    var q = read(NOTES);
    if (!q.length) return Promise.resolve(true);
    noteBusy = true;
    var before = q.length;
    return post({ k: KEY, v: 2, notes: q.slice(0, BATCH) }).then(function (res) {
      if (!res || !res.ok) return false;
      var kept = {};
      // "saved" also lists notes the server refused as not ours: they leave the queue too, or they
      // would be sent for ever and hold every photo behind them
      (res.saved || []).forEach(function (id) { kept[id] = true; });
      write(NOTES, read(NOTES).filter(function (n) { return !kept[n.id]; }), KEEP_NOTES);
      return true;
    }).catch(function () { return false; }).then(function (ok) {
      noteBusy = false;
      var left = read(NOTES).length;
      if (ok && left && left < before) return flushNotes();      // more to send, and progress was made
      return ok && !left;
    });
  }

  /* ---------------- photos ---------------- */

  var lastSave = 0;
  var photoBusy = false;

  function base64(blob) {
    return new Promise(function (ok, bad) {
      var r = new FileReader();
      r.onload = function () { ok(String(r.result).split(',')[1] || ''); };
      r.onerror = bad;
      r.readAsDataURL(blob);
    });
  }

  function dropJob(job) {
    write(PHOTOS, read(PHOTOS).filter(function (j) { return !(j.noteId === job.noteId && j.i === job.i); }), KEEP_PHOTOS);
  }
  function bumpJob(job) {
    var q = read(PHOTOS);
    q.forEach(function (j) { if (j.noteId === job.noteId && j.i === job.i) j.tries = (j.tries || 0) + 1; });
    write(PHOTOS, q.filter(function (j) { return (j.tries || 0) < MAX_TRIES; }), KEEP_PHOTOS);
  }

  function flushPhotos() {
    if (!URL_ || photoBusy || navigator.onLine === false || document.hidden) return;
    if (Date.now() - lastSave < PHOTO_DELAY) return;
    if (read(NOTES).length) return;                     // notes first, always
    var job = read(PHOTOS)[0];
    if (!job || !SA.db || !SA.db.bySession || !SA.photo || !SA.photo.shrink) return;
    photoBusy = true;
    SA.db.bySession(job.sessionId).then(function (records) {
      var rec = records.filter(function (r) { return r.id === job.recordId; })[0];
      var photo = rec && rec.photos && rec.photos[job.i];
      if (!photo || !photo.blob) { dropJob(job); return 'next'; }   // report deleted on the phone: nothing to send
      return SA.photo.shrink(photo.blob, PHOTO_SIZE, PHOTO_CEILING).then(base64).then(function (data) {
        return post({ k: KEY, v: 1, photo: Object.assign({}, job, { data: data, takenAt: photo.takenAt || '' }) });
      }).then(function (res) {
        if (res && res.ok) { dropJob(job); return 'next'; }
        bumpJob(job);
        return 'stop';
      });
    }).catch(function (err) {
      if (!(err && err.network)) bumpJob(job);           // a photo that cannot be read or sent counts; no signal does not
      return 'stop';
    }).then(function (go) {
      photoBusy = false;
      if (go === 'next') setTimeout(flushPhotos, 1500);
    });
  }

  function flush() {
    flushNotes().then(function () { flushPhotos(); });
  }

  SA.monitor = {
    /** Call with the record SA.db.add resolved with (it carries its id). Never throws. */
    note: function (record) {
      try {
        if (!URL_ || !record || !TEAMS[record.team] || record.id == null) return;
        var n = noteOf(record);
        var q = read(NOTES);
        q.push(n);
        write(NOTES, q, KEEP_NOTES);
        var photos = record.photos || [];
        if (photos.length) {
          var jobs = read(PHOTOS);
          var rep = n.report || {};
          photos.forEach(function (p, i) {
            jobs.push({ noteId: n.id, sessionId: record.sessionId, recordId: record.id, i: i, n: photos.length,
              team: n.team, kind: n.kind, unit: n.unit, shift: n.shift, shiftDate: n.shiftDate,
              hour: n.hour, point: n.point, pass: n.pass,
              sheet: rep.sheet || '', col: (rep.photoCols || [])[i] || '', code: p.sealCode || '' });
          });
          write(PHOTOS, jobs, KEEP_PHOTOS);
        }
        lastSave = Date.now();
        flushNotes();
        setTimeout(flushPhotos, PHOTO_DELAY + 1000);
      } catch (e) { /* the monitor must never stop a report */ }
    },
    pending: function () { return read(NOTES).length; },
    waitingPhotos: function () { return read(PHOTOS).length; },
    flush: flush
  };

  window.addEventListener('online', flush);
  document.addEventListener('visibilitychange', function () { if (!document.hidden) flush(); });
  setInterval(flush, 60000);
  setTimeout(flush, 4000);
})(SA);
