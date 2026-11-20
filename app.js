/* Web Box: shows any https page inside an Excel content add-in.
   The link and zoom are saved in the workbook (Office document settings),
   so each box reopens on its own page on Windows, Mac and Excel for the web. */
(function () {
  "use strict";

  var KEY_URL = "webBoxUrl";
  var KEY_ZOOM = "webBoxZoom";
  var ZOOMS = [0.5, 0.6, 0.7, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2];

  var $ = function (id) { return document.getElementById(id); };
  var stage = $("stage"), frame = $("frame"), setup = $("setup"), bar = $("bar");
  var input = $("urlInput"), err = $("urlError"), tableauFix = $("tableauFix");

  var state = { url: "", zoom: 1, inOffice: false };

  /* ---------- storage: workbook settings in Office, localStorage in a plain browser ---------- */
  function load(key) {
    if (state.inOffice) { return Office.context.document.settings.get(key); }
    try { return JSON.parse(localStorage.getItem(key)); } catch (e) { return null; }
  }
  function save(values) {
    if (state.inOffice) {
      var s = Office.context.document.settings;
      Object.keys(values).forEach(function (k) { s.set(k, values[k]); });
      s.saveAsync(function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) { console.warn("Web Box: settings not saved", r.error); }
      });
      return;
    }
    try { Object.keys(values).forEach(function (k) { localStorage.setItem(k, JSON.stringify(values[k])); }); } catch (e) {}
  }

  /* ---------- link clean-up ---------- */
  function extractUrl(text) {
    text = (text || "").trim();
    // Pasted embed code: <iframe src="..."> or <tableau-viz src="..."> or <param name='name' ...> (old Tableau embed)
    var m = text.match(/\bsrc\s*=\s*["']([^"']+)["']/i);
    if (m) { return m[1]; }
    var host = text.match(/<param\s+name=['"]host_url['"]\s+value=['"]([^'"]+)['"]/i);
    var name = text.match(/<param\s+name=['"]name['"]\s+value=['"]([^'"]+)['"]/i);
    if (host && name) {
      return decodeURIComponent(host[1]).replace(/\/+$/, "") + "/views/" + name[1].replace(/&#47;/g, "/");
    }
    return text;
  }

  function normalise(text, fixTableau) {
    var raw = extractUrl(text);
    if (!raw) { throw new Error("Paste a link first."); }
    if (!/^[a-z][a-z0-9+.-]*:/i.test(raw)) { raw = "https://" + raw; }
    var u;
    try { u = new URL(raw); } catch (e) { throw new Error("That doesn't look like a web link."); }
    if (u.protocol === "http:") { throw new Error("Use an https:// link. Excel blocks http:// pages inside add-ins."); }
    if (u.protocol !== "https:") { throw new Error("Only https:// links can be shown."); }

    // Tableau's portal (address-bar) links keep the view in the hash: /#/views/WB/Sheet?... or
    // /#/site/SITE/views/WB/Sheet?... Tableau refuses to show the portal inside another page
    // (X-Frame-Options: SAMEORIGIN), so turn them into the frameable view path /views/... or /t/SITE/views/...
    var search = u.search, path = u.pathname, hash = u.hash;
    var hm = fixTableau && hash.match(/^#(\/(?:site\/[^/?]+\/)?views\/[^?]+)(\?.*)?$/);
    if (hm) {
      path = hm[1].replace(/^\/site\/([^/]+)\//, "/t/$1/");
      search = hm[2] || "";
      hash = "";
    } else if (fixTableau && /^#\//.test(hash) && !/\/views\//.test(path)) {
      throw new Error("That's a Tableau page (home, a project or a workbook list), which Tableau won't show inside " +
                      "another page. Open the dashboard itself and copy its link or Share link.");
    }

    if (fixTableau && /\/views\//.test(path)) {
      // Rebuild the query by hand so Tableau's ":" parameters stay literal (URLSearchParams would write %3A).
      var drop = { "publish": 1, ":origin": 1, ":iid": 1, ":embed": 1, ":showVizHome": 1 };
      var keep = search.replace(/^\?/, "").split("&").filter(function (part) {
        if (!part) { return false; }
        var k = decodeURIComponent(part.split("=")[0]);
        return !drop[k];
      });
      keep.unshift(":embed=y", ":showVizHome=no");
      return u.origin + path + "?" + keep.join("&") + hash;
    }
    return u.toString();
  }

  /* ---------- views ---------- */
  function showSetup(prefill) {
    stage.hidden = true; bar.hidden = true; setup.hidden = false;
    input.value = prefill || "";
    err.textContent = "";
    $("cancelSetup").hidden = !state.url;
    setTimeout(function () { input.focus(); input.select(); }, 0);
  }

  function showPage(url) {
    state.url = url;
    setup.hidden = true; stage.hidden = false; bar.hidden = false;
    if (frame.getAttribute("src") !== url) { frame.setAttribute("src", url); }
    applyZoom();
  }

  function applyZoom() {
    var z = state.zoom;
    frame.style.transform = z === 1 ? "" : "scale(" + z + ")";
    frame.style.width = (100 / z) + "%";
    frame.style.height = (100 / z) + "%";
    $("btnZoomPct").textContent = Math.round(z * 100) + "%";
  }

  function stepZoom(dir) {
    var i = ZOOMS.indexOf(state.zoom);
    if (i < 0) { i = ZOOMS.indexOf(1); }
    i = Math.max(0, Math.min(ZOOMS.length - 1, i + dir));
    state.zoom = ZOOMS[i];
    applyZoom();
    save(obj(KEY_ZOOM, state.zoom));
  }

  function obj(k, v) { var o = {}; o[k] = v; return o; }

  function openInBrowser() {
    var url = state.url;
    try {
      if (state.inOffice && Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1")) {
        Office.context.ui.openBrowserWindow(url);
        return;
      }
    } catch (e) {}
    window.open(url, "_blank", "noopener");
  }

  /* ---------- sign in ----------
     Tableau can't show a company sign-in inside a framed page, so its own "Sign in" button opens a
     pop-up. Inside Excel that pop-up can land in another browser, so the box never gets the sign-in
     and the pop-up hangs on "You may close this window now". Instead, sign in through an Office
     dialog: a window Excel opens itself, which keeps the sign-in where the box can use it. */
  var noteTimer = null;
  function note(text, ms) {
    var el = $("note");
    clearTimeout(noteTimer);
    el.textContent = text || "";
    el.hidden = !text;
    if (text && ms) { noteTimer = setTimeout(function () { el.hidden = true; }, ms); }
  }

  function reloadFrame() {
    var u = state.url;
    frame.setAttribute("src", "about:blank");
    setTimeout(function () { frame.setAttribute("src", u); }, 30);
  }

  // The page as a normal (not embedded) page, so Tableau runs its usual sign-in, not the pop-up one.
  function signInTarget(url) {
    try {
      var u = new URL(url);
      var q = u.search.replace(/^\?/, "").split("&").filter(function (part) {
        if (!part) { return false; }
        var k = decodeURIComponent(part.split("=")[0]);
        return k !== ":embed" && k !== ":showVizHome";
      });
      return u.origin + u.pathname + (q.length ? "?" + q.join("&") : "") + u.hash;
    } catch (e) { return url; }
  }

  function signIn() {
    if (!state.url) { return; }
    var target = signInTarget(state.url);
    var dialogOk = false;
    try { dialogOk = state.inOffice && Office.context.requirements.isSetSupported("DialogApi", "1.1"); } catch (e) {}
    if (!dialogOk) {
      window.open(target, "_blank");
      note("Sign in in the new window, then press reload (↻) here.", 8000);
      return;
    }
    // Office only opens a dialog on the add-in's own site; signin.html then sends it on to the target.
    var page = new URL("signin.html", location.href);
    page.searchParams.set("to", target);
    Office.context.ui.displayDialogAsync(page.toString(), { height: 75, width: 60, promptBeforeOpen: false }, function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        var code = r.error && r.error.code;
        note(code === 12007 ? "A sign-in window is already open. Finish there and close it."
                            : "Couldn't open the sign-in window (" + (r.error && r.error.message || code) + ").", 10000);
        return;
      }
      note("Sign in in the window that opened. When the dashboard shows there, close that window.");
      // Fires when the user closes the window (12006) or it can't load a page; either way, try the box again.
      r.value.addEventHandler(Office.EventType.DialogEventReceived, function () {
        note("");
        reloadFrame();
      });
    });
  }

  /* ---------- wiring ---------- */
  function wire() {
    setup.addEventListener("submit", function (e) {
      e.preventDefault();
      var url;
      try { url = normalise(input.value, tableauFix.checked); }
      catch (ex) { err.textContent = ex.message; return; }
      save(obj(KEY_URL, url));
      frame.removeAttribute("src");
      showPage(url);
    });
    $("cancelSetup").addEventListener("click", function () { showPage(state.url); });

    $("pill").addEventListener("click", function () { bar.classList.toggle("collapsed"); });
    $("btnEdit").addEventListener("click", function () { showSetup(state.url); });
    // The parent's history includes the frame's navigations, so this steps the embedded page back/forward.
    $("btnBack").addEventListener("click", function () { history.back(); });
    $("btnFwd").addEventListener("click", function () { history.forward(); });
    $("btnReload").addEventListener("click", reloadFrame);
    $("btnSignIn").addEventListener("click", signIn);
    $("btnHome").addEventListener("click", function () { frame.setAttribute("src", "about:blank"); setTimeout(function () { frame.setAttribute("src", state.url); }, 30); });
    $("btnZoomOut").addEventListener("click", function () { stepZoom(-1); });
    $("btnZoomIn").addEventListener("click", function () { stepZoom(1); });
    $("btnZoomPct").addEventListener("click", function () { state.zoom = 1; applyZoom(); save(obj(KEY_ZOOM, 1)); });
    $("btnOpen").addEventListener("click", openInBrowser);
  }

  function start() {
    wire();
    var z = Number(load(KEY_ZOOM));
    state.zoom = ZOOMS.indexOf(z) >= 0 ? z : 1;
    var saved = load(KEY_URL);
    var fromQuery = new URLSearchParams(location.search).get("url"); // handy for testing outside Excel
    if (fromQuery && !saved) { try { saved = normalise(fromQuery, true); } catch (e) {} }
    // Repair links saved before address-bar (/#/views/...) links were converted: Tableau refuses to frame those.
    if (saved && /#\/(site\/[^/]+\/)?views\//.test(saved)) {
      try { var fixed = normalise(saved, true); if (fixed !== saved) { saved = fixed; save(obj(KEY_URL, saved)); } } catch (e) {}
    }
    if (saved) { showPage(saved); } else { showSetup(""); }
  }

  var started = false;
  function startOnce(inOffice) { if (started) { return; } started = true; state.inOffice = inOffice; start(); }

  if (window.Office && Office.onReady) {
    // Resolves inside Excel with info.host set, and in a plain browser tab with host null.
    Office.onReady(function (info) { startOnce(!!(info && info.host)); });
  } else {
    // office.js could not load (offline / blocked): still show the page.
    startOnce(false);
  }
})();
