/*
 * Ficha HR en Excel
 * Rellena la plantilla plantilla-hr.xlsx con los datos del evento y la descarga.
 * Edita directamente el XML del libro, así la plantilla conserva intactos
 * formato, logotipo, celdas combinadas y área de impresión.
 */
(function (root) {
  "use strict";

  var TEMPLATE_URL = "./plantilla-hr.xlsx";
  var SHEET = "xl/worksheets/sheet1.xml";
  var SHEET_RELS = "xl/worksheets/_rels/sheet1.xml.rels";
  var DAYS = ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"];
  var MONTHS_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

  // Celdas de la plantilla (columna de valor de cada fila)
  var CELLS = {
    fecha: "C11", llegada: "J11", evento: "C12", recinto: "D13",
    poblacion: "C16", provincia: "G16", cliente: "C17",
    personal: ["D26", "D27", "D28", "D29"],
    vehiculoFilas: [34, 35, 36, 37, 38, 39, 40, 41],
    limpiar: ["G22"] // valor de prueba que venía en la plantilla
  };
  // Hipervínculos de la plantilla (rId de sheet1.xml.rels)
  var LINKS = { recinto: "rId1", acceso: "rId2", parking: "rId3", hotel: "rId4", meteo: "rId5" };

  function str(v) { return v === null || v === undefined ? "" : String(v).trim(); }
  function norm(v) { return str(v).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, ""); }
  function xml(v) {
    return String(v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  }

  function longDate(iso) {
    var p = iso.split("-").map(Number), d = new Date(p[0], p[1] - 1, p[2]);
    return DAYS[d.getDay()] + ", " + p[2] + " de " + MONTHS_ES[p[1] - 1] + " de " + p[0];
  }

  function mapsUrl(q) { return "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(q); }
  function meteoUrl(poblacion, recinto) {
    var slug = norm(poblacion).replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
    if (slug) return "https://www.eltiempo.es/" + slug + ".html";
    return "https://www.google.com/search?q=" + encodeURIComponent("tiempo " + (recinto || ""));
  }

  function extractPlate(name) {
    var m = str(name).toUpperCase().match(/\b(\d{4})\s*-?\s*([B-DF-HJ-NP-TV-Z]{3})\b/);
    return m ? m[1] + " " + m[2] : "";
  }

  // Escribe un texto en una celda existente conservando su estilo
  function setCell(sheetXml, ref, value) {
    var re = new RegExp('<c r="' + ref + '"([^>]*?)(/>|>[\\s\\S]*?</c>)');
    var m = sheetXml.match(re);
    if (!m) throw new Error("La plantilla no tiene la celda " + ref + ".");
    var style = (m[1].match(/\ss="\d+"/) || [""])[0];
    var cell = str(value) === ""
      ? '<c r="' + ref + '"' + style + "/>"
      : '<c r="' + ref + '"' + style + ' t="inlineStr"><is><t xml:space="preserve">' + xml(value) + "</t></is></c>";
    return sheetXml.replace(re, cell);
  }

  function setLink(relsXml, id, url) {
    var re = new RegExp('(<Relationship Id="' + id + '"[^>]*?Target=")[^"]*(")');
    return relsXml.replace(re, "$1" + xml(url) + "$2");
  }

  // data: {date, event, client, space, poblacion, provincia, llegada, staff:[name], vehicles:[{name, plate, driver}]}
  function fillWorkbook(zip, data) {
    return Promise.all([zip.file(SHEET).async("string"), zip.file(SHEET_RELS).async("string")]).then(function (parts) {
      var s = parts[0], rels = parts[1];
      var lugar = [data.space, data.poblacion, data.provincia].filter(str).join(", ");

      s = setCell(s, CELLS.fecha, data.date ? longDate(data.date) : "");
      s = setCell(s, CELLS.llegada, data.llegada);
      s = setCell(s, CELLS.evento, data.event);
      s = setCell(s, CELLS.recinto, data.space);
      s = setCell(s, CELLS.poblacion, data.poblacion);
      s = setCell(s, CELLS.provincia, data.provincia);
      s = setCell(s, CELLS.cliente, data.client);
      CELLS.limpiar.forEach(function (ref) { s = setCell(s, ref, ""); });

      var staff = (data.staff || []).filter(str);
      CELLS.personal.forEach(function (ref, i) {
        var isLast = i === CELLS.personal.length - 1;
        s = setCell(s, ref, isLast ? staff.slice(i).join(", ") : staff[i] || "");
      });

      var vehicles = data.vehicles || [];
      CELLS.vehiculoFilas.forEach(function (row, i) {
        var v = vehicles[i] || {};
        s = setCell(s, "B" + row, v.driver);
        s = setCell(s, "E" + row, v.plate);
        s = setCell(s, "H" + row, v.name);
      });

      if (lugar) {
        rels = setLink(rels, LINKS.recinto, mapsUrl(lugar));
        rels = setLink(rels, LINKS.acceso, mapsUrl(lugar));
        rels = setLink(rels, LINKS.parking, mapsUrl(lugar));
        rels = setLink(rels, LINKS.hotel, mapsUrl("hoteles cerca de " + lugar));
      }
      rels = setLink(rels, LINKS.meteo, meteoUrl(data.poblacion, lugar));

      zip.file(SHEET, s);
      zip.file(SHEET_RELS, rels);
      return zip;
    });
  }

  function fileName(data) {
    var slug = norm(data.event || data.client || "evento").replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "").slice(0, 40);
    return "HR_" + data.date + "_" + (slug || "evento") + ".xlsx";
  }

  /* ---------- Integración con el portal ---------- */

  function eventTokens(ev) {
    var stop = { "evento": 1, "concierto": 1, "festival": 1, "plaza": 1, "teatro": 1, "auditorio": 1, "pabellon": 1, "recinto": 1 };
    return norm([ev.event, ev.client, ev.space].join(" ")).split(/[^a-z0-9]+/).filter(function (t) { return t.length >= 4 && !stop[t]; });
  }

  function dayAssignments(records, date) {
    var out = [];
    (records || []).forEach(function (r) {
      var a = (r.assignments || []).find(function (x) { return x.date === date; });
      if (a && (a.type === "Trabajo" || !a.type)) out.push({ name: r.name, value: a.value });
    });
    return out;
  }

  function preselect(list, tokens) {
    var any = false;
    list.forEach(function (x) {
      var v = norm(x.value);
      x.checked = tokens.some(function (t) { return v.indexOf(t) >= 0; });
      any = any || x.checked;
    });
    if (!any) list.forEach(function (x) { x.checked = true; });
    return list;
  }

  function injectStyles() {
    if (document.getElementById("hrStyles")) return;
    var css = ".hr-backdrop{position:fixed;inset:0;z-index:50;display:grid;place-items:center;padding:16px;background:rgba(16,45,40,.45)}" +
      ".hr-modal{width:min(560px,100%);max-height:calc(100vh - 32px);overflow:auto;border-radius:10px;background:#fff;padding:22px;box-shadow:0 20px 50px rgba(0,0,0,.25)}" +
      ".hr-modal h3{margin:0 0 4px;font:700 19px var(--display)}.hr-modal h4{margin:18px 0 8px;font-size:12px;text-transform:uppercase;color:var(--muted)}" +
      ".hr-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}.hr-grid label{display:grid;gap:3px;font-size:12px;color:var(--muted)}.hr-grid .field{min-width:0;width:100%}" +
      ".hr-check{display:flex;gap:9px;align-items:flex-start;padding:7px 0;border-bottom:1px solid var(--line)}.hr-check small{display:block;color:var(--muted)}" +
      ".hr-actions{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}.hr-note{font-size:12px;color:var(--muted);margin-top:6px}";
    var el = document.createElement("style"); el.id = "hrStyles"; el.textContent = css; document.head.appendChild(el);
  }

  function checklist(list, prefix, extra) {
    if (!list.length) return '<div class="hr-note">Nada asignado este día.</div>';
    return list.map(function (x, i) {
      return '<label class="hr-check"><input type="checkbox" data-' + prefix + '="' + i + '"' + (x.checked ? " checked" : "") + '><span>' +
        esc(x.name) + "<small>" + esc(x.value) + (extra ? extra(x) : "") + "</small></span></label>";
    }).join("");
  }
  function esc(v) { return root.escapeHtml ? root.escapeHtml(str(v)) : xml(str(v)); }

  function open(date) {
    var model = root.model || {};
    var ev = (model.events || []).find(function (x) { return x.date === date; });
    if (!ev) return;
    if (!root.JSZip) { alert("No se pudo cargar el generador de Excel. Revisa la conexión y recarga la página."); return; }
    injectStyles();

    var tokens = eventTokens(ev);
    var staff = preselect(dayAssignments(model.workers, date), tokens);
    var workerNames = (model.workers || []).map(function (w) { return w.name; });
    var vehicles = preselect(dayAssignments(model.vehicles, date), tokens).map(function (v) {
      var driver = workerNames.find(function (n) { return norm(n) && norm(v.value).indexOf(norm(n)) >= 0; }) || "";
      v.plate = extractPlate(v.name); v.driver = driver; return v;
    });

    var wrap = document.createElement("div");
    wrap.className = "hr-backdrop";
    wrap.innerHTML = '<div class="hr-modal" role="dialog" aria-modal="true" aria-label="Generar ficha HR">' +
      "<h3>Ficha HR · " + esc(ev.event || "Evento") + '</h3><div class="muted">' + esc(longDate(date)) + (ev.client ? " · " + esc(ev.client) : "") + (ev.space ? " · " + esc(ev.space) : "") + "</div>" +
      '<h4>Datos del evento</h4><div class="hr-grid">' +
      '<label>Población<input class="field" id="hrPoblacion"></label><label>Provincia<input class="field" id="hrProvincia"></label>' +
      '<label style="grid-column:1/-1">Llegada camiones / comienzo montaje<input class="field" id="hrLlegada" placeholder="Ej. 08:00"></label></div>' +
      "<h4>Personal TUA (el primero marcado será el responsable)</h4>" + checklist(staff, "staff") +
      "<h4>Vehículos</h4>" + checklist(vehicles, "veh", function (v) { return v.driver ? " · conductor: " + esc(v.driver) : ""; }) +
      '<div class="hr-note">Se rellenan contactos del personal, vehículos, enlaces de mapas y meteorología. El resto se completa en el Excel.</div>' +
      '<div class="hr-actions"><button class="button" data-hr="cancel">Cancelar</button><button class="button primary" data-hr="go">Descargar Excel</button></div></div>';
    document.body.appendChild(wrap);

    function close() { if (wrap.parentNode) wrap.parentNode.removeChild(wrap); }
    wrap.addEventListener("click", function (e) { if (e.target === wrap) close(); });
    wrap.querySelector('[data-hr="cancel"]').addEventListener("click", close);
    wrap.querySelector('[data-hr="go"]').addEventListener("click", function () {
      var btn = this;
      var picked = function (attr, list) {
        return Array.prototype.filter.call(wrap.querySelectorAll("[data-" + attr + "]"), function (c) { return c.checked; })
          .map(function (c) { return list[Number(c.getAttribute("data-" + attr))]; });
      };
      var data = {
        date: date, event: ev.event, client: ev.client, space: ev.space,
        poblacion: wrap.querySelector("#hrPoblacion").value, provincia: wrap.querySelector("#hrProvincia").value,
        llegada: wrap.querySelector("#hrLlegada").value,
        staff: picked("staff", staff).map(function (x) { return x.name; }),
        vehicles: picked("veh", vehicles)
      };
      btn.disabled = true; btn.textContent = "Generando…";
      generate(data).then(close).catch(function (err) {
        alert("No se pudo generar la ficha: " + err.message);
        btn.disabled = false; btn.textContent = "Descargar Excel";
      });
    });
  }

  function generate(data) {
    return fetch(TEMPLATE_URL + "?_=" + Date.now(), { cache: "no-store" }).then(function (r) {
      if (!r.ok) throw new Error("No se encontró la plantilla (HTTP " + r.status + ").");
      return r.arrayBuffer();
    }).then(function (buf) { return root.JSZip.loadAsync(buf); })
      .then(function (zip) { return fillWorkbook(zip, data); })
      .then(function (zip) { return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", compression: "DEFLATE" }); })
      .then(function (blob) {
        var a = document.createElement("a");
        a.href = URL.createObjectURL(blob); a.download = fileName(data);
        document.body.appendChild(a); a.click();
        setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
      });
  }

  var api = { open: open, fillWorkbook: fillWorkbook, fileName: fileName, extractPlate: extractPlate, longDate: longDate };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.TUAFichaHR = api;
})(typeof window !== "undefined" ? window : globalThis);
