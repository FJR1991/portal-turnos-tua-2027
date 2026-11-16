/*
 * Ficha HR en Excel
 * Rellena la plantilla plantilla-hr.xlsx con los datos del evento y la descarga.
 * Edita directamente el XML del libro, así la plantilla conserva intactos
 * formato, logotipo, celdas combinadas y área de impresión.
 */
(function (root) {
  "use strict";

  var TEMPLATE_URL = "./plantilla-hr.xlsx";
  // SHA-256 de la contraseña para generar fichas (no se guarda en claro)
  var PASSWORD_SHA256 = "ce512dbba6d0ce6c832884ba56d555db8f67670d667b56549037deadf0499677";
  var SHEET = "xl/worksheets/sheet1.xml";
  var SHEET_RELS = "xl/worksheets/_rels/sheet1.xml.rels";
  var WORKBOOK = "xl/workbook.xml";
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


  // Inserta "count" copias de la fila "after" justo debajo y desplaza todo lo que hay por debajo
  // (celdas, celdas combinadas, enlaces, saltos de página y área de impresión).
  function shiftRow(n, after, count) { return n > after ? n + count : n; }
  function shiftRefs(text, after, count) {
    return text.replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g, function (m, col, row) { return col + shiftRow(Number(row), after, count); });
  }
  function insertRows(sheetXml, workbookXml, after, count, label) {
    if (count <= 0) return [sheetXml, workbookXml];
    var template = "";
    sheetXml = sheetXml.replace(/<row r="(\d+)"([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g, function (m, r, attrs, tail, inner) {
      r = Number(r);
      if (r === after) template = m;
      if (r <= after) return m;
      var body = inner === undefined ? "/>" : ">" + inner.replace(/ r="([A-Z]{1,3})(\d+)"/g, function (x, c, n) { return ' r="' + c + (Number(n) + count) + '"'; }) + "</row>";
      return '<row r="' + (r + count) + '"' + attrs + body;
    });
    var copies = "";
    for (var i = 1; i <= count; i++) {
      var n = after + i;
      var row = template.replace(/<row r="\d+"/, '<row r="' + n + '"').replace(/ r="([A-Z]{1,3})\d+"/g, ' r="$1' + n + '"');
      if (label) {
        var lab = label(i);
        row = row.replace(new RegExp('<c r="' + label.col + n + '"([^>]*?)(/>|>[\\s\\S]*?</c>)'), function (x, a) {
          var st = (a.match(/\ss="\d+"/) || [""])[0];
          return '<c r="' + label.col + n + '"' + st + ' t="inlineStr"><is><t xml:space="preserve">' + xml(lab) + "</t></is></c>";
        });
      }
      copies += row;
    }
    sheetXml = sheetXml.replace(template, template + copies);
    // celdas combinadas: desplazar y copiar las de la fila modelo
    var newMerges = "";
    sheetXml = sheetXml.replace(/<mergeCell ref="([^"]+)"\s*\/>/g, function (m, ref) {
      var rows = ref.match(/\d+/g).map(Number);
      if (rows.every(function (x) { return x === after; })) {
        for (var i = 1; i <= count; i++) newMerges += '<mergeCell ref="' + ref.replace(/\d+/g, String(after + i)) + '"/>';
      }
      return '<mergeCell ref="' + shiftRefs(ref, after, count) + '"/>';
    });
    sheetXml = sheetXml.replace("</mergeCells>", newMerges + "</mergeCells>")
      .replace(/<mergeCells count="\d+"/, function () { return '<mergeCells count="' + (sheetXml.match(/<mergeCell ref=/g) || []).length + '"'; });
    sheetXml = sheetXml.replace(/<mergeCells count="\d+"/, '<mergeCells count="' + (sheetXml.match(/<mergeCell ref=/g) || []).length + '"');
    sheetXml = sheetXml.replace(/<hyperlink ref="([^"]+)"/g, function (m, ref) { return '<hyperlink ref="' + shiftRefs(ref, after, count) + '"'; })
      .replace(/<dimension ref="([^"]+)"/, function (m, ref) { return '<dimension ref="' + shiftRefs(ref, after, count) + '"'; })
      .replace(/<brk id="(\d+)"/g, function (m, id) { return '<brk id="' + shiftRow(Number(id), after, count) + '"'; });
    workbookXml = workbookXml.replace(/(<definedName name="_xlnm\.Print_Area"[^>]*>)([^<]+)(<\/definedName>)/, function (m, a, ref, b) { return a + shiftRefs(ref, after, count) + b; });
    return [sheetXml, workbookXml];
  }

  // data: {date, event, client, space, poblacion, provincia, llegada, staff:[name], vehicles:[{name, plate, driver}]}
  function fillWorkbook(zip, data) {
    return Promise.all([zip.file(SHEET).async("string"), zip.file(SHEET_RELS).async("string"), zip.file(WORKBOOK).async("string")]).then(function (parts) {
      var s = parts[0], rels = parts[1], wbx = parts[2];
      var V = data.venue || null, vv = function (k) { return V ? str(V[k]) : ""; };
      if (V) {
        s = setCell(s, CELLS.recinto, [data.space || vv("NOMBRE RECINTO"), vv("DIRECCIÓN")].filter(str).join(" · "));
        s = setCell(s, "D14", vv("DIRECCIÓN ACCESO CARGA"));
        s = setCell(s, "D15", vv("PARKING"));
        [["D44", "TIPO DE ESPACIO"], ["J44", "AFORO"], ["D45", "TIPO ESCENARIO"], ["J45", "MEDIDAS ESCENARIO"], ["D46", "ESCALERAS"], ["J46", "RAMPAS"],
         ["D47", "ACCESO CARGA"], ["J47", "PASACABLES"], ["D48", "RIGGING PA"], ["J48", "RIGGING LUCES"], ["D49", "RIGGING VIDEO"], ["J49", "CAÑÓN"],
         ["D50", "UBICACIÓN CONTROL"], ["J50", "MEDIDAS CONTROL"], ["D51", "OTROS DATOS ESPACIO"],
         ["D22", "PRODUCCIÓN LOCAL"], ["G22", "TELÉFONO PROD. LOCAL"], ["J22", "EMAIL PROD. LOCAL"],
         ["D30", "ELÉCTRICO CONTACTO"], ["G30", "TELÉFONO ELÉCTRICO"], ["J30", "EMAIL ELÉCTRICO"]].forEach(function (p) {
          if (vv(p[1])) s = setCell(s, p[0], vv(p[1]));
        });
        ["SONIDO", "ILUMINACIÓN", "VIDEO", "MOTORES", "CONTROL", "CAÑONES", "PARKING BUS"].forEach(function (t, i) {
          if (vv("TOMA " + t)) s = setCell(s, "D" + (55 + i), vv("TOMA " + t));
          if (vv("POSICIÓN TOMA " + t)) s = setCell(s, "I" + (55 + i), vv("POSICIÓN TOMA " + t));
        });
      }
      var staff = (data.staff || []).filter(str);
      var vehicles = data.vehicles || [];
      // una fila por persona y por vehículo: si no caben, se añaden filas a la plantilla
      var nP = CELLS.personal.length, extraP = Math.max(0, staff.length - nP);
      var lastP = Number(CELLS.personal[nP - 1].replace(/\D/g, ""));
      var labelP = function (i) { return "Personal TUA " + (nP + i); }; labelP.col = "B";
      var r1 = insertRows(s, wbx, lastP, extraP, labelP); s = r1[0]; wbx = r1[1];
      var vRows = CELLS.vehiculoFilas.map(function (r) { return r + extraP; });
      var extraV = Math.max(0, vehicles.length - vRows.length);
      var r2 = insertRows(s, wbx, vRows[vRows.length - 1], extraV, null); s = r2[0]; wbx = r2[1];
      for (var i = 1; i <= extraV; i++) vRows.push(vRows[vRows.length - 1] + 1);
      var pRefs = CELLS.personal.slice();
      for (var j = 1; j <= extraP; j++) pRefs.push("D" + (lastP + j));
      var lugar = [data.space, data.poblacion, data.provincia].filter(function (v, i, all) { return str(v) && !all.slice(0, i).some(function (p) { return norm(p).indexOf(norm(v)) >= 0; }); }).join(", ");

      s = setCell(s, CELLS.fecha, data.date ? longDate(data.date) : "");
      s = setCell(s, CELLS.llegada, data.llegada);
      s = setCell(s, CELLS.evento, data.event);
      if (!data.venue) s = setCell(s, CELLS.recinto, data.space);
      s = setCell(s, CELLS.poblacion, data.poblacion);
      s = setCell(s, CELLS.provincia, data.provincia);
      s = setCell(s, CELLS.cliente, data.client);
      if (!(V && vv("TELÉFONO PROD. LOCAL"))) CELLS.limpiar.forEach(function (ref) { s = setCell(s, ref, ""); });

      pRefs.forEach(function (ref, i) { s = setCell(s, ref, staff[i] || ""); });

      vRows.forEach(function (row, i) {
        var v = vehicles[i] || {};
        s = setCell(s, "B" + row, v.driver);
        s = setCell(s, "E" + row, v.plate);
        s = setCell(s, "H" + row, v.name);
      });

      var linkRecinto = vv("MAPS RECINTO") || vv("MAPS ACCESO CARGA") || (lugar ? mapsUrl(lugar) : "");
      if (linkRecinto) {
        rels = setLink(rels, LINKS.recinto, linkRecinto);
        rels = setLink(rels, LINKS.acceso, vv("MAPS ACCESO CARGA") || linkRecinto);
        rels = setLink(rels, LINKS.parking, vv("MAPS PARKING") || linkRecinto);
      }
      if (lugar) rels = setLink(rels, LINKS.hotel, mapsUrl("hoteles cerca de " + lugar));
      rels = setLink(rels, LINKS.meteo, meteoUrl(data.poblacion, lugar));

      zip.file(SHEET, s);
      zip.file(SHEET_RELS, rels);
      zip.file(WORKBOOK, wbx);
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

  function sha256Hex(text) {
    if (!root.crypto || !root.crypto.subtle) return Promise.reject(new Error("Este navegador no permite comprobar la contraseña."));
    return root.crypto.subtle.digest("SHA-256", new TextEncoder().encode(text)).then(function (buf) {
      return Array.prototype.map.call(new Uint8Array(buf), function (b) { return ("0" + b.toString(16)).slice(-2); }).join("");
    });
  }

  function isStatus(v) { return /reserva|firme|tbc|confirm|opci[oó]n|pendiente|\?/i.test(str(v)); }

  // Separa los eventos de un día: cada columna de PRODUCCION 27 es un evento distinto
  function dayEvents(date, fallback) {
    var rows = (root.workbookSheets || {})["PRODUCCION 27"] || [], toISO = root.dateISO, out = [];
    if (toISO) {
      for (var r = 2; r < rows.length; r++) {
        if (toISO(rows[r] && rows[r][0]) !== date) continue;
        var width = Math.max((rows[r] || []).length, (rows[r + 1] || []).length, (rows[r + 2] || []).length);
        for (var c = 2; c < width; c++) {
          var cell = function (i) { var v = str((rows[i] || [])[c]); return v === "." ? "" : v; };
          var e = { event: cell(r), status: cell(r + 1), space: cell(r + 2) };
          if (e.event || e.status || e.space) out.push(e);
        }
        break;
      }
    }
    if (!out.length && fallback) out.push({ event: fallback.event, status: fallback.client, space: fallback.space });
    return out;
  }

  function splitPlace(space) {
    var m = str(space).match(/^(.*?)\s*\(([^)]+)\)\s*$/);
    return m ? { poblacion: m[1], provincia: m[2] } : { poblacion: str(space), provincia: "" };
  }

  /* ---------- Listado de recintos (Google Sheet compartido de solo lectura) ---------- */
  var RECINTOS_ID_DEFAULT = "1XfgjgyqQp2dE2haA1amK6CQ1kSsfB13PS3MGM0Vk_2A";
  var recintosPromise = null;
  function recintosId() { var c = root.TUA2027_DATA_SOURCE || {}; return c.recintosId || RECINTOS_ID_DEFAULT; }

  function loadRecintos() {
    if (recintosPromise) return recintosPromise;
    recintosPromise = new Promise(function (resolve, reject) {
      var cb = "tuaRecintos_" + Date.now(), script = document.createElement("script");
      var timer = setTimeout(function () { done(); reject(new Error("tiempo de espera")); }, 20000);
      function done() { clearTimeout(timer); delete root[cb]; if (script.parentNode) script.parentNode.removeChild(script); }
      root[cb] = function (payload) {
        done();
        if (!payload || payload.status !== "ok" || !payload.table) { reject(new Error("no se pudo leer")); return; }
        var cols = (payload.table.cols || []).map(function (c) { return str(c.label); });
        var list = (payload.table.rows || []).map(function (row) {
          var o = {};
          cols.forEach(function (h, i) { var c = row && row.c && row.c[i]; if (h) o[h] = c ? str(c.v != null ? c.v : c.f) : ""; });
          return o;
        }).filter(function (o) { return str(o["NOMBRE RECINTO"]); });
        resolve(list);
      };
      script.onerror = function () { done(); reject(new Error("bloqueado")); };
      script.src = "https://docs.google.com/spreadsheets/d/" + encodeURIComponent(recintosId()) + "/gviz/tq?sheet=RECINTOS&headers=1&tqx=" +
        encodeURIComponent("out:json;responseHandler:" + cb) + "&_=" + Date.now();
      document.head.appendChild(script);
    });
    recintosPromise.catch(function () { recintosPromise = null; });
    return recintosPromise;
  }

  function venueLabel(v) { return str(v["NOMBRE RECINTO"]) + (str(v["POBLACIÓN"]) ? " — " + str(v["POBLACIÓN"]) : ""); }

  function bestVenue(list, ev) {
    var STOP = { "de": 1, "del": 1, "la": 1, "el": 1, "los": 1, "las": 1, "y": 1, "en": 1, "municipal": 1 };
    var words = function (t) { return norm(t).split(/[^a-z0-9]+/).filter(function (w) { return w.length >= 3 && !STOP[w]; }); };
    var evWords = words([ev.event, ev.space].join(" ")), evText = " " + norm([ev.event, ev.space].join(" ")) + " ";
    var best = null, bestScore = 0, byTown = {};
    list.forEach(function (v) {
      var town = norm(v["POBLACIÓN"]);
      if (town && evText.indexOf(" " + town + " ") >= 0) byTown[town] = (byTown[town] || 0) + 1;
    });
    list.forEach(function (v) {
      var vw = words([v["NOMBRE RECINTO"], v["OTROS NOMBRES"]].join(" ")), town = norm(v["POBLACIÓN"]);
      var hits = vw.filter(function (w, i) { return vw.indexOf(w) === i && evWords.indexOf(w) >= 0 && w !== town; }).length;
      var townHit = town && evText.indexOf(" " + town + " ") >= 0;
      var score = hits * 2 + (townHit ? 1 : 0) + (townHit && byTown[town] === 1 ? 1 : 0);
      if (score > bestScore || (score === bestScore && best && Number(v["Nº HR"] || 0) > Number(best["Nº HR"] || 0))) { best = v; bestScore = score; }
    });
    return bestScore >= 2 ? best : null;
  }

  // recintos de la población del evento (por el texto del planning o por la población escrita)
  function townVenues(list, ev, poblacion) {
    var evText = " " + norm([ev.event, ev.space, poblacion].join(" ")) + " ";
    return list.filter(function (v) {
      var town = norm(v["POBLACIÓN"]);
      return town && evText.indexOf(" " + town + " ") >= 0;
    }).sort(function (a, b) { return Number(b["Nº HR"] || 0) - Number(a["Nº HR"] || 0); });
  }

  function venueSummary(v) {
    if (!v) return "";
    var parts = [], links = ["MAPS RECINTO", "MAPS ACCESO CARGA", "MAPS PARKING"].filter(function (k) { return str(v[k]); }).length;
    if (str(v["DIRECCIÓN"]) || str(v["DIRECCIÓN ACCESO CARGA"])) parts.push("dirección");
    if (links) parts.push(links + (links > 1 ? " enlaces" : " enlace") + " de Maps");
    if (["AFORO", "TIPO ESCENARIO", "MEDIDAS ESCENARIO", "ACCESO CARGA", "UBICACIÓN CONTROL"].some(function (k) { return str(v[k]); })) parts.push("datos del espacio");
    var tomas = ["TOMA SONIDO", "TOMA ILUMINACIÓN", "TOMA VIDEO", "TOMA MOTORES", "TOMA CONTROL", "TOMA CAÑONES", "TOMA PARKING BUS"].filter(function (k) { return str(v[k]); }).length;
    if (tomas) parts.push(tomas + (tomas > 1 ? " tomas" : " toma") + " eléctricas");
    if (str(v["PRODUCCIÓN LOCAL"])) parts.push("producción local");
    if (str(v["ELÉCTRICO CONTACTO"])) parts.push("eléctrico");
    return parts.length ? "Se rellenará del listado: " + parts.join(", ") + ". (Tomas y contactos son del último evento allí: revísalos.)" : "Este recinto no tiene más datos en el listado.";
  }

  function open(date) {
    var model = root.model || {};
    var base = (model.events || []).find(function (x) { return x.date === date; });
    if (!base) return;
    if (!root.JSZip) { alert("No se pudo cargar el generador de Excel. Revisa la conexión y recarga la página."); return; }
    injectStyles();

    var events = dayEvents(date, base);
    var workerNames = (model.workers || []).map(function (w) { return w.name; });
    var staff = [], vehicles = [], venues = [], venue = null;

    var wrap = document.createElement("div");
    wrap.className = "hr-backdrop";
    wrap.innerHTML = '<div class="hr-modal" role="dialog" aria-modal="true" aria-label="Generar ficha HR">' +
      '<h3>Ficha HR</h3><div class="muted">' + esc(longDate(date)) + "</div>" +
      (events.length > 1 ? '<h4>Evento del día</h4><select class="field" id="hrEvent" style="width:100%">' +
        events.map(function (e, i) { return '<option value="' + i + '">' + esc([e.event, e.space].filter(str).join(" · ") || "Evento " + (i + 1)) + "</option>"; }).join("") + "</select>" : "") +
      '<h4>Datos del evento</h4><div class="hr-grid">' +
      '<label style="grid-column:1/-1">Actividad o evento<input class="field" id="hrEvento"></label>' +
      '<label style="grid-column:1/-1">Recinto del listado<input class="field" id="hrVenue" list="hrVenues" placeholder="Cargando recintos…" autocomplete="off"></label><datalist id="hrVenues"></datalist>' +
      '<label style="grid-column:1/-1;display:none" id="hrTownWrap"><span id="hrTownLabel">Recintos en esta población</span><select class="field" id="hrTown" style="width:100%"></select></label>' +
      '<div class="hr-note" id="hrVenueInfo" style="grid-column:1/-1;margin-top:-2px"></div>' +
      '<label>Recinto (texto en la ficha)<input class="field" id="hrRecinto"></label><label>Cliente<input class="field" id="hrCliente"></label>' +
      '<label>Población<input class="field" id="hrPoblacion"></label><label>Provincia<input class="field" id="hrProvincia"></label>' +
      '<label style="grid-column:1/-1">Llegada camiones / comienzo montaje<input class="field" id="hrLlegada" placeholder="Ej. 08:00"></label></div>' +
      '<div class="hr-note" id="hrStatus"></div>' +
      '<div id="hrLists"></div>' +
      '<div class="hr-note">Revisa los datos: puedes cambiar cualquier campo antes de descargar. El resto se completa en el Excel.</div>' +
      '<h4>Contraseña</h4><input class="field" type="password" id="hrPassword" autocomplete="off" style="width:100%" placeholder="Necesaria para generar la ficha">' +
      '<div class="hr-note" id="hrPassError" style="color:var(--coral)"></div>' +
      '<div class="hr-actions"><button class="button" data-hr="cancel">Cancelar</button><button class="button primary" data-hr="go">Descargar Excel</button></div></div>';
    document.body.appendChild(wrap);
    var $ = function (sel) { return wrap.querySelector(sel); };

    function setVenue(v) {
      venue = v || null;
      $("#hrVenue").value = venue ? venueLabel(venue) : "";
      $("#hrVenueInfo").textContent = venue ? venueSummary(venue) : (venues.length ? "Sin recinto del listado: elígelo en la lista o búscalo arriba; si no, se usa el texto del planning." : "");
      if (typeof townList !== "undefined" && townList.length) { var ti = townList.indexOf(venue); $("#hrTown").value = ti >= 0 ? String(ti) : ""; }
      if (venue) {
        $("#hrRecinto").value = str(venue["NOMBRE RECINTO"]);
        if (str(venue["POBLACIÓN"])) $("#hrPoblacion").value = str(venue["POBLACIÓN"]);
        if (str(venue["PROVINCIA"])) $("#hrProvincia").value = str(venue["PROVINCIA"]);
      }
    }
    function currentEvent() { return events[$("#hrEvent") ? Number($("#hrEvent").value) : 0] || {}; }
    var townList = [];
    function refreshTown() {
      townList = venues.length ? townVenues(venues, currentEvent(), $("#hrPoblacion").value) : [];
      var wrapT = $("#hrTownWrap");
      if (!townList.length) { wrapT.style.display = "none"; return; }
      wrapT.style.display = "";
      $("#hrTownLabel").textContent = "Recintos fichados en " + str(townList[0]["POBLACIÓN"]) + " (" + townList.length + ")";
      $("#hrTown").innerHTML = '<option value="">— Elige recinto —</option>' + townList.map(function (v, i) {
        return '<option value="' + i + '">' + esc(str(v["NOMBRE RECINTO"])) + (Number(v["Nº HR"]) > 1 ? " · " + v["Nº HR"] + " HR" : "") + "</option>";
      }).join("");
      var idx = venue ? townList.indexOf(venue) : -1;
      $("#hrTown").value = idx >= 0 ? String(idx) : "";
    }
    loadRecintos().then(function (list) {
      venues = list;
      $("#hrVenues").innerHTML = list.map(function (v) { return '<option value="' + esc(venueLabel(v)) + '"></option>'; }).join("");
      $("#hrVenue").placeholder = "Escribe para buscar entre " + list.length + " recintos";
      setVenue(bestVenue(list, currentEvent()));
      refreshTown();
    }).catch(function () {
      $("#hrVenue").placeholder = "No se pudo cargar el listado de recintos";
      $("#hrVenue").disabled = true;
    });
    $("#hrTown").addEventListener("change", function () {
      var v = townList[Number($("#hrTown").value)];
      if ($("#hrTown").value !== "" && v) setVenue(v);
    });
    $("#hrPoblacion").addEventListener("change", function () { if (!venue) refreshTown(); });
    $("#hrVenue").addEventListener("change", function () {
      var val = $("#hrVenue").value, v = venues.find(function (x) { return venueLabel(x) === val; });
      if (v) setVenue(v);
      else if (!str(val)) { venue = null; $("#hrVenueInfo").textContent = ""; render(); }
    });

    function render() {
      var ev = currentEvent();
      var place = splitPlace(ev.space);
      $("#hrEvento").value = ev.event || "";
      $("#hrRecinto").value = ev.space || "";
      $("#hrCliente").value = isStatus(ev.status) ? "" : ev.status || "";
      $("#hrPoblacion").value = place.poblacion;
      $("#hrProvincia").value = place.provincia;
      $("#hrStatus").textContent = isStatus(ev.status) ? "Estado en planificación: " + ev.status : "";
      var tokens = eventTokens({ event: ev.event, space: ev.space });
      staff = preselect(dayAssignments(model.workers, date), tokens);
      vehicles = preselect(dayAssignments(model.vehicles, date), tokens).map(function (v) {
        v.plate = extractPlate(v.name);
        v.driver = workerNames.find(function (n) { return norm(n) && norm(v.value).indexOf(norm(n)) >= 0; }) || "";
        return v;
      });
      if (venues.length) { setVenue(bestVenue(venues, ev)); refreshTown(); }
      $("#hrLists").innerHTML = "<h4>Personal TUA (el primero marcado será el responsable)</h4>" + checklist(staff, "staff") +
        "<h4>Vehículos</h4>" + checklist(vehicles, "veh", function (v) { return v.driver ? " · conductor: " + esc(v.driver) : ""; });
    }
    render();
    if ($("#hrEvent")) $("#hrEvent").addEventListener("change", render);

    $("#hrPassword").addEventListener("keydown", function (e) { if (e.key === "Enter") $('[data-hr="go"]').click(); });
    function close() { if (wrap.parentNode) wrap.parentNode.removeChild(wrap); }
    wrap.addEventListener("click", function (e) { if (e.target === wrap) close(); });
    $('[data-hr="cancel"]').addEventListener("click", close);
    $('[data-hr="go"]').addEventListener("click", function () {
      var btn = this;
      var picked = function (attr, list) {
        return Array.prototype.filter.call(wrap.querySelectorAll("[data-" + attr + "]"), function (c) { return c.checked; })
          .map(function (c) { return list[Number(c.getAttribute("data-" + attr))]; });
      };
      var data = {
        date: date, event: $("#hrEvento").value, client: $("#hrCliente").value, space: $("#hrRecinto").value,
        poblacion: $("#hrPoblacion").value, provincia: $("#hrProvincia").value, llegada: $("#hrLlegada").value,
        staff: picked("staff", staff).map(function (x) { return x.name; }),
        vehicles: picked("veh", vehicles),
        venue: venue
      };
      btn.disabled = true; btn.textContent = "Comprobando…";
      $("#hrPassError").textContent = "";
      sha256Hex($("#hrPassword").value).then(function (hash) {
        if (hash !== PASSWORD_SHA256) {
          var err = new Error("Contraseña incorrecta."); err.password = true; throw err;
        }
        btn.textContent = "Generando…";
        return generate(data);
      }).then(close).catch(function (err) {
        if (err.password) {
          $("#hrPassError").textContent = err.message;
          $("#hrPassword").value = ""; $("#hrPassword").focus();
          btn.disabled = false; btn.textContent = "Descargar Excel";
          return;
        }
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

  // En móvil el detalle queda debajo de toda la agenda: al tocar un día, bajar hasta él
  if (typeof document !== "undefined") {
    document.addEventListener("click", function (e) {
      var target = e.target.closest && e.target.closest("[data-event-date], #productionCalendar [data-calendar-date]");
      if (!target || !root.matchMedia || !root.matchMedia("(max-width: 620px)").matches) return;
      setTimeout(function () {
        var panel = document.querySelector("#view-produccion .detail-panel");
        if (!panel) return;
        panel.style.scrollMarginTop = "72px";
        panel.scrollIntoView({ behavior: "smooth", block: "start" });
      }, 50);
    });
  }

  var api = { open: open, fillWorkbook: fillWorkbook, fileName: fileName, extractPlate: extractPlate, longDate: longDate };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.TUAFichaHR = api;
})(typeof window !== "undefined" ? window : globalThis);
