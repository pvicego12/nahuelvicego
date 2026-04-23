function controlHorarioPRO() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getActiveSheet();
  const data = hoja.getDataRange().getValues();

  if (data.length < 2) return;

  const headers = data[0].map(h => limpiarTexto(h));

  const colEmpleado = headers.findIndex(h => h.includes("empleado"));
  const colFecha = headers.findIndex(h => h.includes("fecha"));
  const colHora = headers.findIndex(h => h.includes("hora"));

  if (colEmpleado === -1 || colFecha === -1 || colHora === -1) {
    throw new Error("No encuentro las columnas");
  }

  const horasPorEmpleado = obtenerHorasPorEmpleado(ss);
  const licencias = obtenerLicencias(ss); // 🔥 NUEVO

  const map = {};
  const empleados = new Set();
  let fechas = [];

  // 🔹 AGRUPAR DATOS
  for (let i = 1; i < data.length; i++) {

    const emp = data[i][colEmpleado];
    const fechaObj = new Date(data[i][colFecha]);
    const hora = data[i][colHora];

    if (!emp || !fechaObj || !hora) continue;

    const fecha = formatFecha(fechaObj);
    const minutos = parseHoraREAL(hora);
    if (minutos === null) continue;

    empleados.add(emp);

    const key = emp + "_" + fecha;

    if (!map[key]) map[key] = [];
    map[key].push(minutos);

    fechas.push(fechaObj);
  }

  const minFecha = new Date(Math.min(...fechas));
  const maxFecha = new Date(Math.max(...fechas));

  const diasMes = generarDias(minFecha, maxFecha);
  const feriados = obtenerFeriadosRapido(minFecha, maxFecha);

  const resData = [];
  const filasEmpleado = [];

  empleados.forEach(emp => {

    const filaInicio = resData.length;

    // 🔵 fila azul
    resData.push([emp.toUpperCase(), "", "", "", "", "", ""]);

    diasMes.forEach(fechaObj => {

      const fecha = formatFecha(fechaObj);
      const key = emp + "_" + fecha;

      const dia = fechaObj.getDay();
      const esSabado = dia === 6;
      const esDomingo = dia === 0;
      const esFeriado = feriados.has(fecha);

      let total = 0;
      let ingresos = "";
      let egresos = "";
      let diferencia = 0;
      let estado = "";

      if (map[key]) {

        let arr = map[key].sort((a,b)=>a-b);
        arr = limpiarDuplicados(arr);

        let entrada = null;

        arr.forEach(m => {
          if (entrada === null) {
            entrada = m;
            ingresos += format(m) + " | ";
          } else {
            total += (m - entrada);
            egresos += format(m) + " | ";
            entrada = null;
          }
        });

        if (arr.length < 2) total = 0;

        const jornada = (horasPorEmpleado[emp] || 8) * 60;
        diferencia = total === 0 ? 0 : total - jornada;

      } else {

        const keyLic = emp + "_" + fecha;

        // 🔥 PRIORIDAD: LICENCIAS
        if (licencias[keyLic]) {
          estado = licencias[keyLic];

        } else if (esFeriado) {
          estado = "FERIADO";

        } else if (esSabado) {
          estado = "SABADO";

        } else if (esDomingo) {
          estado = "DOMINGO";

        } else {
          estado = "AUSENTE";
        }
      }

      resData.push([emp, fecha, ingresos, egresos, total, diferencia, estado]);
    });

    const filaFin = resData.length - 1;

    filasEmpleado.push({
      inicio: filaInicio,
      fin: filaFin
    });
  });

  let res = ss.getSheetByName("Resultado");
  if (!res) res = ss.insertSheet("Resultado");

  res.clear();

  res.appendRow([
    "Empleado","Fecha","Ingresos","Egresos","Minutos","Diferencia","Estado"
  ]);

  res.getRange(2,1,resData.length,resData[0].length).setValues(resData);

  // 🔥 SUMATORIA
  filasEmpleado.forEach(f => {
    const filaExcel = f.inicio + 2;
    const inicio = f.inicio + 3;
    const fin = f.fin + 2;
    const formula = `=SUM(F${inicio}:F${fin})`;
    res.getRange(filaExcel, 6).setFormula(formula);
  });

  pintar(res);
}


// -------------------------
function obtenerLicencias(ss) {

  const hoja = ss.getSheetByName("LICENCIAS");
  if (!hoja) return {};

  const data = hoja.getDataRange().getValues();
  let map = {};

  for (let i = 1; i < data.length; i++) {

    const emp = data[i][0];
    const fecha = formatFecha(new Date(data[i][1]));
    const tipo = data[i][2];

    if (!emp || !fecha || !tipo) continue;

    const key = emp + "_" + fecha;
    map[key] = tipo.toUpperCase();
  }

  return map;
}


// -------------------------
function limpiarDuplicados(arr) {

  if (arr.length <= 1) return arr;

  const limpio = [arr[0]];

  for (let i = 1; i < arr.length; i++) {
    const anterior = limpio[limpio.length - 1];
    const actual = arr[i];

    if (Math.abs(actual - anterior) <= 1) continue;

    limpio.push(actual);
  }

  return limpio;
}


// -------------------------
function obtenerFeriadosRapido(inicio, fin) {
  const cal = CalendarApp.getCalendarById("es.ar#holiday@group.v.calendar.google.com");
  const eventos = cal.getEvents(inicio, fin);

  const set = new Set();

  eventos.forEach(ev => {
    const f = Utilities.formatDate(ev.getStartTime(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    set.add(f);
  });

  return set;
}


// -------------------------
function generarDias(inicio, fin) {
  let dias = [];
  let actual = new Date(inicio);

  while (actual <= fin) {
    dias.push(new Date(actual));
    actual.setDate(actual.getDate() + 1);
  }

  return dias;
}


// -------------------------
function pintar(hoja) {

  const data = hoja.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    const fila = data[i];
    const empleado = fila[0];
    const fecha = fila[1];
    const estado = fila[6];
    const dif = fila[5];

    const celdaDif = hoja.getRange(i+1,6);
    const celdaEstado = hoja.getRange(i+1,7);
    const filaCompleta = hoja.getRange(i+1,1,1,7);

    if (empleado && !fecha) {
      filaCompleta.setBackground("#2f75b5");
      continue;
    }

    if (estado === "FERIADO") celdaEstado.setFontColor("#f4a261");
    else if (estado === "AUSENTE") celdaEstado.setFontColor("#ff0000");
    else if (estado === "SABADO" || estado === "DOMINGO") celdaEstado.setFontColor("#000000");
    else celdaEstado.setFontColor("#ff0000"); // licencias en rojo

    if (dif > 0) celdaDif.setBackground("#b6f2c2");
    else if (dif < 0) celdaDif.setBackground("#f7b2b2");
  }
}


// -------------------------
function formatFecha(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function limpiarTexto(txt) {
  return String(txt).toLowerCase().trim().replace(/\s+/g, " ");
}

function parseHoraREAL(v) {
  if (v instanceof Date) return v.getHours()*60 + v.getMinutes();
  if (typeof v === "number") return Math.round(v*24*60);
  if (typeof v === "string") {
    const p = v.split(":");
    if (p.length < 2) return null;
    return (+p[0])*60 + (+p[1]);
  }
  return null;
}

function format(min) {
  const h = Math.floor(min / 60);
  const m = min % 60;
  return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}`;
}

function obtenerHorasPorEmpleado(ss) {
  const hoja = ss.getSheetByName("HORARIOS");
  if (!hoja) return {};

  const data = hoja.getDataRange().getValues();
  let map = {};

  for (let i = 1; i < data.length; i++) {
    map[data[i][0]] = Number(data[i][1]);
  }

  return map;
}
