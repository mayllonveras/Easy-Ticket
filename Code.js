/**
 * @fileoverview Easy Ticket – lê e-mails da Expresso Guanabara, cria eventos
 * na agenda “Viagens”, salva PDFs no Drive e rotula os e-mails processados.
 *
 * @typedef {Object} CityData
 * @property {string} code   Sigla da cidade.
 * @property {string} name   Nome completo “CIDADE - UF”.
 */

/* -------------------- CONFIG -------------------- */

/** alertas de lembrete ― ex.: "30m", "1h", "1.5h" */
const ALERTS = ["30m", "1h", "1.5h"];

/** duração padrão do deslocamento (h) */
const TRIP_DURATION_HOURS = 3;

/** cidades tratadas pelo script */
const CITIES = [
  { code: "THE", name: "TERESINA - PI" },
  { code: "PHB", name: "PARNAIBA - PI" },
  { code: "PIR", name: "PIRIPIRI - PI" }
];

/* -------------------- FUNÇÕES AUXILIARES -------------------- */

/** parte regex com todas as cidades escapadas. */
function buildCitiesPart() {
  return `(${CITIES.map(c => c.name.replace(/\s*-\s*/g, "\\s*-\\s*")).join("|")})`;
}

/** regex p/ rota (origem … destino) dentro dos blocos “Viagem de Ida/Volta” */
function buildRouteRegex() {
  const cities = buildCitiesPart();                         // (CIDADE - UF|CIDADE - UF|…)
  // Procura “Viagem de Ida” OU “Viagem de Volta”, depois a primeira cidade
  // e, mais adiante no mesmo bloco, a segunda cidade (pode haver HTML entre elas).
  return new RegExp(`Viagem\\s+de\\s+(?:Ida|Volta)[\\s\\S]*?${cities}[\\s\\S]*?${cities}`, "gi");
}

/** label “Passagem Guanabara agendada” */
function ensureLabel() {
  const n = "Passagem Guanabara agendada";
  return GmailApp.getUserLabelByName(n) || GmailApp.createLabel(n);
}

/** pasta “Bilhetes - passagens Guanabara” */
function ensureFolder() {
  const n = "Bilhetes - passagens Guanabara";
  const f = DriveApp.getFoldersByName(n);
  return f.hasNext() ? f.next() : DriveApp.createFolder(n);
}

/** calendário “Viagens” */
function ensureCalendar() {
  const n = "Viagens";
  const c = (Calendar.CalendarList.list().items || []).find(i => i.summary === n);
  return c ? c.id : Calendar.Calendars.insert({ summary: n }).id;
}

/**
 * converte string pt-BR (“29 jul, terça 10:01” | “16 de janeiro de 2025 às 13:11”) em Date.
 * @param {string} txt
 * @returns {Date|null}
 */
function parseDatetime(txt) {
  const full  = /(\d{1,2})\s+de\s+([a-zçã-ú]+)\s+de\s+(\d{4})\s+às\s+(\d{1,2}):(\d{2})/i;
  const short = /(\d{1,2})\s+(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[^0-9]*?(\d{1,2}):(\d{2})/i;
  const monthsFull = { janeiro:0, fevereiro:1, março:2, abril:3, maio:4, junho:5,
                       julho:6, agosto:7, setembro:8, outubro:9, novembro:10, dezembro:11 };
  const monthsShort= { jan:0, fev:1, mar:2, abr:3, mai:4, jun:5,
                       jul:6, ago:7, set:8, out:9, nov:10, dez:11 };

  let m = txt.match(full);
  if (m) return new Date(+m[3], monthsFull[m[2].toLowerCase()], +m[1], +m[4], +m[5]);
  m = txt.match(short);
  if (m) {
    const year = new Date().getFullYear();
    return new Date(year, monthsShort[m[2].toLowerCase()], +m[1], +m[3], +m[4]);
  }
  return null;
}

/** map sigla ←→ nome completo */
function getCityCode(name) {
  const c = CITIES.find(c => name.includes(c.name));
  return c ? c.code : "???";
}

/** extrai rota do nome do anexo */
function routeFromFilename(name) {
  for (const o of CITIES) {
    for (const d of CITIES) {
      if (o === d) continue;
      const p = `${o.name.replace(/\s*[-–—]\s*/g, "\\s*-\\s*")}\\s*-\\s*${d.name.replace(/\s*[-–—]\s*/g, "\\s*-\\s*")}`;
      if (new RegExp(p, "i").test(name)) return { origin:o.name, destination:d.name };
    }
  }
  return null;
}

/** “30m” → 30; “1.5h” → 90 */
function alertToMinutes(s) {
  return s.endsWith("m") ? Math.round(parseFloat(s))
       : s.endsWith("h") ? Math.round(parseFloat(s) * 60)
       : 0;
}

/* -------------------- NOVA FUNÇÃO -------------------- */

/** normaliza nomes de cidade para comparação (maiúsculas, hífens, espaços) */
function normalizeCity(txt) {
  return txt
    .replace(/\s*[-–—]\s*/g, " - ") // hífens ou travessões → espaço-hífen-espaço
    .replace(/\s+/g, " ")           // reduz espaços múltiplos
    .trim()
    .toUpperCase();
}

/* -------------------- SCRIPT PRINCIPAL -------------------- */

function runGuanabaraTicketScript() {
  const lbl     = ensureLabel();
  const calId   = ensureCalendar();
  const folder  = ensureFolder();
  const routeRe = buildRouteRegex();
  const threads = GmailApp.search('subject:"Expresso Guanabara - Compra confirmada com sucesso" newer_than:1m');

  console.log(`Total de threads encontradas: ${threads.length}`);

  threads.forEach((th, ti) => {
    if (th.getLabels().some(l => l.getName() === lbl.getName())) {
      console.log(`Thread #${ti} já tem label, pulando...`);
      return;
    }

    console.log(`Thread #${ti} com ${th.getMessageCount()} mensagens.`);
    let created = false;

    th.getMessages().forEach((msg, mi) => {
      const body = msg.getBody();

      // --- Datas ---
      const dateMatches =
        body.match(/(?:\d{1,2}\s+de\s+[a-zçã-ú]+\s+de\s+\d{4}\s+às\s+\d{2}:\d{2})|(?:\d{1,2}\s+(?:jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[^0-9]{1,15}\d{2}:\d{2})/gi) || [];

      // --- Rotas ---
      const routes = [];
      let m; routeRe.lastIndex = 0;
      while ((m = routeRe.exec(body)) !== null)
        routes.push({ origin: m[1].trim(), destination: m[2].trim() });

      console.log(`Mensagem #${mi}: datas=${dateMatches.length}, rotas=${routes.length}`);

      // --- Anexos mapeados ---
      const routeMap = {};
      msg.getAttachments({ includeAttachments:true }).forEach((at,i) => {
        const r = routeFromFilename(at.getName());
        console.log(`  Anexo #${i}: ${at.getName()}`);
        if (r) {
          const key = `${normalizeCity(r.origin)}|${normalizeCity(r.destination)}`;
          routeMap[key] = at;
          console.log(`    Mapeado routeMap[${key}]`);
        }
      });

      const limit = Math.min(dateMatches.length, routes.length);
      for (let i = 0; i < limit; i++) {
        const start = parseDatetime(dateMatches[i]);
        if (!start) continue;

        let origin      = routes[i].origin;
        let destination = routes[i].destination;

        // procura anexo correspondente
        const key        = `${normalizeCity(origin)}|${normalizeCity(destination)}`;
        const swappedKey = `${normalizeCity(destination)}|${normalizeCity(origin)}`;

        let attachment = routeMap[key];

        // fallback: se rota no corpo veio invertida
        if (!attachment && routeMap[swappedKey]) {
          attachment  = routeMap[swappedKey];
          origin      = routes[i].destination;
          destination = routes[i].origin;
          console.log(`    ⚠️  Rota invertida detectada, usando ${origin} ➜ ${destination}`);
        }

        const end = new Date(start.getTime() + TRIP_DURATION_HOURS * 3600000);

        console.log(`    Evento #${i}: ${origin} ➜ ${destination}`);

        // lembretes
        const overrides = ALERTS.map(a => ({ method:"popup", minutes:alertToMinutes(a) }))
                                .filter(o => o.minutes);

        const ev = {
          summary: `Viagem ${origin.replace(/\s*-\s*PI/i,"")} : ${destination.replace(/\s*-\s*PI/i,"")}`,
          start:   { dateTime: start.toISOString() },
          end:     { dateTime: end.toISOString() },
          reminders:{ useDefault:false, overrides }
        };

        // anexo do bilhete
        if (attachment) {
          const sigO = getCityCode(origin), sigD = getCityCode(destination);
          const tag  = Utilities.formatDate(start, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
          const blob = attachment.copyBlob().setName(`Bilhete Guanabara ${sigO}>${sigD}-${tag}.pdf`);
          const file = folder.createFile(blob).setSharing(
            DriveApp.Access.ANYONE_WITH_LINK,
            DriveApp.Permission.VIEW
          );
          ev.attachments = [{
            fileId: file.getId(),
            fileUrl:`https://drive.google.com/file/d/${file.getId()}/view?usp=sharing`,
            title: file.getName()
          }];
        }

        Calendar.Events.insert(ev, calId, { supportsAttachments:true });
        created = true;
        console.log("      ✓ evento criado");
      }
    });

    if (created) {
      th.addLabel(lbl);
      console.log(`Label aplicada na thread #${ti}`);
    }
  });

  console.log("Script finalizado.");
}
