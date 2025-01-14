/**
 * @fileoverview Script do Easy Ticket que pesquisa e-mails do Gmail para compras confirmadas da Expresso Guanabara,
 * cria eventos no calendário "Viagens", salva PDFs em uma pasta específica do Drive e marca as threads com a label
 * "Passagem Guanabara agendada".
 * 
 * @author
 *   Mayllon Veras (mayllonveras@gmail.com)
 */

/**
 * @typedef {Object} CityData
 * @property {string} code - Sigla da cidade (por exemplo, "PHB").
 * @property {string} name - Nome completo da cidade com estado (por exemplo, "PARNAIBA - PI").
 */

/** @constant {CityData[]} */
const CITIES = [
  { code: "PHB", name: "PARNAIBA - PI" },
  { code: "PIR", name: "PIRIPIRI - PI" }
];

/**
 * Cria a label "Passagem Guanabara agendada" no Gmail, se não existir.
 * @returns {GoogleAppsScript.Gmail.GmailLabel} A instância da label.
 */
function ensureLabelExists() {
  const labelName = "Passagem Guanabara agendada";
  const existingLabel = GmailApp.getUserLabelByName(labelName);
  return existingLabel || GmailApp.createLabel(labelName);
}

/**
 * Cria ou encontra a pasta "Bilhetes - passagens Guanabara" no Google Drive.
 * @returns {GoogleAppsScript.Drive.Folder} A referência para a pasta de bilhetes.
 */
function ensureTicketsFolderExists() {
  const folderName = "Bilhetes - passagens Guanabara";
  const folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

/**
 * Cria ou encontra o calendário "Viagens" usando o serviço avançado do Google Calendar.
 * @returns {string} O ID do calendário.
 */
function ensureCalendarExists() {
  const summaryName = "Viagens";
  const list = Calendar.CalendarList.list().items || [];
  let calendarId = null;
  for (let i = 0; i < list.length; i++) {
    if (list[i].summary === summaryName) {
      calendarId = list[i].id;
      break;
    }
  }
  if (!calendarId) {
    const newCalendar = Calendar.Calendars.insert({ summary: summaryName });
    calendarId = newCalendar.id;
  }
  return calendarId;
}

/**
 * Converte uma data em português (ex: "16 de janeiro de 2025 às 13:11") em objeto Date.
 * @param {string} ptBrDatetimeString - A string contendo data e hora em português.
 * @returns {Date|null} O objeto Date convertido ou null se falhar.
 */
function parseDatetimePtBr(ptBrDatetimeString) {
  const months = {
    janeiro: 0, fevereiro: 1, março: 2, abril: 3, maio: 4, junho: 5,
    julho: 6, agosto: 7, setembro: 8, outubro: 9, novembro: 10, dezembro: 11
  };
  const regex = /(?:[a-zçã-ú-]+,\s*)?(\d{1,2})\s+de\s+([a-zçã-ú]+)\s+de\s+(\d{4})\s+às\s+(\d{1,2}):(\d{2})/i;
  const match = ptBrDatetimeString.match(regex);
  if (!match) return null;
  const day = Number(match[1]);
  const month = months[match[2].toLowerCase()];
  const year = Number(match[3]);
  const hour = Number(match[4]);
  const minute = Number(match[5]);
  return new Date(year, month, day, hour, minute);
}

/**
 * Retorna a sigla da cidade com base no nome completo definido no array CITIES.
 * @param {string} cityText - Texto da cidade (ex: "PARNAIBA - PI").
 * @returns {string} A sigla da cidade se encontrada, caso contrário "???".
 */
function getCityCode(cityText) {
  for (let i = 0; i < CITIES.length; i++) {
    if (cityText.includes(CITIES[i].name)) {
      return CITIES[i].code;
    }
  }
  return "???";
}

/**
 * Encontra, ou cria se necessário, todos os recursos necessários (label, calendário, pasta) e,
 * em seguida, pesquisa mensagens do Gmail sobre compra confirmada da Expresso Guanabara para criar
 * eventos no calendário "Viagens" e armazenar PDFs de passagens no Drive.
 */
function runGuanabaraTicketScript() {
  const label = ensureLabelExists();
  const calendarId = ensureCalendarExists();
  const ticketsFolder = ensureTicketsFolderExists();
  const threads = GmailApp.search('subject:"Expresso Guanabara - Compra confirmada com sucesso" newer_than:7d');

  for (let t = 0; t < threads.length; t++) {
    const thread = threads[t];
    const threadLabels = thread.getLabels().map(lbl => lbl.getName());
    if (threadLabels.includes(label.getName())) continue;

    let createdEvent = false;
    const messages = thread.getMessages();

    for (let m = 0; m < messages.length; m++) {
      const body = messages[m].getBody();
      const datesFound = body.match(/(?:[a-zçã-ú-]+,\s*)?\d{1,2}\s+de\s+[a-zçã-ú]+\s+de\s+\d{4}\s+às\s+\d{2}:\d{2}/gi) || [];
      const routesFound = body.match(/(PARNAIBA\s*-\s*PI|PIRIPIRI\s*-\s*PI)\s*<span[^>]*>[^<]*<\/span>\s*(PARNAIBA\s*-\s*PI|PIRIPIRI\s*-\s*PI)/gi) || [];
      const limit = Math.min(datesFound.length, routesFound.length);

      const attachments = messages[m].getAttachments({ includeInlineImages: false, includeAttachments: true });
      const routeMap = {};

      for (let a = 0; a < attachments.length; a++) {
        const filename = attachments[a].getName();
        if (filename.includes("PARNAIBA - PI - PIRIPIRI - PI")) {
          routeMap["PARNAIBA - PI|PIRIPIRI - PI"] = attachments[a];
        } else if (filename.includes("PIRIPIRI - PI - PARNAIBA - PI")) {
          routeMap["PIRIPIRI - PI|PARNAIBA - PI"] = attachments[a];
        }
      }

      for (let i = 0; i < limit; i++) {
        const start = parseDatetimePtBr(datesFound[i]);
        if (!start) continue;
        const end = new Date(start.getTime() + 3 * 3600000);
        const routeMatch = routesFound[i].match(/(PARNAIBA\s*-\s*PI|PIRIPIRI\s*-\s*PI)\s*<span[^>]*>[^<]*<\/span>\s*(PARNAIBA\s*-\s*PI|PIRIPIRI\s*-\s*PI)/i);
        if (!routeMatch) continue;

        const origin = routeMatch[1].trim();
        const destination = routeMatch[2].trim();
        const routeKey = `${origin}|${destination}`;
        const eventTitle = `Viagem ${origin.replace(" - PI", "")} : ${destination.replace(" - PI", "")}`;
        const eventData = {
          summary: eventTitle,
          start: { dateTime: start.toISOString() },
          end: { dateTime: end.toISOString() }
        };

        if (routeMap[routeKey]) {
          const originSig = getCityCode(origin);
          const destSig = getCityCode(destination);
          const dateStr = Utilities.formatDate(start, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

          const pdfBlob = routeMap[routeKey].copyBlob();
          const newFilename = `Bilhete Guanabara ${originSig}>${destSig}-${dateStr}.pdf`;
          pdfBlob.setName(newFilename);

          const driveFile = ticketsFolder.createFile(pdfBlob);
          driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          const fileUrl = `https://drive.google.com/file/d/${driveFile.getId()}/view?usp=sharing`;
          eventData.attachments = [{
            fileId: driveFile.getId(),
            fileUrl,
            title: driveFile.getName()
          }];
        }

        Calendar.Events.insert(eventData, calendarId, { supportsAttachments: true });
        createdEvent = true;
      }
    }

    if (createdEvent) {
      thread.addLabel(label);
    }
  }
}
