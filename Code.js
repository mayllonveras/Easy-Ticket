/**
 * @fileoverview Script do Easy Ticket que pesquisa e-mails do Gmail para compras confirmadas
 * da Expresso Guanabara, cria eventos no calendário "Viagens", salva PDFs em uma pasta
 * específica do Drive e marca as threads com a label "Passagem Guanabara agendada".
 * A duração da viagem é controlada por uma variável no topo. As rotas são determinadas
 * dinamicamente a partir do array CITIES, para que o nome do anexo e a rota no corpo
 * do e-mail sejam considerados de forma coerente. Logs foram incluídos para depuração.
 * 
 * Autor: Mayllon Veras (mayllonveras@gmail.com)
 */

/**
 * @typedef {Object} CityData
 * @property {string} code - Sigla da cidade (ex.: "PHB").
 * @property {string} name - Nome completo da cidade (ex.: "PARNAIBA - PI").
 */

/** 
 * Número de horas de duração padrão para cada viagem. 
 * @constant {number} 
 */
const TRIP_DURATION_HOURS = 3;

/** 
 * Lista de cidades. Adicione/edite conforme necessário.
 * @constant {CityData[]} 
 */
const CITIES = [
  { code: "THE", name: "TERESINA - PI" },
  { code: "PHB", name: "PARNAIBA - PI" },
  { code: "PIR", name: "PIRIPIRI - PI" }
];

/**
 * Gera uma subexpressão de regex unindo todos os valores de nome do array CITIES.
 * @returns {string} Regex pattern para representar todas as cidades no array CITIES.
 */
function buildCitiesRegexPart() {
  const escapedList = CITIES.map(city => city.name.replace(/\s*-\s*/g, '\\s*\\-\\s*'));
  return `(${escapedList.join('|')})`;
}

/**
 * Cria ou recupera a label "Passagem Guanabara agendada" no Gmail.
 * @returns {GoogleAppsScript.Gmail.GmailLabel} A label de agendamento.
 */
function ensureLabelExists() {
  const labelName = "Passagem Guanabara agendada";
  const label = GmailApp.getUserLabelByName(labelName);
  if (label) return label;
  return GmailApp.createLabel(labelName);
}

/**
 * Cria ou recupera a pasta "Bilhetes - passagens Guanabara" no Drive.
 * @returns {GoogleAppsScript.Drive.Folder} A pasta para armazenar PDFs.
 */
function ensureTicketsFolderExists() {
  const folderName = "Bilhetes - passagens Guanabara";
  const folderSearch = DriveApp.getFoldersByName(folderName);
  if (folderSearch.hasNext()) return folderSearch.next();
  return DriveApp.createFolder(folderName);
}

/**
 * Cria ou recupera o calendário "Viagens" usando o serviço avançado do Calendar.
 * @returns {string} O ID do calendário.
 */
function ensureCalendarExists() {
  const summaryName = "Viagens";
  const calList = Calendar.CalendarList.list().items || [];
  let calendarId = null;
  for (let i = 0; i < calList.length; i++) {
    if (calList[i].summary === summaryName) {
      calendarId = calList[i].id;
      break;
    }
  }
  if (!calendarId) {
    const newCal = Calendar.Calendars.insert({ summary: summaryName });
    calendarId = newCal.id;
  }
  return calendarId;
}

/**
 * Converte uma string em português com data e hora (ex.: "16 de janeiro de 2025 às 13:11")
 * em um objeto Date do JavaScript.
 * @param {string} ptBrDatetimeString - A string de data/hora em português.
 * @returns {Date|null} O objeto Date convertido ou null se falhar.
 */
function parseDatetimePtBr(ptBrDatetimeString) {
  const monthsMap = {
    janeiro: 0, fevereiro: 1, março: 2, abril: 3, maio: 4, junho: 5,
    julho: 6, agosto: 7, setembro: 8, outubro: 9, novembro: 10, dezembro: 11
  };
  const regex = /(?:[a-zçã-ú-]+,\s*)?(\d{1,2})\s+de\s+([a-zçã-ú]+)\s+de\s+(\d{4})\s+às\s+(\d{1,2}):(\d{2})/i;
  const match = ptBrDatetimeString.match(regex);
  if (!match) return null;
  const day = parseInt(match[1], 10);
  const month = monthsMap[match[2].toLowerCase()];
  const year = parseInt(match[3], 10);
  const hour = parseInt(match[4], 10);
  const minute = parseInt(match[5], 10);
  return new Date(year, month, day, hour, minute);
}

/**
 * Retorna a sigla de uma cidade com base no array CITIES.
 * @param {string} cityText - Texto contendo cidade e estado (ex.: "PARNAIBA - PI").
 * @returns {string} Sigla correspondente ou "???" se não encontrada.
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
 * Extrai a rota (origem e destino) do nome do arquivo do anexo com base nas cidades definidas.
 * @param {string} filename - Nome do arquivo do anexo.
 * @returns {Object|null} Objeto contendo origin e destination ou null se não encontrado.
 */
function extractRouteFromFilename(filename) {
  for (let i = 0; i < CITIES.length; i++) {
    for (let j = 0; j < CITIES.length; j++) {
      if (i === j) continue; // Ignorar rota de uma cidade para ela mesma
      const origin = CITIES[i].name;
      const destination = CITIES[j].name;
      const expectedPattern = `${origin.replace(/\s*-\s*/g, '\\s*\\-\\s*')}\\s*\\-\\s*${destination.replace(/\s*-\s*/g, '\\s*\\-\\s*')}`;
      const regex = new RegExp(expectedPattern, 'i');
      if (regex.test(filename)) {
        return { origin, destination };
      }
    }
  }
  return null;
}

/**
 * Cria ou identifica todos os recursos necessários e pesquisa mensagens no Gmail
 * sobre compras confirmadas da Expresso Guanabara, criando eventos no Calendar
 * e salvando PDFs no Drive. Marca as threads processadas com a label.
 */
function runGuanabaraTicketScript() {
  const label = ensureLabelExists();
  const calendarId = ensureCalendarExists();
  const ticketsFolder = ensureTicketsFolderExists();

  const threads = GmailApp.search('subject:"Expresso Guanabara - Compra confirmada com sucesso" newer_than:1m');
  console.log(`Total de threads encontradas: ${threads.length}`);

  const cityRegexPart = buildCitiesRegexPart();
  // Monta a regex para achar algo como "CIDADE1 <span>...</span> CIDADE2"
  const routeRegex = new RegExp(`${cityRegexPart}\\s*<span[^>]*>[^<]*<\\/span>\\s*${cityRegexPart}`, 'gi');

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const hasLabel = thread.getLabels().some(lbl => lbl.getName() === label.getName());
    if (hasLabel) {
      console.log(`Thread #${i} já tem label, pulando...`);
      continue;
    }

    let eventCreated = false;
    const messages = thread.getMessages();
    console.log(`Thread #${i} com ${messages.length} mensagens.`);

    for (let m = 0; m < messages.length; m++) {
      const msg = messages[m];
      const body = msg.getBody();
      const datesFound = body.match(/(?:[a-zçã-ú-]+,\s*)?\d{1,2}\s+de\s+[a-zçã-ú]+\s+de\s+\d{4}\s+às\s+\d{2}:\d{2}/gi) || [];
      
      let routesFound = [];
      let matchRoute;
      while ((matchRoute = routeRegex.exec(body)) !== null) {
        if (matchRoute.length === 3) { // matchRoute[1]: origin, matchRoute[2]: destination
          routesFound.push({
            origin: matchRoute[1].trim(),
            destination: matchRoute[2].trim()
          });
        }
      }

      const limit = Math.min(datesFound.length, routesFound.length);
      console.log(`Mensagem #${m}: datas=${datesFound.length}, rotas=${routesFound.length}, limit=${limit}`);

      const attachments = msg.getAttachments({
        includeInlineImages: false,
        includeAttachments: true
      });
      console.log(`Mensagem #${m} possui ${attachments.length} anexos.`);

      const routeMap = {};
      for (let a = 0; a < attachments.length; a++) {
        const filename = attachments[a].getName();
        console.log(`  Anexo #${a}: ${filename}`);

        const route = extractRouteFromFilename(filename);
        if (route) {
          const routeKey = `${route.origin}|${route.destination}`;
          routeMap[routeKey] = attachments[a];
          console.log(`    Mapeado routeMap[${routeKey}]`);
        }
      }

      for (let k = 0; k < limit; k++) {
        const dateStr = datesFound[k];
        const startDate = parseDatetimePtBr(dateStr);
        if (!startDate) {
          console.log(`    Data inválida: ${dateStr}`);
          continue;
        }
        const endDate = new Date(startDate.getTime() + TRIP_DURATION_HOURS * 60 * 60 * 1000);

        const route = routesFound[k];
        const origin = route.origin;
        const destination = route.destination;
        console.log(`    Trecho #${k}: origem=${origin}, destino=${destination}`);

        const routeKey = `${origin}|${destination}`;
        console.log(`    routeKey=${routeKey}`);

        const eventTitle = `Viagem ${origin.replace(/\s*-\s*PI/i, "")} : ${destination.replace(/\s*-\s*PI/i, "")}`;

        const eventData = {
          summary: eventTitle,
          start: { dateTime: startDate.toISOString() },
          end: { dateTime: endDate.toISOString() }
        };

        if (routeMap[routeKey]) {
          console.log(`    Encontrado anexo para ${routeKey}`);
          const originSig = getCityCode(origin);
          const destinationSig = getCityCode(destination);
          const formattedDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

          const pdfBlob = routeMap[routeKey].copyBlob();
          const newFilename = `Bilhete Guanabara ${originSig}>${destinationSig}-${formattedDate}.pdf`;
          pdfBlob.setName(newFilename);

          const driveFile = ticketsFolder.createFile(pdfBlob);
          driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          const fileUrl = `https://drive.google.com/file/d/${driveFile.getId()}/view?usp=sharing`;

          eventData.attachments = [{
            fileId: driveFile.getId(),
            fileUrl,
            title: driveFile.getName()
          }];

          console.log(`    PDF salvo no Drive: ${newFilename}`);
          console.log(`    PDF anexado ao evento: ${fileUrl}`);
        } else {
          console.log(`    Nenhum anexo no routeMap para ${routeKey}`);
        }

        Calendar.Events.insert(eventData, calendarId, { supportsAttachments: true });
        console.log(`    Evento criado: ${eventTitle}`);
        eventCreated = true;
      }
    }

    if (eventCreated) {
      console.log(`Adicionando label na thread #${i}`);
      thread.addLabel(label);
    }
  }

  console.log("Script finalizado.");
}
