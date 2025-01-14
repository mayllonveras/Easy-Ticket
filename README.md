# Easy Ticket - Encontrando facilmente as passagens da Expresso Guanabara no Google Calendar

Sobre a implantação do script do **Easy Ticket** no Google Apps Script. O script encontra-se no arquivo [Code.js](Code.js) e é responsável por:

1. Pesquisar no Gmail por e-mails de compra confirmada da Expresso Guanabara.  
2. Criar eventos no calendário **"Viagens"** (criando-o se não existir).  
3. Salvar PDFs das passagens na pasta **"Bilhetes - passagens Guanabara"** no Drive (criando-a se não existir).  
4. Marcar as threads do Gmail com a label **"Passagem Guanabara agendada"** (criando essa label se não existir).  

## Funcionalidades

- **Duração da viagem**: Controlada por meio de uma constante `TRIP_DURATION_HOURS` no código.
- **Pesquisa de E-mails**: Lê e-mails recentes (últimos 7 dias) com assunto `"Expresso Guanabara - Compra confirmada com sucesso"`.
- **Extração de Rota e Data/Hora**: Usa expressões regulares para identificar a data/hora (em português) e a rota (ex.: `PARNAIBA - PI → PIRIPIRI - PI`).
- **Criação de Eventos**:  
  - Duração baseada em `TRIP_DURATION_HOURS`.  
  - Inclui anexos se o PDF correspondente for encontrado.  
- **Organização dos PDFs**: Salva e renomeia os PDFs no Drive e anexa ao evento do Calendar.
- **Evita Reprocessamento**: Aplica a label `"Passagem Guanabara agendada"` para não repetir a operação na mesma thread.

## Requisitos

1. **Conta Google** com acesso a Gmail, Drive e Calendar.  
2. Acesso a [https://script.google.com/](https://script.google.com/) para criar e editar o projeto no Apps Script.  
3. **Serviço Avançado** do Google Calendar habilitado (para usar recursos da Calendar API, como `Calendar.Events.insert`).
4. **Serviço do Google Drive** habilitado (o script acessará o Drive para salvar arquivos).

## Passo a Passo de Instalação

1. **Criar Projeto no Google Apps Script**  
   - Acesse [https://script.google.com/](https://script.google.com/).  
   - Crie um novo projeto.

2. **Criar o Arquivo [Code.js](Code.js)**  
   - No projeto, crie um arquivo chamado `Code.js`.  
   - Copie e cole o conteúdo de [Code.js](Code.js).

3. **Ativar o Google Calendar API**  
   - No editor do Apps Script, acesse **Serviços avançados do Google**.  
   - Procure por **Calendar API** e ative.  
   - Se solicitado, ative também no Google Cloud Console vinculado ao projeto.

4. **Habilitar o Google Drive**  
   - No editor do Apps Script, acesse o menu **Serviços Avançados do Google**.  
   - Confirme que o **Drive API** está ativado. Caso não esteja:
     - Vá para o [Google Cloud Console](https://console.cloud.google.com/).  
     - Acesse **APIs e Serviços > Biblioteca**.  
     - Ative a **Drive API**.

5. **Configurar o Acionador (Trigger)**  
   - No editor do Apps Script, vá em **Acionadores**.  
   - Selecione a função `runGuanabaraTicketScript`.  
   - Tipo de acionador: **Time-driven**.  
   - Intervalo: **A cada hora** ou outro desejado.  
   - Salve o acionador para que o script seja executado periodicamente.

## Ajustes

- **TRIP_DURATION_HOURS**: No topo de [Code.js](Code.js), ajuste o valor para mudar a duração padrão da viagem.  
- **Cidades**: Se quiser incluir mais rotas, abra o array `CITIES` em [Code.js](Code.js) e adicione `{ code, name }`.  
- **Pasta no Drive**: Padrão: `"Bilhetes - passagens Guanabara"`.  
- **Calendário**: Padrão: **"Viagens"**.  
- **Label no Gmail**: Padrão: **"Passagem Guanabara agendada"**.

## Como Funciona

1. **Pesquisa E-mails**:  
   `'subject:"Expresso Guanabara - Compra confirmada com sucesso" newer_than:7d'`  
   Filtra os últimos 7 dias.

2. **Processa Threads**:  
   - Se a thread já tiver a label `"Passagem Guanabara agendada"`, ignora.  
   - Caso contrário, extrai datas, rotas e anexos correspondentes.

3. **Cria Evento**:  
   - Usa a data/hora identificada no e-mail, definindo a duração pela constante `TRIP_DURATION_HOURS`.  
   - Se houver PDF correspondente à rota, anexa ao evento.  

4. **Marca Threads**:  
   - Se criar ao menos um evento, marca a thread com **"Passagem Guanabara agendada"** para não repetir.

## Autor

- **Mayllon Veras** – [mayllonveras@gmail.com](mailto:mayllonveras@gmail.com)

Contribuições e sugestões são bem-vindas! Abra uma *issue* ou envie um *pull request* para melhorias.
