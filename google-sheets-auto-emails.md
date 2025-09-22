# Google Sheets Auto Emails

Este projeto contém um script em Google Apps Script que envia emails automáticos a partir de dados no Google Sheets.

## Estrutura esperada

### Aba `Envio`
Contém os dados dos destinatários:
```
Nome | Sender | Recipient | Type | Tempo | Data | Manager | A/O | Cargo | Enviado
```

### Aba `Texto`
Contém os templates de email:
```
Type | Título | Texto
```

## Script (Code.gs)

```javascript
function generateAndSendMessages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sendSheet = ss.getSheetByName("Envio"); 
  const templateSheet = ss.getSheetByName("Texto"); 

  const sendData = sendSheet.getDataRange().getValues();
  const templateData = templateSheet.getDataRange().getValues();

  const headersSend = sendData.shift(); 
  const headersTemplate = templateData.shift();

  const colEnviado = headersSend.indexOf("Enviado") + 1; 

  // Dicionário de templates
  const templates = {};
  templateData.forEach(row => {
    templates[row[0]] = { title: row[1], text: row[2] };
  });

  sendData.forEach((row, idx) => {
    let rowObj = {};
    headersSend.forEach((h, i) => rowObj[h] = row[i]);

    // Ignorar já enviados
    if (rowObj["Enviado"] === "SIM") return;

    let type = rowObj["Type"];
    let template = templates[type];
    if (!template) return;

    // Substituições
    let message = template.text
      .replace(/\(Nome\)/g, rowObj["Nome"])
      .replace(/\(Cargo\)/g, rowObj["Cargo"])
      .replace(/\(Manager\)/g, rowObj["Manager"])
      .replace(/\(A\/O\)/g, rowObj["A/O"])
      .replace(/\(Tempo\)/g, rowObj["Tempo"])
      .replace(/\(Data\)/g, formatDatePT(rowObj["Data"]))
      .replace(/\(Date\)/g, formatDateEN(rowObj["Data"]));

    let subject = template.title.replace(/\(Nome\)/g, rowObj["Nome"]);

    let recipient = rowObj["Recipient"];
    let sender = rowObj["Sender"];
    if (!recipient || !sender) return;

    try {
      GmailApp.createDraft(recipient, subject, "", {
        from: sender,
        name: "Equipe RH",
        htmlBody: message.replace(/\n/g, "<br>")
      });

      // Marca como enviado
      if (colEnviado > 0) sendSheet.getRange(idx + 2, colEnviado).setValue("SIM");
      sendSheet.getRange(idx + 2, 1, 1, headersSend.length).setBackground("#d9ead3");

    } catch (e) {
      Logger.log("Erro ao enviar para " + recipient + ": " + e.message);
    }
  });
}

// Formatar datas
function formatDatePT(dateStr) {
  if (!dateStr) return "";
  let d = new Date(dateStr);
  return d.toLocaleDateString("pt-BR", { day: "numeric", month: "long", year: "numeric" });
}

function formatDateEN(dateStr) {
  if (!dateStr) return "";
  let d = new Date(dateStr);
  let day = d.getDate();
  let suffix = (day >= 11 && day <= 13) ? "th" : {1:"st",2:"nd",3:"rd"}[day % 10] || "th";
  let month = d.toLocaleDateString("en-US", { month: "long" });
  return `${month} ${day}${suffix}, ${d.getFullYear()}`;
}
```
