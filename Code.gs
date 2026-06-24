function safeGetChildText(parent, childName, ns) {
  var child = parent.getChild(childName, ns);
  return child ? child.getText() : '[niet beschikbaar]';
}

function buildAddOn(e) {
  var messageId = e.gmail.messageId;
  var message = GmailApp.getMessageById(messageId);
  var attachments = message.getAttachments();

  var gzAttachment = attachments.find(function(att) {
    var name = att.getName().toLowerCase();
    return name.endsWith('.xml.gz') || name.endsWith('.zip');
  });

  if (!gzAttachment) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("Geen geldige bijlage"))
      .addSection(
        CardService.newCardSection()
          .addWidget(CardService.newTextParagraph()
            .setText("Deze e-mail bevat geen .xml.gz of .zip DMARC rapport."))
      )
      .build();
  }

  var xmlText = gzAttachment.getName().toLowerCase().endsWith('.zip')
    ? extractXmlFromZip(gzAttachment)
    : decompressGzWithPako(gzAttachment);
  if (!xmlText) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("Fout bij decompressie"))
      .addSection(
        CardService.newCardSection()
          .addWidget(CardService.newTextParagraph()
            .setText("Kon het bestand niet decompressen."))
      )
      .build();
  }

  var doc = XmlService.parse(xmlText);
  var root = doc.getRootElement(); // <feedback>
  var ns = root.getNamespace();    // DMARC XML gebruikt namespaces

  // Zoek het <report_metadata> element
  var metadata = root.getChild('report_metadata', ns);

  var orgName = "Onbekend";
  var orgEmail = "Onbekend";

  if (metadata) {
    var orgNameElement = metadata.getChild('org_name', ns);
    var emailElement = metadata.getChild('email', ns);

    if (orgNameElement) {
      orgName = orgNameElement.getText();
    }
    if (emailElement) {
      orgEmail = emailElement.getText();
    }
  }
  var records = root.getChildren('record', ns);
  if (!records || records.length === 0) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("Geen records gevonden"))
      .addSection(
        CardService.newCardSection()
          .addWidget(CardService.newTextParagraph()
            .setText("Er zijn geen DMARC records gevonden in de XML."))
      )
      .build();
  }

  var cardBuilder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("DMARC XML Parser - " + gzAttachment.getName()));

  var generalSection = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
      .setText("📄 DMARC rapport gevonden en gedecomprimeerd. Hieronder een overzicht van de records."));

  cardBuilder.addSection(generalSection);

  cardBuilder.addSection(
    CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(`📤 <b>Rapport verzonden door:</b><br>${orgName} (${orgEmail})`))
  );

  var geminiExplanation = callGemini(buildGeminiPrompt(orgName, orgEmail, records, ns));
  cardBuilder.addSection(
    CardService.newCardSection()
      .setHeader("🤖 Uitleg door Gemini")
      .addWidget(CardService.newTextParagraph()
        .setText(geminiExplanation || "⚠️ Gemini kon geen uitleg genereren. Controleer de GEMINI_API_KEY in Script Properties en de uitvoeringslogboeken."))
  );

  records.forEach(function(record, i) {
    var row = record.getChild('row', ns);
    if (!row) return;  // skip als geen row

    var sourceIp = safeGetChildText(row, 'source_ip', ns);

    var policyEvaluated = row.getChild('policy_evaluated', ns);

    // Check of policyEvaluated bestaat
    var disposition = policyEvaluated ? safeGetChildText(policyEvaluated, 'disposition', ns) : '[niet beschikbaar]';
    var dkim = policyEvaluated ? safeGetChildText(policyEvaluated, 'dkim', ns) : '[niet beschikbaar]';
    var spf = policyEvaluated ? safeGetChildText(policyEvaluated, 'spf', ns) : '[niet beschikbaar]';

    var count = safeGetChildText(row, 'count', ns);

    var recordSection = CardService.newCardSection()
      .setHeader("Record " + (i + 1) + " — IP: " + sourceIp)
      .addWidget(CardService.newKeyValue()
        .setTopLabel("Aantal e-mails")
        .setContent("📧 " + count))
      .addWidget(CardService.newKeyValue()
        .setTopLabel("Dispositie")
        .setContent("📋 " + disposition))
      .addWidget(CardService.newKeyValue()
        .setTopLabel("DKIM")
        .setContent(dkim === "pass" ? "✅ geslaagd" : "❌ mislukt!"))
      .addWidget(CardService.newKeyValue()
        .setTopLabel("SPF")
        .setContent(spf === "pass" ? "✅ geslaagd" : "❌ mislukt!"));

    cardBuilder.addSection(recordSection);
  });

  return cardBuilder.build();
}

function reverseLookup(ip) {
  var reversedIp = ip.split('.').reverse().join('.') + '.in-addr.arpa';
  var url = 'https://dns.google/resolve?name=' + reversedIp + '&type=PTR';

  try {
    var response = UrlFetchApp.fetch(url);
    var json = JSON.parse(response.getContentText());

    if (json.Answer && json.Answer.length > 0) {
      // Neem eerste PTR resultaat
      return json.Answer[0].data.replace(/\.$/, '');  // . wegknippen aan eind
    }
  } catch (e) {
    Logger.log('Reverse lookup failed: ' + e);
  }
  return "Onbekend";
}



// Parse XML met XmlService
function parseDmarcXml(xmlText) {
  var doc = XmlService.parse(xmlText);
  var root = doc.getRootElement();

  var reportMetadata = root.getChild("report_metadata");
  var orgName = reportMetadata.getChildText("org_name");
  var dateRange = reportMetadata.getChild("date_range");
  var begin = new Date(Number(dateRange.getChildText("begin")) * 1000);
  var end = new Date(Number(dateRange.getChildText("end")) * 1000);

  var records = root.getChildren("record");
  var spfPass = 0;
  var dkimPass = 0;
  var total = records.length;

  records.forEach(function(rec) {
    var authResults = rec.getChild("auth_results");
    var spf = authResults.getChild("spf");
    var dkim = authResults.getChild("dkim");

    if (spf && spf.getChildText("result") === "pass") spfPass++;
    if (dkim && dkim.getChildText("result") === "pass") dkimPass++;
  });

  var dmarcPass = (spfPass === total) && (dkimPass === total);

  return {
    orgName: orgName,
    beginDate: begin.toDateString(),
    endDate: end.toDateString(),
    totalRecords: total,
    spfPassCount: spfPass,
    dkimPassCount: dkimPass,
    isPass: dmarcPass
  };
}

// Maak kaart met samenvatting
function buildSummaryCard(summary, filename) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("DMARC rapport: " + filename))
    .addSection(
      CardService.newCardSection()
        .addWidget(CardService.newKeyValue()
          .setTopLabel("Organisatie")
          .setContent(summary.orgName))
        .addWidget(CardService.newKeyValue()
          .setTopLabel("Rapportperiode")
          .setContent(summary.beginDate.toLocaleDateString() + " - " + summary.endDate.toLocaleDateString()))
        .addWidget(CardService.newKeyValue()
          .setTopLabel("Domein")
          .setContent(summary.domain))
        .addWidget(CardService.newKeyValue()
          .setTopLabel("Totaal berichten")
          .setContent(summary.totalMessages.toString()))
        .addWidget(CardService.newKeyValue()
          .setTopLabel("DKIM failures")
          .setContent(summary.dkimFails.toString()))
        .addWidget(CardService.newKeyValue()
          .setTopLabel("SPF failures")
          .setContent(summary.spfFails.toString()))
    )
    .build();
}


function buildGeminiPrompt(orgName, orgEmail, records, ns) {
  var lines = [
    'Je bent een e-mail security expert. Leg in eenvoudige Nederlandse taal uit wat dit DMARC rapport betekent voor de e-mailbezorging van het domein.',
    'Geef een korte samenvatting van 3 tot 5 zinnen en vermeld duidelijk of er acties nodig zijn.',
    'Gebruik geen markdown, geen opsommingstekens, gewoon lopende tekst.',
    '',
    'Rapport van: ' + orgName + ' (' + orgEmail + ')',
    'Aantal records: ' + records.length,
    ''
  ];
  records.forEach(function(record, i) {
    var row = record.getChild('row', ns);
    if (!row) return;
    var policyEvaluated = row.getChild('policy_evaluated', ns);
    lines.push('Record ' + (i + 1) + ':'
      + ' IP=' + safeGetChildText(row, 'source_ip', ns)
      + ', aantal=' + safeGetChildText(record, 'count', ns)
      + ', dispositie=' + (policyEvaluated ? safeGetChildText(policyEvaluated, 'disposition', ns) : 'onbekend')
      + ', DKIM=' + (policyEvaluated ? safeGetChildText(policyEvaluated, 'dkim', ns) : 'onbekend')
      + ', SPF=' + (policyEvaluated ? safeGetChildText(policyEvaluated, 'spf', ns) : 'onbekend'));
  });
  return lines.join('\n');
}

function callGemini(prompt) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    Logger.log('❌ Gemini: geen GEMINI_API_KEY gevonden in Script Properties');
    return null;
  }
  var url = 'https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash-lite:generateContent?key=' + apiKey;
  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
      muteHttpExceptions: true
    });
    var responseText = response.getContentText();
    Logger.log('Gemini HTTP status: ' + response.getResponseCode());
    Logger.log('Gemini response: ' + responseText.substring(0, 800));
    var json = JSON.parse(responseText);
    if (!json.candidates || !json.candidates[0]) {
      Logger.log('❌ Gemini: geen candidates in response');
      return null;
    }
    return json.candidates[0].content.parts[0].text;
  } catch(e) {
    Logger.log('❌ Gemini fout: ' + e);
    return null;
  }
}

function listGeminiModels() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  var response = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1/models?key=' + apiKey, { muteHttpExceptions: true });
  Logger.log(response.getContentText());
}

function authorizeUrlFetch() {
  loadPako();  // dwingt het script om toestemming te vragen
}

function loadPako() {
  var response = UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/pako/2.1.0/pako.min.js');
  eval(response.getContentText());
}

function extractXmlFromZip(attachment) {
  try {
    var unzipped = Utilities.unzip(attachment.copyBlob());
    var xmlBlob = unzipped.find(function(blob) {
      return blob.getName().toLowerCase().endsWith('.xml');
    });
    if (!xmlBlob) return null;
    return xmlBlob.getDataAsString();
  } catch(e) {
    Logger.log('❌ Fout bij ZIP extractie: ' + e);
    return null;
  }
}

function decompressGzWithPako(attachment) {
  // Zorg dat pako geladen is
  if (typeof pako === 'undefined') {
    loadPako();
  }

  var bytes = attachment.getBytes(); // krijg raw gzip data als byte array
  try {
    // decompressen met pako
    var decompressed = pako.ungzip(new Uint8Array(bytes), { to: 'string' });
    Logger.log('Decompressed XML (eerste 200 chars): ' + decompressed.substring(0, 200));
    return decompressed;
  } catch(e) {
    Logger.log('❌ Fout bij pako decompressie: ' + e);
    return null;
  }
}
