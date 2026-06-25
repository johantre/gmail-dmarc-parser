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
    .setHeader(CardService.newCardHeader().setTitle("DMARC Parser (Claude) - " + gzAttachment.getName()));

  var generalSection = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph()
      .setText("📄 DMARC rapport gevonden en gedecomprimeerd. Hieronder een overzicht van de records."));

  cardBuilder.addSection(generalSection);

  cardBuilder.addSection(
    CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(`📤 <b>Rapport verzonden door:</b><br>${orgName} (${orgEmail})`))
  );

  var claudeRaw = callClaude(buildPrompt(orgName, orgEmail, records, ns));
  var claudeParsed = parseClaudeResponse(claudeRaw);
  var claudeSection = CardService.newCardSection().setHeader("🤖 Uitleg door Claude");
  if (claudeParsed.oordeel || claudeParsed.uitleg || claudeParsed.actie) {
    if (claudeParsed.oordeel) {
      claudeSection.addWidget(CardService.newTextParagraph().setText('<b>' + claudeParsed.oordeel + '</b>'));
    }
    if (claudeParsed.uitleg) {
      claudeSection.addWidget(CardService.newTextParagraph().setText(claudeParsed.uitleg));
    }
    if (claudeParsed.actie) {
      claudeSection.addWidget(CardService.newKeyValue()
        .setTopLabel("Actie")
        .setContent(claudeParsed.actie));
    }
  } else {
    claudeSection.addWidget(CardService.newTextParagraph()
      .setText(claudeRaw || "⚠️ Claude kon geen uitleg genereren. Controleer de ANTHROPIC_API_KEY in Script Properties."));
  }
  cardBuilder.addSection(claudeSection);

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


function parseClaudeResponse(text) {
  var result = { oordeel: null, uitleg: null, actie: null };
  if (!text) return result;
  text.split('\n').forEach(function(line) {
    if (line.startsWith('OORDEEL:')) result.oordeel = line.replace('OORDEEL:', '').trim();
    else if (line.startsWith('UITLEG:')) result.uitleg = line.replace('UITLEG:', '').trim();
    else if (line.startsWith('ACTIE:')) result.actie = line.replace('ACTIE:', '').trim();
  });
  return result;
}

function buildPrompt(orgName, orgEmail, records, ns) {
  var lines = [
    'Je bent een e-mail security expert. Analyseer dit DMARC rapport en geef je antwoord EXACT in dit formaat (drie regels, geen extra tekst):',
    'OORDEEL: [kies één van: ✅ Alles in orde | ⚠️ Aandacht vereist | ❌ Actie vereist]',
    'UITLEG: [2 tot 3 zinnen uitleg in eenvoudig Nederlands wat dit rapport betekent]',
    'ACTIE: [concrete actie of "Geen actie nodig"]',
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
      + ', aantal=' + safeGetChildText(row, 'count', ns)
      + ', dispositie=' + (policyEvaluated ? safeGetChildText(policyEvaluated, 'disposition', ns) : 'onbekend')
      + ', DKIM=' + (policyEvaluated ? safeGetChildText(policyEvaluated, 'dkim', ns) : 'onbekend')
      + ', SPF=' + (policyEvaluated ? safeGetChildText(policyEvaluated, 'spf', ns) : 'onbekend'));
  });
  return lines.join('\n');
}

function callClaude(prompt) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) {
    Logger.log('❌ Claude: geen ANTHROPIC_API_KEY gevonden in Script Properties');
    return null;
  }
  try {
    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5',
        max_tokens: 1024,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });
    var responseText = response.getContentText();
    Logger.log('Claude HTTP status: ' + response.getResponseCode());
    Logger.log('Claude response: ' + responseText.substring(0, 500));
    var json = JSON.parse(responseText);
    if (!json.content || !json.content[0]) {
      Logger.log('❌ Claude: geen content in response');
      return null;
    }
    return json.content[0].text;
  } catch(e) {
    Logger.log('❌ Claude fout: ' + e);
    return null;
  }
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
