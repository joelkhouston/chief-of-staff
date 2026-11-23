/**
 * Phase 4: Multimodal PDF Processing & VeriCite Grounded Synthesis
 * Maps academic PDFs to Joel's 3 Policy Pillars and 5-Color Better Notes schema.
 */

const POLICY_TIMEZONE = 'America/Bogota';
const MAX_BATCH_RUNTIME_MS = 270000; // 4.5 min safety checkpoint (under 6-min limit)

/**
 * One-click setup: creates Drive drop-folders and the Policy_Research tracking tab.
 */
function setupPolicyFoldersAndTab() {
  const props = PropertiesService.getScriptProperties();

  const folders = {
    POLICY_QUEUE_FOLDER_ID: 'CoS_Policy_Queue',
    POLICY_ARCHIVE_FOLDER_ID: 'CoS_Policy_Archive',
    POLICY_NOTES_FOLDER_ID: 'CoS_Policy_Notes'
  };

  Object.entries(folders).forEach(([propKey, folderName]) => {
    const existingId = props.getProperty(propKey);
    if (existingId) {
      try {
        DriveApp.getFolderById(existingId);
        return;
      } catch (e) {
        // Folder was deleted; recreate below
      }
    }
    const iter = DriveApp.getFoldersByName(folderName);
    const folder = iter.hasNext() ? iter.next() : DriveApp.createFolder(folderName);
    props.setProperty(propKey, folder.getId());
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Policy_Research');
  if (!sheet) sheet = ss.insertSheet('Policy_Research');

  if (sheet.getLastRow() === 0) {
    const headers = [
      'Doc_ID', 'Processed_At', 'Citation_Key', 'Title', 'Primary_Pillar',
      'Relevance_1_to_5', 'Yellow_Claims', 'Green_Evidence', 'Blue_Methods',
      'Orange_Connections', 'Red_Gaps', 'Markdown_Note_URL'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  Logger.log('Policy folders (CoS_Policy_Queue, CoS_Policy_Archive, CoS_Policy_Notes) and Policy_Research tab ready.');
}

/**
 * Processes unread PDFs in CoS_Policy_Queue sequentially with VeriCite grounding,
 * saves Better Notes Markdown to CoS_Policy_Notes, logs to Policy_Research,
 * and archives the source PDF.
 */
function processPolicyQueue() {
  const startTime = Date.now();
  const props = PropertiesService.getScriptProperties();
  const queueId = props.getProperty('POLICY_QUEUE_FOLDER_ID');
  const archiveId = props.getProperty('POLICY_ARCHIVE_FOLDER_ID');
  const notesId = props.getProperty('POLICY_NOTES_FOLDER_ID');

  if (!queueId || !archiveId || !notesId) {
    throw new Error('Run setupPolicyFoldersAndTab() first.');
  }

  const queueFolder = DriveApp.getFolderById(queueId);
  const archiveFolder = DriveApp.getFolderById(archiveId);
  const notesFolder = DriveApp.getFolderById(notesId);
  const files = queueFolder.getFilesByType(MimeType.PDF);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const researchSheet = ss.getSheetByName('Policy_Research');
  let processedCount = 0;

  while (files.hasNext()) {
    // Checkpoint before hitting the 6-minute Apps Script execution limit
    if (Date.now() - startTime > MAX_BATCH_RUNTIME_MS) {
      ScriptApp.newTrigger('processPolicyQueue')
        .timeBased()
        .after(60 * 1000)
        .create();
      Logger.log(`Runtime checkpoint reached after ${processedCount} PDF(s). Continuation scheduled in 1 minute.`);
      return;
    }

    const file = files.next();
    const fileSize = file.getSize();
    const fileName = file.getName();

    if (fileSize > 50 * 1024 * 1024) {
      Logger.log(`Skipped "${fileName}" — exceeds 50 MB Gemini PDF limit.`);
      continue;
    }

    try {
      const extraction = extractPdfWithVeriCite_(file);
      const mdContent = buildBetterNotesMarkdown_(extraction, fileName);
const firstAuthor = (extraction.citation_key || 'Unknown').replace(/[0-9]+$/, '') || 'Unknown';
const fiveWords = (extraction.title || fileName.replace(/\.pdf$/i, ''))
  .trim()
  .split(/\s+/)
  .slice(0, 5)
  .join(' ')
  .replace(/[:;,.\-—]+$/, '');
const safeName = `${firstAuthor} - ${fiveWords}`.replace(/[\\/:*?"<>|]/g, '');
      const mdFile = notesFolder.createFile(`${safeName}.md`, mdContent, MimeType.PLAIN_TEXT);

      const countByColor = (color) =>
        (extraction.annotations || []).filter(a => a.color === color).length;

      const docId = 'DOC_' + Utilities.getUuid().slice(0, 6).toUpperCase();
      researchSheet.appendRow([
        docId,
        Utilities.formatDate(new Date(), POLICY_TIMEZONE, 'yyyy-MM-dd HH:mm'),
        extraction.citation_key || safeName,
        extraction.title || fileName,
        extraction.primary_pillar,
        extraction.relevance_score,
        countByColor('Yellow'),
        countByColor('Green'),
        countByColor('Blue'),
        countByColor('Orange'),
        countByColor('Red'),
        mdFile.getUrl()
      ]);

      file.moveTo(archiveFolder);
      processedCount++;

      // Brief pause to avoid HTTP 429 rate bursts
      Utilities.sleep(1500);
    } catch (err) {
      Logger.log(`Failed processing "${fileName}": ${err.message}`);
    }
  }

  Logger.log(`Policy queue run complete. Processed ${processedCount} PDF(s).`);
}

/**
 * Calls gemini-3.8-flash on a PDF (<20MB inline, 20-50MB via Files API)
 * enforcing VeriCite verbatim quote and page-number anchoring.
 */
function extractPdfWithVeriCite_(file) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('Missing GEMINI_API_KEY in Script Properties.');

  const blob = file.getBlob();
  const sizeBytes = file.getSize();
  let filePart = null;
  let uploadedFileName = null;

  if (sizeBytes < 20 * 1024 * 1024) {
    filePart = {
      inlineData: {
        mimeType: 'application/pdf',
        data: Utilities.base64Encode(blob.getBytes())
      }
    };
  } else {
    const uploaded = uploadLargePdfToFilesApi_(blob, apiKey);
    uploadedFileName = uploaded.name;
    filePart = {
      fileData: {
        mimeType: 'application/pdf',
        fileUri: uploaded.uri
      }
    };
  }

  const prompt = `Analyze this academic or policy PDF for Joel's K-12 School AI Policy research.

STRICT VERICITE GROUNDING RULES:
1. Evaluate each section independently for relevance to K-12 AI governance across Joel's 3 Pillars:
   - "Education Policies" (institutional rules, teacher/student guidelines, classroom adoption)
   - "Data Collection" (student privacy, telemetry, vendor data sharing, minor protections)
   - "Ethics" (algorithmic bias, equity, academic integrity, human agency)
2. Classify every extracted finding into Joel's 5-Color Zotero Better Notes schema:
   - "Yellow": Core arguments, theses, and policy claims (What)
   - "Green": Empirical evidence, RCT outcomes, and quantitative/qualitative findings
   - "Blue": Data collection methods, sample sizes, and evaluation frameworks (How)
   - "Orange": Connections, alignments, or contradictions with broader K-12 AI frameworks
   - "Red": Missing information, methodological limitations, blind spots, and follow-up citation gaps
3. Every annotation MUST include an exact verbatim quote from the PDF and its page number.
4. NEVER infer or fabricate unstated statistics, dates, or claims. Return null if metadata is absent.`;

  const schema = {
    type: 'OBJECT',
    properties: {
      title: { type: 'STRING', description: 'Exact document title.' },
      authors_year: { type: 'STRING', nullable: true, description: 'Authors and publication year, or null.' },
      citation_key: { type: 'STRING', description: 'AuthorYear key, e.g., Demszky2025.' },
      primary_pillar: {
        type: 'STRING',
        enum: ['Education Policies', 'Data Collection', 'Ethics'],
        description: 'Dominant research pillar.'
      },
      relevance_score: { type: 'INTEGER', description: 'Relevance to K-12 school AI policy from 1 (low) to 5 (core).' },
      executive_abstract: { type: 'STRING', nullable: true, description: '3-sentence grounded summary, or null.' },
      annotations: {
        type: 'ARRAY',
        description: 'Up to 15 highest-signal VeriCite-anchored extractions across the 5 colors.',
        items: {
          type: 'OBJECT',
          properties: {
            color: {
              type: 'STRING',
              enum: ['Yellow', 'Green', 'Blue', 'Orange', 'Red'],
              description: '5-color annotation category.'
            },
            pillar: {
              type: 'STRING',
              enum: ['Education Policies', 'Data Collection', 'Ethics'],
              description: 'Specific pillar this quote supports.'
            },
            verbatim_quote: { type: 'STRING', description: 'Exact verbatim quote copied from the PDF.' },
            page_num: { type: 'INTEGER', description: '1-based PDF page number where quote appears.' },
            analytical_note: { type: 'STRING', description: 'Concise 1-sentence takeaway linking quote to K-12 AI policy.' }
          },
          required: ['color', 'pillar', 'verbatim_quote', 'page_num', 'analytical_note']
        }
      }
    },
    required: ['title', 'citation_key', 'primary_pillar', 'relevance_score', 'annotations']
  };

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=${apiKey}`;
  const payload = {
    systemInstruction: { parts: [{ text: getSystemCharter() }] },
    contents: [{ role: 'user', parts: [filePart, { text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json',
      responseSchema: schema
    }
  };

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    const body = JSON.parse(res.getContentText());
    const tokens = body.usageMetadata ? body.usageMetadata.totalTokenCount : 0;

    const auditSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Audit_Log');
    if (auditSheet) {
      auditSheet.appendRow([
        Utilities.formatDate(new Date(), POLICY_TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
        'pdf_extract', 'gemini-3.8-flash', code, tokens, code === 200 ? 'EXTRACTED' : 'ERROR'
      ]);
    }

    if (code !== 200) {
      throw new Error(`Gemini PDF API Error (${code}): ${body.error ? body.error.message : res.getContentText()}`);
    }

    return JSON.parse(body.candidates[0].content.parts[0].text);
  } finally {
    if (uploadedFileName) {
      UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/${uploadedFileName}?key=${apiKey}`, {
        method: 'delete',
        muteHttpExceptions: true
      });
    }
  }
}

/**
 * Resumable upload helper for PDFs between 20 MB and 50 MB via Gemini Files API.
 */
function uploadLargePdfToFilesApi_(blob, apiKey) {
  const bytes = blob.getBytes();
  const initRes = UrlFetchApp.fetch(
    `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${apiKey}`,
    {
      method: 'post',
      headers: {
        'X-Goog-Upload-Protocol': 'resumable',
        'X-Goog-Upload-Command': 'start',
        'X-Goog-Upload-Header-Content-Length': String(bytes.length),
        'X-Goog-Upload-Header-Content-Type': 'application/pdf',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ file: { display_name: blob.getName() } }),
      muteHttpExceptions: true
    }
  );

  const uploadUrl = initRes.getAllHeaders()['x-goog-upload-url'] || initRes.getAllHeaders()['X-Goog-Upload-URL'];
  if (!uploadUrl) throw new Error('Failed to obtain Gemini Files API upload URL.');

  const finalRes = UrlFetchApp.fetch(uploadUrl, {
    method: 'post',
    headers: {
      'Content-Length': String(bytes.length),
      'X-Goog-Upload-Offset': '0',
      'X-Goog-Upload-Command': 'upload, finalize'
    },
    payload: bytes,
    muteHttpExceptions: true
  });

  const fileObj = JSON.parse(finalRes.getContentText()).file;
  if (!fileObj || !fileObj.uri) throw new Error('Gemini Files API upload returned no URI.');

  // Poll briefly if file state is PROCESSING
  for (let i = 0; i < 6; i++) {
    const checkRes = UrlFetchApp.fetch(
      `https://generativelanguage.googleapis.com/v1beta/${fileObj.name}?key=${apiKey}`,
      { muteHttpExceptions: true }
    );
    const meta = JSON.parse(checkRes.getContentText());
    if (meta.state === 'ACTIVE') return meta;
    Utilities.sleep(2000);
  }
  return fileObj;
}

/**
 * Formats VeriCite JSON extractions into Better Notes / Obsidian Markdown
 * grouped by Joel's 5-Color Schema.
 */
function buildBetterNotesMarkdown_(data, fallbackFileName) {
  const pillarTag = (data.primary_pillar || 'General').toLowerCase().replace(/\s+/g, '-');
  const lines = [
    '---',
    `title: "${(data.title || fallbackFileName).replace(/"/g, '\\"')}"`,
    `citation_key: "${data.citation_key || ''}"`,
    `authors_year: "${data.authors_year || 'Unknown'}"`,
    `primary_pillar: "${data.primary_pillar}"`,
    `relevance: ${data.relevance_score || 0}`,
    `tags: [ai-policy, ${pillarTag}, vericite-extracted]`,
    '---',
    '',
    `# ${data.title || fallbackFileName}`,
    '',
    `- **Citation Key:** \`@${data.citation_key}\``,
    `- **Authors / Year:** ${data.authors_year || 'Not stated in document'}`,
    `- **Primary Pillar:** ${data.primary_pillar} (Relevance: ${data.relevance_score}/5)`,
    ''
  ];

  if (data.executive_abstract) {
    lines.push('## Grounded Executive Summary', '', data.executive_abstract, '');
  }

  const colorSections = [
    { key: 'Yellow', heading: '## 🟡 Core Arguments & Claims (What)' },
    { key: 'Green', heading: '## 🟢 Empirical Evidence & Results' },
    { key: 'Blue', heading: '## 🔵 Data Collection Methods & Frameworks (How)' },
    { key: 'Orange', heading: '## 🟠 Cross-Paper Connections & Contradictions' },
    { key: 'Red', heading: '## 🔴 Missing Info, Blind Spots & Citation Gaps' }
  ];

  colorSections.forEach(sec => {
    const matches = (data.annotations || []).filter(a => a.color === sec.key);
    if (matches.length === 0) return;
    lines.push(sec.heading, '');
    matches.forEach(item => {
      lines.push(`> "${item.verbatim_quote}" — *@${data.citation_key}, p. ${item.page_num}*`);
      lines.push(`- **[${item.pillar}]** ${item.analytical_note}`);
      lines.push('');
    });
  });

  return lines.join('\n');
}

/**
 * VeriCite Two-Step Cross-Paper Synthesis:
 * Step 1: Evaluates extracted notes independently for relevance to target pillar.
 * Step 2: Synthesizes strictly from verified verbatim quotes—dropping unanchored claims.
 */
function synthesizePolicyPillar(pillarName) {
  const targetPillar = pillarName || 'Data Collection';
  const props = PropertiesService.getScriptProperties();
  const notesId = props.getProperty('POLICY_NOTES_FOLDER_ID');
  if (!notesId) throw new Error('Run setupPolicyFoldersAndTab() first.');

  const notesFolder = DriveApp.getFolderById(notesId);
  const files = notesFolder.getFilesByType(MimeType.PLAIN_TEXT);
  const corpus = [];

  while (files.hasNext()) {
    const f = files.next();
    if (f.getName().startsWith('Synthesis_')) continue;
    corpus.push(`=== SOURCE FILE: ${f.getName()} ===\n${f.getBlob().getDataAsText()}`);
  }

  if (corpus.length === 0) {
    throw new Error('No extracted notes found in CoS_Policy_Notes yet.');
  }

  const prompt = `You are executing a strict VeriCite Two-Step Literature Synthesis on the pillar: "${targetPillar}".

STEP 1 (INDEPENDENT RELEVANCE GATE):
Inspect each source note below independently. Keep ONLY verbatim quotes and findings that directly address "${targetPillar}" in K-12 schools.

STEP 2 (ANCHORED SYNTHESIS):
Write a structured policy synthesis brief from those verified passages only.
- Every factual or empirical sentence MUST end with an inline citation key and page number, e.g., (@AuthorYear, p. 12).
- Do NOT introduce any claim, statistic, or recommendation that lacks an explicit verbatim quote in the corpus below.
- Highlight contradictions across papers and list open policy gaps at the end.

CORPUS:\n${corpus.join('\n\n')}`;

  const schema = {
    type: 'OBJECT',
    properties: {
      verified_passage_count: { type: 'INTEGER', description: 'Number of passages passing Step 1 relevance gate.' },
      synthesis_markdown: { type: 'STRING', description: 'Complete Markdown synthesis brief with inline (@Key, p. X) citations.' },
      dropped_topics: { type: 'STRING', nullable: true, description: 'Topics excluded due to lacking verbatim evidence.' }
    },
    required: ['verified_passage_count', 'synthesis_markdown']
  };

  const result = callGeminiStructured_(
    'gemini-3.8-flash',
    getSystemCharter(),
    prompt,
    schema,
    'vericite_synthesis'
  );

  const dateStamp = Utilities.formatDate(new Date(), POLICY_TIMEZONE, 'yyyy-MM-dd');
  const slug = targetPillar.replace(/\s+/g, '_');
  const outFile = notesFolder.createFile(
    `Synthesis_${slug}_${dateStamp}.md`,
    result.synthesis_markdown,
    MimeType.PLAIN_TEXT
  );

  Logger.log(`VeriCite synthesis created (${result.verified_passage_count} verified passages): ${outFile.getUrl()}`);
  return outFile.getUrl();
}

/**
 * Schedules automated nightly PDF processing at 2:00 AM Cartagena time.
 */
function installPolicyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'processPolicyQueue') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('processPolicyQueue')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();

  Logger.log('Nightly Policy PDF trigger installed (2:00 AM).');
}