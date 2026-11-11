const SHEET_INPUT = "Input";
const SHEET_OUTPUT = "Output";
const SHEET_DOCUMENTS = "Document";
const SHEET_PAGES = "Page";
const SHEET_TASKS = "Tasks";
const SHEET_RESULTS = "Results";
const SHEET_KNOWLEDGE = "Knowledge";
const SHEET_TASK_TEMPLATE = "Task_Template";
const SHEET_PROJECT_INPUT = "Project_Input";
const SHEET_STORYLINE = "StoryLine";

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "health";
  let result;

  if (action === "health") {
    result = { ok: true, message: "Sheet Bridge is running", version: "v118_she_esg_attach_restore" };
    return outputResponse(e, result);
  }

  if (action === "test") {
    result = handleTest();
    return outputResponse(e, result);
  }

  if (action === "readOutput") {
    result = handleReadOutput(e);
    return outputResponse(e, result);
  }

  if (action === "readStoryline") {
    result = handleReadStoryline();
    return outputResponse(e, result);
  }


  if (action === "clearData") {
    result = handleClearData();
    return outputResponse(e, result);
  }

  if (action === "generateResultsFromTasks" || action === "generateResults") {
    result = handleGenerateResultsFromTasks();
    return outputResponse(e, result);
  }

  if (action === "fetchUrlDocument" || action === "fetchArticle") {
    result = handleFetchUrlDocument((e && e.parameter) || {});
    return outputResponse(e, result);
  }

  result = { ok: false, error: "Unknown action", action: action };
  return outputResponse(e, result);
}

function doPost(e) {
  let body = {};

  if (e && e.parameter && e.parameter.action) {
    body = e.parameter;
  } else if (e && e.postData && e.postData.contents) {
    try {
      body = JSON.parse(e.postData.contents || "{}");
    } catch (err) {
      body = {};
    }
  }

  const action = body.action || "";

  if (action === "writeProjectInfo") {
    return jsonResponse(handleWriteProjectInfo(body));
  }

  if (action === "writeInput") {
    return jsonResponse(handleWriteInput(body));
  }

  if (action === "writeDocumentPages") {
    return jsonResponse(handleWriteDocumentPages(body));
  }

  if (action === "writeDocumentMeta") {
    return jsonResponse(handleWriteDocumentMeta(body));
  }

  if (action === "appendDocumentPages") {
    return jsonResponse(handleAppendDocumentPages(body));
  }

  if (action === "createKnowledgeTasks") {
    return jsonResponse(handleCreateKnowledgeTasks(body));
  }

  if (action === "generateResultsFromTasks" || action === "generateResults") {
    return jsonResponse(handleGenerateResultsFromTasks());
  }

  if (action === "fetchUrlDocument") {
    return jsonResponse(handleFetchUrlDocument(body));
  }

  if (action === "clearData") {
    return jsonResponse(handleClearData());
  }

  return jsonResponse({ ok: false, error: "Unknown action", action: action });
}


function handleReadStoryline() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_STORYLINE);
  if (!sheet) {
    return { ok: false, error: "StoryLine 시트를 찾을 수 없습니다." };
  }
  const values = sheet.getRange("B2:B6").getDisplayValues();
  return {
    ok: true,
    storyline: values[0][0] || "",
    trendInsight: values[1][0] || "",
    coreRisk: values[2][0] || "",
    solutionDirection: values[3][0] || "",
    expectedEffect: values[4][0] || "",
    b2: values[0][0] || "",
    b3: values[1][0] || "",
    b4: values[2][0] || "",
    b5: values[3][0] || "",
    b6: values[4][0] || ""
  };
}

function handleTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCoreSheets(ss);

  const documents = getOrCreateSheet(ss, SHEET_DOCUMENTS);
  const pages = getOrCreateSheet(ss, SHEET_PAGES);

  const testDocId = "TEST_DOC_" + Date.now();
  documents.appendRow([
    testDocId,
    "TEST_COMPANY",
    "connection_test",
    "connection_test.txt",
    new Date(),
    "uploaded"
  ]);

  pages.appendRow([
    testDocId,
    1,
    "Google Sheet 연결 테스트 페이지입니다."
  ]);

  return {
    ok: true,
    document_id: testDocId,
    message: "Connection test saved to Document / Page"
  };
}


function handleWriteProjectInfo(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_PROJECT_INPUT);

  // Project_Input의 기존 구조와 B8 이하 수식은 건드리지 않습니다.
  // 프로젝트 등록값만 B2:B7에 순서대로 입력합니다.
  const values = [
    [body.clientName || ""],
    [body.industry || ""],
    [body.proposalTopic || ""],
    [body.proposalTarget || ""],
    [body.analysisBaseDate || ""],
    [body.analysisPeriodLabel || body.analysisPeriod || ""]
  ];
  sheet.getRange("B2:B7").setValues(values);

  return {
    ok: true,
    clientName: body.clientName || "",
    message: "Project information saved to Project_Input B2:B7"
  };
}

function handleWriteInput(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const input = getOrCreateSheet(ss, SHEET_INPUT);

  ensureHeaders(input, ["session_id", "source_type", "file_name", "raw_text", "created_at", "request_type", "requested_sections", "meta_json"]);

  input.appendRow([
    body.session_id || "",
    body.source_type || "",
    body.file_name || "",
    body.raw_text || "",
    new Date(),
    body.request_type || "",
    body.requested_sections || "[]",
    body.meta || "{}"
  ]);

  return { ok: true, session_id: body.session_id || "", message: "Input saved to Google Sheet" };
}


function parseDocumentMeta_(body) {
  let doc = {};
  let meta = {};
  try { doc = JSON.parse(body.document_json || "{}"); } catch (err) { doc = {}; }
  try { meta = JSON.parse(body.meta_json || "{}"); } catch (err) { meta = {}; }

  const documentId = doc.document_id || body.document_id || ("DOC_" + Date.now());
  return {
    document_id: documentId,
    company: doc.company || body.company || "UNKNOWN_COMPANY",
    document_type: doc.document_type || body.document_type || "business_report",
    title: doc.title || body.file_name || "uploaded_document",
    upload_date: doc.upload_date ? new Date(doc.upload_date) : new Date(),
    status: doc.status || "uploaded",
    meta: meta
  };
}

function handleWriteDocumentMeta(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCoreSheets(ss);
  const documentsSheet = getOrCreateSheet(ss, SHEET_DOCUMENTS);
  const doc = parseDocumentMeta_(body);

  documentsSheet.appendRow([
    doc.document_id,
    doc.company,
    doc.document_type,
    doc.title,
    doc.upload_date,
    doc.status
  ]);

  return {
    ok: true,
    document_id: doc.document_id,
    company: doc.company,
    document_type: doc.document_type,
    message: "Document metadata saved"
  };
}

function handleAppendDocumentPages(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCoreSheets(ss);
  const pagesSheet = getOrCreateSheet(ss, SHEET_PAGES);
  const documentId = body.document_id || "";
  let pages = [];
  try { pages = JSON.parse(body.pages_json || "[]"); } catch (err) { pages = []; }

  if (pages.length) {
    const pageRows = pages.map(function(p) {
      return [
        p.document_id || documentId,
        p.page_no || "",
        trimForCell(p.page_text || "")
      ];
    });
    pagesSheet.getRange(pagesSheet.getLastRow() + 1, 1, pageRows.length, 3).setValues(pageRows);
  }

  return {
    ok: true,
    document_id: documentId,
    pages_saved: pages.length,
    batch_no: body.batch_no || "",
    message: "Page batch saved"
  };
}

function handleCreateKnowledgeTasks(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCoreSheets(ss);
  const doc = {
    document_id: body.document_id || "",
    company: body.company || "UNKNOWN_COMPANY",
    document_type: body.document_type || "business_report",
    title: body.title || body.file_name || "uploaded_document"
  };
  const tasksCreated = createKnowledgeTasksForDocument_(ss, doc);
  return {
    ok: true,
    document_id: doc.document_id,
    document_type: doc.document_type,
    tasks_created: tasksCreated,
    message: "Knowledge tasks created"
  };
}

function handleWriteDocumentPages(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCoreSheets(ss);

  const documentsSheet = getOrCreateSheet(ss, SHEET_DOCUMENTS);
  const pagesSheet = getOrCreateSheet(ss, SHEET_PAGES);

  let doc = {};
  let pages = [];
  let meta = {};

  try {
    doc = JSON.parse(body.document_json || "{}");
  } catch (err) {
    doc = {};
  }

  try {
    pages = JSON.parse(body.pages_json || "[]");
  } catch (err) {
    pages = [];
  }

  try {
    meta = JSON.parse(body.meta_json || "{}");
  } catch (err) {
    meta = {};
  }

  const documentId = doc.document_id || body.document_id || ("DOC_" + Date.now());
  const company = doc.company || body.company || "UNKNOWN_COMPANY";
  const documentType = doc.document_type || body.document_type || "business_report";
  const title = doc.title || body.file_name || "uploaded_document";
  const uploadDate = doc.upload_date ? new Date(doc.upload_date) : new Date();
  const status = doc.status || "uploaded";

  documentsSheet.appendRow([
    documentId,
    company,
    documentType,
    title,
    uploadDate,
    status
  ]);

  if (pages.length) {
    const pageRows = pages.map(function(p) {
      return [
        p.document_id || documentId,
        p.page_no || "",
        trimForCell(p.page_text || "")
      ];
    });
    pagesSheet.getRange(pagesSheet.getLastRow() + 1, 1, pageRows.length, 3).setValues(pageRows);
  }

  const tasksCreated = createKnowledgeTasksForDocument_(ss, {
    document_id: documentId,
    company: company,
    document_type: documentType,
    title: title
  });

  return {
    ok: true,
    document_id: documentId,
    company: company,
    document_type: documentType,
    pages_saved: pages.length,
    tasks_created: tasksCreated,
    meta: meta,
    message: "Document, pages and knowledge tasks saved to Google Sheet"
  };
}


function createKnowledgeTasksForDocument_(ss, doc) {
  const taskDefs = getTaskDefinitionsForDocumentType_(doc.document_type);
  if (!taskDefs.length) return 0;

  const tasksSheet = getOrCreateSheet(ss, SHEET_TASKS);
  ensureHeaders(tasksSheet, ["task_id", "document_id", "agent", "knowledge_domain", "category", "property", "task_key", "task_instruction", "status", "created_at"]);

  const now = new Date();
  const rows = taskDefs.map(function(def, idx) {
    return [
      makeTaskId_(doc.document_id, def.task_key, idx),
      doc.document_id,
      def.agent,
      def.knowledge_domain,
      def.category,
      def.property,
      def.task_key,
      def.task_instruction,
      "READY",
      now
    ];
  });

  if (rows.length) {
    tasksSheet.getRange(tasksSheet.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
  }
  return rows.length;
}

function getTaskDefinitionsForDocumentType_(documentType) {
  const type = String(documentType || "").trim();

  const businessTasks = [
    taskDef_("백장표 컨설턴트", "Business", "Overview", "사업개요", "business_overview", "사업보고서에서 회사의 사업 개요, 사업모델, 핵심 영위 사업을 추출한다. 근거 페이지와 원문 근거를 남기고, 없으면 NOT_FOUND로 둔다."),
    taskDef_("백장표 컨설턴트", "Business", "Product", "주요 제품·서비스", "products_services", "사업보고서에서 주요 제품 및 서비스, 매출 품목, 서비스 구조를 추출한다. 근거 페이지와 원문 근거를 남긴다."),
    taskDef_("백장표 컨설턴트", "Business", "Operation", "생산설비·운영특성", "operation", "사업보고서에서 원재료, 생산설비, 사업장, 공급망, 운영 특성을 추출한다. 제조업이 아니면 해당 산업 특성에 맞춰 운영 구조를 정리한다."),
    taskDef_("백장표 컨설턴트", "Business", "Financial", "매출·영업이익", "financials", "사업보고서에서 최근 3개년 매출과 영업이익을 추출한다. 별도 기준을 우선하고, 없으면 연결 기준임을 명시한다."),
    taskDef_("백장표 컨설턴트", "Business", "Risk", "사업 리스크", "business_risk", "사업보고서에서 회사가 언급한 주요 사업 리스크, 시장 리스크, 재무 리스크, 규제 리스크를 추출한다.")
  ];

  const industryTasks = [
    taskDef_("백장표 컨설턴트", "Industry", "Market", "시장 규모·성장률·세분화", "market_size_segmentation", "산업 리포트에서 전체 시장 규모, CAGR, 성장 세그먼트, 지역·제품·고객별 세분화 정보를 추출한다."),
    taskDef_("백장표 컨설턴트", "Industry", "Competition", "경쟁구도·플레이어", "competitive_landscape", "산업 리포트에서 시장 집중도, 주요 경쟁사, 점유율, M&A, 투자, 사업철수 등 Strategic Move를 추출한다."),
    taskDef_("백장표 컨설턴트", "Industry", "Customer", "고객 세그먼트·수요동인", "customer_demand_drivers", "산업 리포트에서 고객 유형, 구매 기준, Key Drivers, 고객 Pain Point와 Unmet Needs를 추출한다."),
    taskDef_("백장표 컨설턴트", "Industry", "Regulation", "거시환경·규제변화", "macro_regulatory_impact", "산업 리포트에서 정부 정책, 법적 규제, 보조금, 공급망, 금리·환율·원자재 등 외부 변수 영향을 추출한다."),
    taskDef_("백장표 컨설턴트", "Industry", "ValueChain", "밸류체인·파괴적 혁신", "value_chain_disruption", "산업 리포트에서 Value Chain 구조, 단계별 마진풀, 신규 진입자, 기술 변화에 따른 밸류체인 붕괴 가능성을 추출한다."),
    taskDef_("백장표 컨설턴트", "Industry", "Technology", "기술 변화·디지털 전환", "industry_technology_trend", "산업 리포트에서 AI, 디지털, 자동화, 플랫폼, 데이터 등 기술 변화가 산업에 미치는 영향을 추출한다."),
    taskDef_("백장표 컨설턴트", "Industry", "Opportunity", "컨설팅 기회", "industry_consulting_opportunity", "산업 리포트의 외부 환경 변화가 고객사에 줄 수 있는 컨설팅 기회와 제안 가능 영역을 도출한다. 단, 근거가 부족하면 LOW 또는 NOT_FOUND로 표시한다.")
  ];

  const technologyTasks = [
    taskDef_("최백업 매니저", "Technology", "AI", "AI 활용 이슈", "ai_issue", "입력 자료나 뉴스에서 AI 활용 이슈, 생성형 AI, Agent, 자동화 관련 내용을 추출한다."),
    taskDef_("최백업 매니저", "Technology", "DX", "디지털 전환 니즈", "dx_need", "입력 자료나 고객 니즈에서 DX 추진 배경, 데이터 활용, 시스템 고도화 필요성을 추출한다."),
    taskDef_("최백업 매니저", "Technology", "Opportunity", "AI/DX 제안 기회", "ai_dx_opportunity", "AI/DX 관점에서 고객에게 제안 가능한 기회, 솔루션 방향, 적용 가능 업무 영역을 도출한다.")
  ];

  const esgTasks = [
    taskDef_("최백업 매니저", "ESG", "Environment", "환경 이슈", "environment_issue", "지속가능경영보고서나 관련 자료에서 환경, 탄소, 에너지, 폐기물 관련 이슈를 추출한다."),
    taskDef_("최백업 매니저", "ESG", "Safety", "안전보건 이슈", "safety_issue", "지속가능경영보고서, 중대재해 공시, 보도자료에서 안전보건, SHE, 중대재해, 사업장 리스크를 추출한다."),
    taskDef_("최백업 매니저", "ESG", "Governance", "거버넌스·컴플라이언스", "governance_compliance", "ESG 자료에서 지배구조, 윤리, 준법, 컴플라이언스 관련 이슈를 추출한다."),
    taskDef_("최백업 매니저", "ESG", "Opportunity", "SHE/ESG 제안 기회", "esg_she_opportunity", "비재무 리스크와 지속가능경영 관점에서 제안 가능한 컨설팅 기회를 도출한다.")
  ];

  const clientTasks = [
    taskDef_("최백업 매니저", "Client", "Pain", "고객 Pain Point", "client_pain_point", "회의록이나 입력 내용에서 고객이 직접 언급한 Pain Point를 추출한다."),
    taskDef_("최백업 매니저", "Client", "Need", "명시적 요구사항", "client_need", "회의록이나 입력 내용에서 고객의 명시적 요구사항과 기대 산출물을 추출한다."),
    taskDef_("최백업 매니저", "Client", "Issue", "주요 이슈", "client_issue", "회의록이나 입력 내용에서 고객사가 현재 겪는 주요 이슈, 제약사항, 의사결정 기준을 추출한다."),
    taskDef_("최백업 매니저", "Client", "DecisionCriteria", "의사결정 기준", "client_decision_criteria", "회의록에서 의사결정권자, 평가 기준, 우선순위, 일정 제약을 추출한다.")
  ];

  if (type === "business_report") return businessTasks;
  if (type === "industry_report") return industryTasks;
  if (type === "technology_news" || type === "ai_dx_issue") return technologyTasks;
  if (type === "esg_report" || type === "sustainability_report" || type === "dart_disclosure" || type === "major_accident_notice") return esgTasks;
  if (type === "meeting_minutes") return clientTasks;
  return [];
}

function taskDef_(agent, domain, category, property, key, instruction) {
  return {
    agent: agent,
    knowledge_domain: domain,
    category: category,
    property: property,
    task_key: key,
    task_instruction: instruction
  };
}

function makeTaskId_(documentId, taskKey, idx) {
  const base = String(documentId || "DOC").replace(/[^0-9a-zA-Z가-힣_-]+/g, "_").slice(0, 45);
  const key = String(taskKey || "task").replace(/[^0-9a-zA-Z가-힣_-]+/g, "_").slice(0, 35);
  return "TASK_" + base + "_" + (idx + 1) + "_" + key;
}


function handleGenerateResultsFromTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCoreSheets(ss);

  const tasksSheet = getOrCreateSheet(ss, SHEET_TASKS);
  const pagesSheet = getOrCreateSheet(ss, SHEET_PAGES);
  const resultsSheet = getOrCreateSheet(ss, SHEET_RESULTS);
  const templateSheet = getOrCreateSheet(ss, SHEET_TASK_TEMPLATE);

  const tasksData = getSheetObjects_(tasksSheet);
  const pagesData = getSheetObjects_(pagesSheet);
  const resultsData = getSheetObjects_(resultsSheet);
  const templates = loadTaskTemplates_(templateSheet);

  const existingResultTaskIds = {};
  resultsData.rows.forEach(function(row) {
    const taskId = String(row.task_id || "").trim();
    if (taskId) existingResultTaskIds[taskId] = true;
  });

  const pagesByDocument = {};
  pagesData.rows.forEach(function(row) {
    const documentId = String(row.document_id || "").trim();
    if (!documentId) return;
    if (!pagesByDocument[documentId]) pagesByDocument[documentId] = [];
    pagesByDocument[documentId].push({
      page_no: row.page_no,
      page_text: String(row.page_text || "")
    });
  });

  const now = new Date();
  const resultRows = [];
  const taskStatusUpdates = [];

  tasksData.rows.forEach(function(task, idx) {
    const taskId = String(task.task_id || "").trim();
    if (!taskId || existingResultTaskIds[taskId]) return;

    const status = String(task.status || "").toUpperCase();
    if (status && status !== "READY" && status !== "PENDING") return;

    const documentId = String(task.document_id || "").trim();
    const taskKey = String(task.task_key || task.task_type || task.property || "").trim();
    const pages = pagesByDocument[documentId] || [];
    const context = selectContextForTask_(pages, taskKey, task);
    const template = templates[taskKey] || buildDefaultTemplateFromTask_(task);
    const resultFormula = buildResultAiFormula_(template, task, context.text);

    resultRows.push([
      taskId,
      taskKey || String(task.property || "result"),
      resultFormula,
      "AI 결과 내 EVIDENCE 항목 참조",
      context.source_pages || "NOT_FOUND",
      "PENDING",
      "FORMULA_CREATED",
      now
    ]);

    taskStatusUpdates.push({ rowIndex: idx + 2, status: "RESULT_FORMULA_CREATED" });
  });

  if (resultRows.length) {
    resultsSheet.getRange(resultsSheet.getLastRow() + 1, 1, resultRows.length, 8).setValues(resultRows);
  }

  const statusCol = tasksData.headerMap.status;
  if (statusCol) {
    taskStatusUpdates.forEach(function(item) {
      tasksSheet.getRange(item.rowIndex, statusCol).setValue(item.status);
    });
  }

  return {
    ok: true,
    results_created: resultRows.length,
    message: "Result AI formulas created from READY tasks"
  };
}

function getSheetObjects_(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (!values.length) return { headers: [], headerMap: {}, rows: [] };

  const headers = values[0].map(function(h) { return String(h || "").trim(); });
  const headerMap = {};
  headers.forEach(function(h, idx) {
    if (h) headerMap[h] = idx + 1;
  });

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const obj = {};
    headers.forEach(function(h, c) {
      if (h) obj[h] = values[r][c];
    });
    rows.push(obj);
  }

  return { headers: headers, headerMap: headerMap, rows: rows };
}

function loadTaskTemplates_(sheet) {
  const data = getSheetObjects_(sheet);
  const templates = {};

  data.rows.forEach(function(row) {
    const active = String(row.is_active || "TRUE").toUpperCase();
    if (active === "FALSE" || active === "N") return;
    const type = String(row.task_type || "").trim();
    if (!type) return;
    templates[type] = {
      task_type: type,
      knowledge_domain: row.knowledge_domain || "",
      category: row.category || "",
      property: row.property || "",
      prompt: row.prompt || "",
      output_format: row.output_format || "TEXT"
    };
  });

  return templates;
}

function buildDefaultTemplateFromTask_(task) {
  return {
    task_type: task.task_key || task.task_type || "generic_task",
    knowledge_domain: task.knowledge_domain || "",
    category: task.category || "",
    property: task.property || "",
    prompt: task.task_instruction || "문서 텍스트를 기반으로 핵심 내용을 정리하라. 근거가 없으면 NOT_FOUND라고 작성하라.",
    output_format: "TEXT"
  };
}

function selectContextForTask_(pages, taskKey, task) {
  const key = String(taskKey || "").toLowerCase();
  const category = String((task && task.category) || "").toLowerCase();
  const prop = String((task && task.property) || "").toLowerCase();
  const hayKey = key + " " + category + " " + prop;
  const keywords = getKeywordsForTask_(hayKey);
  const candidates = [];

  pages.forEach(function(p) {
    const text = String(p.page_text || "");
    const score = scorePageForKeywords_(text, keywords);
    if (score > 0) {
      candidates.push({ page_no: p.page_no, page_text: text, score: score });
    }
  });

  candidates.sort(function(a, b) {
    if (b.score !== a.score) return b.score - a.score;
    return Number(a.page_no || 0) - Number(b.page_no || 0);
  });

  let selected = candidates.slice(0, 5);
  if (!selected.length) selected = pages.slice(0, 4);

  let text = selected.map(function(p) {
    return "[p." + p.page_no + "]\n" + String(p.page_text || "").slice(0, 3500);
  }).join("\n\n---\n\n");

  if (!text) text = "NOT_FOUND";
  if (text.length > 18000) text = text.slice(0, 18000) + "\n...[TRUNCATED_FOR_AI_FORMULA]";

  const sourcePages = selected.map(function(p) { return "p." + p.page_no; }).join(", ") || "NOT_FOUND";
  return { text: text, source_pages: sourcePages };
}

function getKeywordsForTask_(key) {
  const common = [];
  if (key.indexOf("overview") >= 0 || key.indexOf("사업개요") >= 0 || key.indexOf("overview") >= 0) {
    return common.concat(["사업의 개요", "사업 개요", "사업개요", "사업의 내용", "영위", "사업부문", "비즈니스", "Business"]);
  }
  if (key.indexOf("product") >= 0 || key.indexOf("제품") >= 0 || key.indexOf("service") >= 0) {
    return common.concat(["주요 제품", "제품 및 서비스", "서비스", "제품", "매출 품목", "상품", "용역"]);
  }
  if (key.indexOf("operation") >= 0 || key.indexOf("manufacturing") >= 0 || key.indexOf("생산") >= 0 || key.indexOf("운영") >= 0) {
    return common.concat(["원재료", "생산설비", "생산능력", "생산실적", "사업장", "공장", "시설", "공급망", "조달"]);
  }
  if (key.indexOf("financial") >= 0 || key.indexOf("매출") >= 0 || key.indexOf("영업이익") >= 0 || key.indexOf("financial") >= 0) {
    return common.concat(["포괄손익계산서", "손익계산서", "매출액", "매출", "영업이익", "별도", "연결", "수익"]);
  }
  if (key.indexOf("risk") >= 0 || key.indexOf("리스크") >= 0 || key.indexOf("위험") >= 0) {
    return common.concat(["위험", "리스크", "위험관리", "시장위험", "신용위험", "유동성위험", "규제", "소송", "우발"]);
  }
  if (key.indexOf("market") >= 0 || key.indexOf("시장") >= 0) {
    return common.concat(["시장", "시장규모", "성장률", "CAGR", "전망", "수요", "성장", "segment", "market"]);
  }
  if (key.indexOf("competition") >= 0 || key.indexOf("경쟁") >= 0) {
    return common.concat(["경쟁", "경쟁사", "점유율", "시장점유율", "player", "M&A", "투자", "선도", "기업"]);
  }
  if (key.indexOf("customer") >= 0 || key.indexOf("고객") >= 0 || key.indexOf("demand") >= 0) {
    return common.concat(["고객", "수요", "구매", "Pain", "니즈", "Needs", "수요처", "고객사", "동인"]);
  }
  if (key.indexOf("regulation") >= 0 || key.indexOf("규제") >= 0 || key.indexOf("compliance") >= 0) {
    return common.concat(["규제", "정책", "정부", "법", "인허가", "컴플라이언스", "FDA", "중대재해", "공시"]);
  }
  if (key.indexOf("value") >= 0 || key.indexOf("chain") >= 0 || key.indexOf("밸류") >= 0) {
    return common.concat(["밸류체인", "Value Chain", "공급망", "마진", "원가", "upstream", "downstream", "disruption"]);
  }
  if (key.indexOf("technology") >= 0 || key.indexOf("ai") >= 0 || key.indexOf("dx") >= 0 || key.indexOf("기술") >= 0) {
    return common.concat(["AI", "인공지능", "디지털", "DX", "자동화", "데이터", "플랫폼", "Agent", "LLM", "클라우드"]);
  }
  if (key.indexOf("environment") >= 0 || key.indexOf("환경") >= 0) {
    return common.concat(["환경", "탄소", "온실가스", "에너지", "폐기물", "수질", "대기"]);
  }
  if (key.indexOf("safety") >= 0 || key.indexOf("안전") >= 0 || key.indexOf("she") >= 0) {
    return common.concat(["안전", "보건", "SHE", "중대재해", "재해", "사고", "위험성평가", "산업안전"]);
  }
  if (key.indexOf("governance") >= 0 || key.indexOf("거버넌스") >= 0) {
    return common.concat(["거버넌스", "지배구조", "윤리", "준법", "컴플라이언스", "이사회"]);
  }
  if (key.indexOf("client") >= 0 || key.indexOf("pain") >= 0 || key.indexOf("need") >= 0 || key.indexOf("meeting") >= 0 || key.indexOf("회의") >= 0) {
    return common.concat(["요구", "니즈", "Pain", "이슈", "회의", "요청", "과제", "문제", "제약", "의사결정"]);
  }
  return common.concat(["사업", "시장", "고객", "리스크", "기회", "제안"]);
}

function scorePageForKeywords_(text, keywords) {
  const t = String(text || "").toLowerCase();
  let score = 0;
  keywords.forEach(function(k) {
    const kw = String(k || "").toLowerCase();
    if (!kw) return;
    const idx = t.indexOf(kw);
    if (idx >= 0) score += kw.length > 3 ? 3 : 1;
  });
  return score;
}

function buildResultAiFormula_(template, task, contextText) {
  const taskInfo = [
    "너는 컨설팅펌 리서치 어시스턴트다.",
    "아래 문서 근거만 사용해서 결과를 작성하라.",
    "추정하지 말고, 근거가 없으면 NOT_FOUND라고 작성하라.",
    "반드시 한국어로 작성하라.",
    "",
    "[Knowledge Domain] " + (template.knowledge_domain || task.knowledge_domain || ""),
    "[Category] " + (template.category || task.category || ""),
    "[Property] " + (template.property || task.property || ""),
    "",
    "[Task Instruction]",
    String(template.prompt || task.task_instruction || "문서 텍스트를 기반으로 핵심 내용을 정리하라."),
    "",
    "[Output Format]",
    "VALUE: 핵심 정리 1~5개 bullet 또는 3줄 요약",
    "EVIDENCE: 원문 근거 1~2문장",
    "SOURCE_PAGE: p.숫자",
    "CONFIDENCE: HIGH / MEDIUM / LOW / NOT_FOUND",
    "",
    "[Document Text]",
    contextText || "NOT_FOUND"
  ].join("\n");

  let body = taskInfo;
  if (body.length > 24000) body = body.slice(0, 24000) + "\n...[TRUNCATED_FOR_FORMULA]";
  return '=AI("' + escapeForFormulaString_(body) + '")';
}

function escapeForFormulaString_(text) {
  return String(text || "").replace(/"/g, '""');
}


function handleFetchUrlDocument(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCoreSheets(ss);

  const url = String(body.url || "").trim();
  if (!/^https?:\/\//i.test(url)) {
    return { ok: false, error: "Invalid URL", url: url };
  }

  const company = body.company || "UNKNOWN_COMPANY";
  const documentType = body.document_type || "technology_news";
  const fetchedAt = new Date();
  let html = "";
  let statusCode = "";
  let title = url;
  let text = "";

  try {
    const res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true,
      headers: {
        "User-Agent": "Mozilla/5.0 ProposalAgentRPG/1.0"
      }
    });
    statusCode = res.getResponseCode();
    html = res.getContentText("UTF-8");
    title = extractTitle_(html) || url;
    text = stripHtmlForKnowledge_(html);

    if (!text || text.length < 80) {
      text = "Source URL: " + url + "\n" +
        "본문을 충분히 추출하지 못했습니다. 언론사 차단/로그인/동적 렌더링 가능성이 있습니다. 기사 본문을 직접 붙여넣어 주세요.";
    } else {
      text = "Source URL: " + url + "\nHTTP Status: " + statusCode + "\nTitle: " + title + "\n\n" + text;
    }
  } catch (err) {
    text = "Source URL: " + url + "\nURL 본문 수집 실패: " + (err.message || err) + "\n기사 본문을 직접 붙여넣어 주세요.";
  }

  const safeId = sanitizeForId_(title || url).slice(0, 45) || "url";
  const documentId = "DOC_URL_" + Date.now() + "_" + safeId;
  const documentsSheet = getOrCreateSheet(ss, SHEET_DOCUMENTS);
  const pagesSheet = getOrCreateSheet(ss, SHEET_PAGES);

  documentsSheet.appendRow([
    documentId,
    company,
    documentType,
    title,
    fetchedAt,
    "uploaded"
  ]);

  pagesSheet.appendRow([
    documentId,
    1,
    trimForCell(text)
  ]);

  const tasksCreated = createKnowledgeTasksForDocument_(ss, {
    document_id: documentId,
    company: company,
    document_type: documentType,
    title: title
  });

  return {
    ok: true,
    document_id: documentId,
    company: company,
    document_type: documentType,
    url: url,
    title: title,
    http_status: statusCode,
    tasks_created: tasksCreated,
    message: "URL document fetched and saved"
  };
}

function extractTitle_(html) {
  const m = String(html || "").match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  if (!m) return "";
  return decodeHtml_(m[1]).replace(/\s+/g, " ").trim();
}

function stripHtmlForKnowledge_(html) {
  let text = String(html || "");
  text = text.replace(/<script[\s\S]*?<\/script>/gi, " ");
  text = text.replace(/<style[\s\S]*?<\/style>/gi, " ");
  text = text.replace(/<noscript[\s\S]*?<\/noscript>/gi, " ");
  text = text.replace(/<header[\s\S]*?<\/header>/gi, " ");
  text = text.replace(/<footer[\s\S]*?<\/footer>/gi, " ");
  text = text.replace(/<nav[\s\S]*?<\/nav>/gi, " ");
  text = text.replace(/<[^>]+>/g, " ");
  text = decodeHtml_(text);
  text = text.replace(/\s+/g, " ").trim();
  return text.length > 45000 ? text.slice(0, 45000) + "\n...[TRUNCATED_FOR_CELL_LIMIT]" : text;
}

function decodeHtml_(text) {
  return String(text || "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
}

function sanitizeForId_(text) {
  return String(text || "")
    .replace(/[^0-9a-zA-Z가-힣_-]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "");
}

function handleClearData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const cleared = [];

  // Project_Input: 입력값 B2:B7만 삭제하고 B8 이하 수식/참조 영역은 보존합니다.
  const projectInput = ss.getSheetByName(SHEET_PROJECT_INPUT);
  if (!projectInput) {
    throw new Error("Project_Input 시트를 찾을 수 없습니다.");
  }
  projectInput.getRange("B2:B7").clearContent();
  cleared.push(SHEET_PROJECT_INPUT + "!B2:B7");

  // Document / Page: 1행 헤더는 유지하고 2행 이하의 값만 삭제합니다.
  [SHEET_DOCUMENTS, SHEET_PAGES].forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      throw new Error(name + " 시트를 찾을 수 없습니다.");
    }
    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    }
    cleared.push(name + "!2:" + Math.max(lastRow, 2));
  });

  // Research Board, Storyline, Partner PPT, Review 및 그 밖의 시트는 절대 수정하지 않습니다.
  return {
    ok: true,
    message: "RESET complete: Project_Input B2:B7, Document rows 2+, Page rows 2+ cleared only",
    cleared_sheets: cleared,
    preserved_sheets: ["Research Board", "Storyline", "Partner PPT", "Review"]
  };
}

function handleReadOutput(e) {
  const sessionId = (e && e.parameter && e.parameter.session_id) || "";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const output = getOrCreateSheet(ss, SHEET_OUTPUT);

  ensureHeaders(output, ["session_id", "result_type", "result_json", "updated_at"]);

  const values = output.getDataRange().getValues();

  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]) === String(sessionId)) {
      return {
        ok: true,
        session_id: sessionId,
        result_type: values[i][1],
        result_json: values[i][2],
        updated_at: values[i][3]
      };
    }
  }

  return { ok: false, session_id: sessionId, error: "No output found" };
}

function ensureCoreSheets(ss) {
  ensureHeaders(getOrCreateSheet(ss, SHEET_DOCUMENTS), ["document_id", "company", "document_type", "title", "upload_date", "status"]);
  ensureHeaders(getOrCreateSheet(ss, SHEET_PAGES), ["document_id", "page_no", "page_text"]);
  ensureHeaders(getOrCreateSheet(ss, SHEET_TASKS), ["task_id", "document_id", "agent", "knowledge_domain", "category", "property", "task_key", "task_instruction", "status", "created_at"]);
  ensureHeaders(getOrCreateSheet(ss, SHEET_RESULTS), ["task_id", "result_type", "result_text", "evidence", "source_page", "confidence", "status", "created_at"]);
  ensureHeaders(getOrCreateSheet(ss, SHEET_KNOWLEDGE), ["company", "knowledge_domain", "category", "property", "value", "evidence", "source_document", "source_page", "confidence", "priority", "created_at"]);
  ensureHeaders(getOrCreateSheet(ss, SHEET_TASK_TEMPLATE), ["task_type", "knowledge_domain", "category", "property", "prompt", "output_format", "is_active"]);
}
function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function ensureHeaders(sheet, headers) {
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const empty = current.every(function(v) { return v === ""; });

  if (empty) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  headers.forEach(function(h, idx) {
    const cellValue = String(current[idx] || "");
    if (!cellValue) sheet.getRange(1, idx + 1).setValue(h);
  });
}

function trimForCell(value) {
  const text = String(value || "");
  // Google Sheets single-cell limit is 50,000 chars. Keep margin for safety.
  return text.length > 45000 ? text.slice(0, 45000) + "\n...[TRUNCATED_FOR_CELL_LIMIT]" : text;
}

function outputResponse(e, obj) {
  const callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    const safeCallback = String(callback).replace(/[^a-zA-Z0-9_.$]/g, "");
    return ContentService
      .createTextOutput(safeCallback + "(" + JSON.stringify(obj) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonResponse(obj);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
