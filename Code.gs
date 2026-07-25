const SETTINGS_SHEET = "Settings";
const SUBMISSIONS_SHEET = "Submissions";

const DEFAULT_RECORD_SCHEMA = {
  title: "챗봇 흔적 분석",
  help: "AI와 주고받은 내용과 이를 검토하며 생각을 발전시킨 과정을 기록합니다.",
  stepName: "탐구 과정",
  toolUrl: "https://gemini.google.com/",
  fields: [
    { key: "q", enabled: true, label: "🙋 AI에게 한 질문·요청", placeholder: "예) 이 현상이 일어나는 까닭을 근거와 함께 설명해 줘." },
    { key: "a", enabled: true, label: "🤖 AI 답변·결과", placeholder: "AI가 답한 내용을 붙여넣거나 핵심만 요약해 적어요." },
    { key: "r", enabled: true, label: "🔎 검토·성찰", placeholder: "확인한 근거, 발견한 오류, 수정한 생각과 더 궁금한 점을 적어요." }
  ]
};

function json_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function settings_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  if (!sheet) throw new Error("Settings 시트를 찾을 수 없습니다.");

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const rows = sheet.getRange(1, 1, lastRow, 6).getDisplayValues();
  const classCodes = rows.slice(1).map(r => r[0].trim()).filter(Boolean);
  const topics = rows.slice(1).map(r => r[1].trim()).filter(Boolean);
  const legacyToolUrl = (rows[0][2] || "").trim();
  const values = {};

  rows.slice(1).forEach(r => {
    const label = (r[3] || "").trim();
    if (label) values[label] = (r[4] || "").trim();
  });

  const truthy = value => !["아니오", "미사용", "false", "0", "off"].includes(String(value || "").trim().toLowerCase());
  const pick = (label, fallback) => values[label] || fallback;
  const toolUrl = pick("AI 도구 주소", legacyToolUrl || DEFAULT_RECORD_SCHEMA.toolUrl);
  const fields = DEFAULT_RECORD_SCHEMA.fields.map((field, i) => {
    const n = i + 1;
    return {
      key: field.key,
      enabled: truthy(pick(`수집 항목 ${n} 사용`, "예")),
      label: pick(`수집 항목 ${n} 이름`, field.label),
      placeholder: pick(`수집 항목 ${n} 도움말`, field.placeholder)
    };
  });
  if (!fields.some(field => field.enabled)) fields[0].enabled = true;

  return {
    teacherPassword: pick("교사 비밀번호", "mirae1234"),
    publicConfig: {
      classCodes,
      topics,
      chatbotUrl: toolUrl,
      recordSchema: {
        title: pick("기록 방식 이름", DEFAULT_RECORD_SCHEMA.title),
        help: pick("학생 안내 문구", DEFAULT_RECORD_SCHEMA.help),
        stepName: pick("반복 단위 이름", DEFAULT_RECORD_SCHEMA.stepName),
        toolUrl,
        fields
      }
    }
  };
}

function submissions_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
  if (!sheet) sheet = ss.insertSheet(SUBMISSIONS_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "ClassCode", "StudentId", "Name", "Topic", "QA_Data_JSON", "TeacherFeedback"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function records_() {
  const sheet = submissions_();
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues().map((r, i) => ({
    row: i + 2,
    timestamp: r[0] instanceof Date ? r[0].toISOString() : r[0],
    classCode: r[1],
    studentId: r[2],
    name: r[3],
    topic: r[4],
    qa: parseQa_(r[5]),
    feedback: r[6] || ""
  })).reverse();
}

function parseQa_(value) {
  try {
    const parsed = JSON.parse(value || "[]");
    return Array.isArray(parsed) ? parsed : [];
  } catch (_) {
    return [];
  }
}

function validPassword_(candidate) {
  const expected = settings_().teacherPassword;
  return String(candidate || "") === String(expected || "");
}

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.action === "config") return json_(settings_().publicConfig);
    if (p.action === "portfolio") {
      const records = records_().filter(r =>
        String(r.classCode) === String(p.classCode || "") &&
        String(r.studentId) === String(p.studentId || "") &&
        String(r.name) === String(p.name || "")
      );
      return json_({ records });
    }
    if (p.action === "all") {
      if (!validPassword_(p.pw)) return json_({ error: "비밀번호가 맞지 않습니다." });
      return json_({ records: records_() });
    }
    return json_({ error: "지원하지 않는 요청입니다." });
  } catch (err) {
    return json_({ error: err.message || String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (body.action === "feedback") {
      if (!validPassword_(body.pw)) return json_({ error: "비밀번호가 맞지 않습니다." });
      const sheet = submissions_();
      const row = Number(body.row);
      if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) {
        return json_({ error: "저장할 기록을 찾을 수 없습니다." });
      }
      sheet.getRange(row, 7).setValue(body.feedback || "");
      return json_({ ok: true });
    }

    const required = ["classCode", "studentId", "name", "topic"];
    if (required.some(key => !String(body[key] || "").trim())) {
      return json_({ error: "필수 입력값을 확인해 주세요." });
    }
    const qa = Array.isArray(body.qa) ? body.qa : [];
    submissions_().appendRow([
      new Date(),
      body.classCode,
      body.studentId,
      body.name,
      body.topic,
      JSON.stringify(qa),
      ""
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ error: err.message || String(err) });
  }
}
