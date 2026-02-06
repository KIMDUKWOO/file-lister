import "./style.css";
import * as XLSX from "xlsx";

const state = {
  srcHandle: null,
  shouldCancel: false,
};

const $ = (sel) => document.querySelector(sel);

function ui() {
  const supports =
    "showDirectoryPicker" in window && "showSaveFilePicker" in window;

  document.querySelector("#app").innerHTML = `
    <div class="container">
      <div class="header">
        <div class="h1">Folder Tree → Excel(.xlsx) Exporter</div>
        <div class="sub">
          Chrome / Edge 전용 · File System Access API 사용
        </div>
      </div>

      <div class="card">
        <div class="row" style="justify-content: space-between;">
          <div class="pill">API: <span class="badge" id="apiBadge"></span></div>
          <div class="pill">폴더: <span class="badge" id="folderBadge">(없음)</span></div>
        </div>

        <div class="hr"></div>

        <div class="row">
          <button class="btn" id="pickBtn">폴더 선택</button>
          <button class="btn primary" id="saveBtn" disabled>Excel로 저장</button>
          <button class="btn danger" id="cancelBtn" disabled>중지</button>

          <label class="pill" style="margin-left:auto;">
            <input type="checkbox" id="recursiveChk" checked />
            하위 폴더 포함
          </label>
        </div>

        <div class="hr"></div>

        <pre id="log"></pre>
      </div>
    </div>
  `;

  $("#apiBadge").textContent = supports ? "지원" : "미지원";
  $("#apiBadge").style.color = supports ? "#7ee7a3" : "#ffb4b4";

  $("#pickBtn").onclick = onPickFolder;
  $("#saveBtn").onclick = onSaveExcel;
  $("#cancelBtn").onclick = () => {
    state.shouldCancel = true;
    log("중지 요청됨");
  };
}

function log(msg) {
  const el = $("#log");
  el.textContent += `[${new Date().toLocaleTimeString()}] ${msg}\n`;
  el.scrollTop = el.scrollHeight;
}

function setBusy(busy) {
  $("#pickBtn").disabled = busy;
  $("#saveBtn").disabled = busy || !state.srcHandle;
  $("#cancelBtn").disabled = !busy;
  $("#recursiveChk").disabled = busy;
}

async function onPickFolder() {
  try {
    state.srcHandle = await window.showDirectoryPicker({ mode: "read" });
    $("#folderBadge").textContent = state.srcHandle.name;
    $("#saveBtn").disabled = false;
    log(`폴더 선택됨: ${state.srcHandle.name}`);
  } catch (e) {
    log("폴더 선택 취소");
  }
}

/**
 * ✅ 핵심 수정 포인트
 * - showSaveFilePicker 를 클릭 이벤트 중 '가장 먼저' 호출
 * - 그 다음에 스캔/엑셀 생성
 */
async function onSaveExcel() {
  if (!state.srcHandle) {
    alert("먼저 폴더를 선택하세요.");
    return;
  }

  let fileHandle;
  try {
    fileHandle = await window.showSaveFilePicker({
      suggestedName: `${state.srcHandle.name}_tree.xlsx`,
      types: [
        {
          description: "Excel",
          accept: {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [
              ".xlsx",
            ],
          },
        },
      ],
    });
  } catch {
    log("저장 취소");
    return;
  }

  setBusy(true);
  state.shouldCancel = false;

  try {
    const recursive = $("#recursiveChk").checked;
    log("스캔 시작");

    const { rows, maxDepth } = await scanToRows(
      state.srcHandle,
      recursive
    );

    const header = Array.from(
      { length: Math.max(1, maxDepth) },
      (_, i) => `Level ${i + 1}`
    );

    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    applyHyperlinks(ws, rows);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Tree");

    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const writable = await fileHandle.createWritable();
    await writable.write(out);
    await writable.close();

    alert("저장 완료");
    log("저장 완료");
  } catch (err) {
    alert("오류 발생");
    console.error(err);
  } finally {
    setBusy(false);
  }
}

async function scanToRows(root, recursive) {
  const rows = [];
  let maxDepth = 1;

  async function walk(dir, parts) {
    for await (const entry of dir.values()) {
      if (state.shouldCancel) return;

      const next = [...parts, entry.name];
      maxDepth = Math.max(maxDepth, next.length);

      if (entry.kind === "directory") {
        rows.push(withMeta(next, "dir"));
        if (recursive) await walk(entry, next);
      } else {
        rows.push(withMeta(next, "file"));
      }
    }
  }

  await walk(root, []);
  return { rows, maxDepth };
}

function withMeta(parts, type) {
  const row = [...parts];
  row.__meta = {
    type,
    path: parts.join("/"),
    depth: parts.length,
  };
  return row;
}

function applyHyperlinks(ws, rows) {
  rows.forEach((r, i) => {
    if (r.__meta.type !== "file") return;

    const cell = XLSX.utils.encode_cell({
      r: i + 1,
      c: r.__meta.depth - 1,
    });

    ws[cell].l = { Target: r.__meta.path };
    ws[cell].s = {
      font: { color: { rgb: "0000FF" }, underline: true },
    };
  });
}

ui();
