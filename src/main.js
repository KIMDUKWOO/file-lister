import "./style.css";
import * as XLSX from "xlsx";

const state = {
  srcHandle: null,
  shouldCancel: false,
};

const $ = (sel) => document.querySelector(sel);

function ui() {
  const supports = "showDirectoryPicker" in window && "showSaveFilePicker" in window;

  document.querySelector("#app").innerHTML = `
    <div class="container">
      <div class="header">
        <div class="h1">Folder Tree → Excel(.xlsx) Exporter</div>
        <div class="sub">
          Chrome/Edge(Chromium) 권장 · File System Access API 사용 · 폴더 트리를 스캔해 Excel로 저장합니다.
        </div>
      </div>

      <div class="card">
        <div class="row" style="justify-content: space-between;">
          <div class="pill">API 지원 여부: <span class="badge" id="apiBadge"></span></div>
          <div class="pill">선택된 폴더: <span class="badge" id="folderBadge">(없음)</span></div>
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

  if (!supports) {
    log("이 브라우저는 File System Access API를 지원하지 않습니다. Chrome/Edge에서 실행하세요.");
  }

  $("#pickBtn").addEventListener("click", onPickFolder);
  $("#saveBtn").addEventListener("click", onSaveExcel);
  $("#cancelBtn").addEventListener("click", () => {
    state.shouldCancel = true;
    log("중지 요청됨… 현재 작업이 끝나는 지점에서 중단합니다.");
  });
}

function log(msg) {
  const el = $("#log");
  const t = new Date().toLocaleTimeString();
  el.textContent += `[${t}] ${msg}\n`;
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
    state.shouldCancel = false;

    $("#folderBadge").textContent = state.srcHandle?.name || "(없음)";
    $("#saveBtn").disabled = !state.srcHandle;

    log(`폴더 선택됨: ${state.srcHandle.name}`);

    // 권한 확인/요청(대용량에서도 중간 실패 방지)
    const ok = await ensurePermission(state.srcHandle, "read");
    if (!ok) {
      state.srcHandle = null;
      $("#folderBadge").textContent = "(없음)";
      $("#saveBtn").disabled = true;
      alert("폴더 읽기 권한이 필요합니다.");
      return;
    }

    log("폴더 읽기 권한 확인 완료");
  } catch (err) {
    // 사용자가 취소한 경우도 여기로 들어올 수 있음
    log(`폴더 선택 실패: ${safeErr(err)}`);
  }
}

async function onSaveExcel() {
  if (!state.srcHandle) {
    alert("먼저 폴더를 선택하세요.");
    return;
  }

  setBusy(true);
  state.shouldCancel = false;

  try {
    const recursive = $("#recursiveChk").checked;

    log(`스캔 시작 (하위 폴더 포함: ${recursive ? "ON" : "OFF"})`);

    const { rows, maxDepth, fileCount, dirCount } = await scanToRows(state.srcHandle, recursive);

    if (state.shouldCancel) {
      log("작업이 중지되었습니다.");
      alert("작업이 중지되었습니다.");
      return;
    }

    log(`스캔 완료: 폴더 ${dirCount}개, 파일 ${fileCount}개, 행 ${rows.length}개`);

    const defaultName = `${state.srcHandle.name}_tree.xlsx`;

    const fileHandle = await window.showSaveFilePicker({
      suggestedName: defaultName,
      types: [
        {
          description: "Excel Workbook",
          accept: {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
          },
        },
      ],
    });

    // 워크북 구성
    const wb = XLSX.utils.book_new();

    // 헤더(깊이에 따라 동적 생성)
    const header = Array.from({ length: Math.max(1, maxDepth) }, (_, i) => `Level ${i + 1}`);

    const aoa = [header, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // 열 너비: 길이 기반으로 대략 계산
    const colWidths = header.map((_, c) => {
      let maxLen = header[c]?.length || 8;
      for (let r = 1; r < aoa.length; r++) {
        const v = aoa[r]?.[c];
        if (typeof v === "string") maxLen = Math.max(maxLen, v.length);
      }
      return { wch: Math.min(Math.max(maxLen + 2, 10), 60) };
    });
    ws["!cols"] = colWidths;

    // 파일 하이퍼링크 + 파란색 스타일
    // rows는 {parts, type, relPath, depth} 정보를 가진 내부 구조를 AOA로 바꾼 결과라,
    // 하이퍼링크는 다시 한 번 스캔 결과를 기준으로 설정합니다.
    // (AOA 행 index는 1부터 데이터 시작)
    applyHyperlinks(ws, rows);

    XLSX.utils.book_append_sheet(wb, ws, "Tree");

    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });

    const writable = await fileHandle.createWritable();
    await writable.write(out);
    await writable.close();

    log("저장 완료");
    alert("저장 완료");
  } catch (err) {
    log(`오류: ${safeErr(err)}`);
    alert(`오류가 발생했습니다.\n\n${safeErr(err)}`);
  } finally {
    setBusy(false);
    state.shouldCancel = false;
  }
}

function safeErr(err) {
  if (!err) return "(unknown error)";
  if (typeof err === "string") return err;
  return err?.message || JSON.stringify(err);
}

async function ensurePermission(handle, mode = "read") {
  try {
    // 이미 권한이 있으면 바로 OK
    const q = await handle.queryPermission({ mode });
    if (q === "granted") return true;

    const r = await handle.requestPermission({ mode });
    return r === "granted";
  } catch {
    return false;
  }
}

// 폴더 트리를 rows(AOA row)로 변환
// - 폴더는 마지막 레벨에 폴더명
// - 파일은 마지막 레벨에 파일명(하이퍼링크)
// - 루트 폴더 이름은 표시하지 않고, 루트 기준 상대 경로로만 기록
async function scanToRows(rootHandle, recursive) {
  const rows = [];
  let maxDepth = 1;
  let fileCount = 0;
  let dirCount = 0;

  // BFS로 돌면 깊이가 섞이지만, 여기서는 트리 정렬 느낌을 위해 DFS에 가까운 순서
  async function walkDir(dirHandle, parts) {
    if (state.shouldCancel) return;

    let entries = [];
    try {
      for await (const entry of dirHandle.values()) {
        entries.push(entry);
      }
    } catch (err) {
      // 권한/읽기 에러가 나도 전체가 죽지 않도록
      log(`경고: 디렉토리 읽기 실패 (${parts.join("/") || dirHandle.name}): ${safeErr(err)}`);
      return;
    }

    // 폴더 먼저, 그 다음 파일(보기 좋게)
    entries.sort((a, b) => {
      const ta = a.kind === "directory" ? 0 : 1;
      const tb = b.kind === "directory" ? 0 : 1;
      if (ta !== tb) return ta - tb;
      return (a.name || "").localeCompare(b.name || "", "ko");
    });

    for (const entry of entries) {
      if (state.shouldCancel) return;

      try {
        if (entry.kind === "directory") {
          dirCount++;
          const newParts = [...parts, entry.name];

          // 폴더 자체도 한 줄 기록
          maxDepth = Math.max(maxDepth, newParts.length);
          rows.push(toRow(newParts, "dir"));

          if (recursive) {
            await walkDir(entry, newParts);
          }
        } else if (entry.kind === "file") {
          fileCount++;
          const newParts = [...parts, entry.name];
          maxDepth = Math.max(maxDepth, newParts.length);
          rows.push(toRow(newParts, "file"));
        }
      } catch (err) {
        log(`경고: 엔트리 처리 실패 (${entry?.name || "(unknown)"}): ${safeErr(err)}`);
      }

      // 대용량 폴더에서 UI 프리징 방지(주기적으로 이벤트 루프에 양보)
      if ((rows.length + fileCount + dirCount) % 250 === 0) {
        await new Promise((r) => setTimeout(r, 0));
        log(`진행 중… rows=${rows.length}, folders=${dirCount}, files=${fileCount}`);
      }
    }
  }

  // 루트는 표시하지 않고, 루트 하위부터 시작
  await walkDir(rootHandle, []);

  return { rows, maxDepth, fileCount, dirCount };
}

function toRow(parts, type) {
  // AOA row는 최대 컬럼 수가 달라도 XLSX가 자동으로 맞춰줌
  // 파일은 parts 마지막(파일명) 셀에 하이퍼링크를 걸 예정
  const row = parts.slice();

  // 파일이면 Excel 링크 대상이 되는 상대 경로도 저장해 둔다
  // (applyHyperlinks에서 사용)
  row.__meta = {
    type,
    relPath: parts.join("/"),
    depth: parts.length,
  };

  return row;
}

function applyHyperlinks(ws, rows) {
  // ws는 헤더가 1행, 데이터는 2행부터
  // rows[i]의 __meta를 이용해 파일 셀에 링크를 적용
  for (let i = 0; i < rows.length; i++) {
    const meta = rows[i].__meta;
    if (!meta || meta.type !== "file") continue;

    const excelRow = i + 2; // 1행은 header
    const excelCol = meta.depth; // depth(1-based)

    const addr = XLSX.utils.encode_cell({ r: excelRow - 1, c: excelCol - 1 });

    const cell = ws[addr];
    if (!cell) continue;

    // 상대 경로 하이퍼링크
    // Excel은 파일 기준 상대경로를 해석합니다.
    // 안정적으로 쓰려면 xlsx를 폴더 트리와 가까운 위치에 저장하는 것이 좋습니다.
    cell.l = { Target: meta.relPath };

    // 파란색 링크 스타일(일부 뷰어/엑셀 설정에 따라 다를 수 있지만 최대한 반영)
    cell.s = {
      font: { color: { rgb: "0000FF" }, underline: true },
    };
  }
}

ui();