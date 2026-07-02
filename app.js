(function () {
  "use strict";

  const folderInput = document.getElementById("folderInput");
  const filesInput = document.getElementById("filesInput");
  const selectedInfo = document.getElementById("selectedInfo");
  const statusPanel = document.getElementById("statusPanel");
  const classNav = document.getElementById("classNav");
  const summaryPanel = document.getElementById("summaryPanel");
  const resultPanel = document.getElementById("resultPanel");
  const unclassifiedPanel = document.getElementById("unclassifiedPanel");

  const HEADER_ALIASES = {
    학번: ["학번", "번호", "학생번호"],
    이름: ["이름", "성명", "학생이름"],
    수업: ["수업", "선택과목", "과목", "선택수업", "프로그램"],
  };

  folderInput.addEventListener("change", (e) => handleFileList(e.target.files));
  filesInput.addEventListener("change", (e) => handleFileList(e.target.files));

  function setStatus(message, isError) {
    statusPanel.hidden = !message;
    statusPanel.textContent = message || "";
    statusPanel.classList.toggle("error", !!isError);
  }

  function isSpreadsheetFile(file) {
    return /\.(xlsx|xls)$/i.test(file.name) && !file.name.startsWith("~$");
  }

  async function handleFileList(fileList) {
    const files = Array.from(fileList || []).filter(isSpreadsheetFile);

    if (files.length === 0) {
      setStatus("선택한 항목에 엑셀 파일(.xlsx, .xls)이 없습니다.", true);
      return;
    }

    selectedInfo.textContent = `${files.length}개 엑셀 파일 선택됨`;
    setStatus(`파일 ${files.length}개를 읽는 중...`, false);

    try {
      const allRows = [];
      const fileWarnings = [];

      for (const file of files) {
        const { rows, warning } = await parseWorkbookFile(file);
        allRows.push(...rows);
        if (warning) fileWarnings.push(`${file.name}: ${warning}`);
      }

      const { classes, unclassified } = groupByClass(allRows);
      render(classes, unclassified, files.length, allRows.length);

      const statusMsg = fileWarnings.length
        ? `처리 완료. 일부 파일에 경고가 있습니다 — ${fileWarnings.join(" / ")}`
        : `처리 완료: 파일 ${files.length}개, 레코드 ${allRows.length}건`;
      setStatus(statusMsg, fileWarnings.length > 0);
    } catch (err) {
      console.error(err);
      setStatus(`처리 중 오류가 발생했습니다: ${err.message}`, true);
    }
  }

  function readArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(file);
    });
  }

  function findHeaderKey(headerRow, aliases) {
    for (const alias of aliases) {
      const idx = headerRow.findIndex(
        (h) => String(h || "").trim() === alias
      );
      if (idx !== -1) return idx;
    }
    return -1;
  }

  async function parseWorkbookFile(file) {
    const buffer = await readArrayBuffer(file);
    const workbook = XLSX.read(buffer, { type: "array" });

    let sheetName = workbook.SheetNames.find((n) => n.toLowerCase() === "data");
    let warning = "";
    if (!sheetName) {
      sheetName = workbook.SheetNames[0];
      warning = `'data' 시트를 찾지 못해 '${sheetName}' 시트를 사용했습니다.`;
    }

    const sheet = workbook.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    if (grid.length === 0) return { rows: [], warning: warning || "빈 시트입니다." };

    const headerRow = grid[0];
    let idIdx = findHeaderKey(headerRow, HEADER_ALIASES.학번);
    let nameIdx = findHeaderKey(headerRow, HEADER_ALIASES.이름);
    let subjectIdx = findHeaderKey(headerRow, HEADER_ALIASES.수업);

    if (idIdx === -1 || nameIdx === -1 || subjectIdx === -1) {
      idIdx = 0;
      nameIdx = 1;
      subjectIdx = 2;
      warning = warning || "헤더명을 찾지 못해 앞 3개 컬럼(학번/이름/수업 순)을 사용했습니다.";
    }

    const rows = [];
    for (let i = 1; i < grid.length; i++) {
      const row = grid[i];
      const studentId = String(row[idIdx] ?? "").trim();
      const name = String(row[nameIdx] ?? "").trim();
      const subject = String(row[subjectIdx] ?? "").trim();
      if (!studentId && !name && !subject) continue;
      rows.push({ studentId, name, subject, sourceFile: file.name });
    }

    return { rows, warning };
  }

  function parseClass(studentId) {
    if (!/^\d{5,}$/.test(studentId)) {
      return { ok: false, reason: "학번이 5자리 숫자 형식이 아닙니다." };
    }
    const grade = Number(studentId[0]);
    const classNo = Number(studentId.slice(1, 3));
    if (grade < 1 || grade > 9 || classNo < 1) {
      return { ok: false, reason: "학번에서 학년/반 값을 해석할 수 없습니다." };
    }
    return { ok: true, grade, classNo, key: `${grade}-${classNo}` };
  }

  function groupByClass(rows) {
    const classes = new Map();
    const unclassified = [];

    for (const row of rows) {
      const parsed = parseClass(row.studentId);
      if (!parsed.ok) {
        unclassified.push({ ...row, reason: parsed.reason });
        continue;
      }

      if (!classes.has(parsed.key)) {
        classes.set(parsed.key, {
          grade: parsed.grade,
          classNo: parsed.classNo,
          students: new Map(),
        });
      }
      const cls = classes.get(parsed.key);

      if (!cls.students.has(row.studentId)) {
        cls.students.set(row.studentId, {
          studentId: row.studentId,
          name: row.name,
          subjects: [],
        });
      }
      const student = cls.students.get(row.studentId);
      if (row.subject && !student.subjects.includes(row.subject)) {
        student.subjects.push(row.subject);
      }
    }

    const sortedClasses = Array.from(classes.values()).sort(
      (a, b) => a.grade - b.grade || a.classNo - b.classNo
    );
    for (const cls of sortedClasses) {
      cls.studentList = Array.from(cls.students.values()).sort((a, b) =>
        a.studentId.localeCompare(b.studentId, undefined, { numeric: true })
      );
    }

    return { classes: sortedClasses, unclassified };
  }

  function el(tag, attrs, children) {
    const node = document.createElement(tag);
    if (attrs) {
      for (const [k, v] of Object.entries(attrs)) {
        if (k === "class") node.className = v;
        else if (k === "text") node.textContent = v;
        else node.setAttribute(k, v);
      }
    }
    (children || []).forEach((c) => c && node.appendChild(c));
    return node;
  }

  function render(classes, unclassified, fileCount, totalRecords) {
    resultPanel.innerHTML = "";
    classNav.innerHTML = "";
    summaryPanel.innerHTML = "";
    unclassifiedPanel.innerHTML = "";

    const totalStudents = classes.reduce((sum, c) => sum + c.studentList.length, 0);

    summaryPanel.hidden = false;
    [
      ["파일", fileCount],
      ["레코드", totalRecords],
      ["학급", classes.length],
      ["학생", totalStudents],
    ].forEach(([label, num]) => {
      summaryPanel.appendChild(
        el("div", { class: "summary-card" }, [
          el("div", { class: "num", text: String(num) }),
          el("div", { class: "label", text: label }),
        ])
      );
    });

    if (classes.length === 0) {
      resultPanel.appendChild(
        el("div", { class: "empty-state", text: "분류된 학급 데이터가 없습니다." })
      );
    } else {
      classNav.hidden = false;
      for (const cls of classes) {
        const anchorId = `class-${cls.grade}-${cls.classNo}`;
        classNav.appendChild(
          el("a", { href: `#${anchorId}`, text: `${cls.grade}학년 ${cls.classNo}반` })
        );

        const rowsBody = el("tbody", null, cls.studentList.map((s) =>
          el("tr", null, [
            el("td", { text: s.studentId }),
            el("td", { text: s.name }),
            el(
              "td",
              null,
              s.subjects.map((subj) => el("span", { class: "subject-badge", text: subj }))
            ),
          ])
        ));

        const table = el("table", null, [
          el("thead", null, [
            el("tr", null, [
              el("th", { text: "학번" }),
              el("th", { text: "이름" }),
              el("th", { text: "선택 과목" }),
            ]),
          ]),
          rowsBody,
        ]);

        resultPanel.appendChild(
          el("div", { class: "class-section", id: anchorId }, [
            el("h2", null, [
              document.createTextNode(`${cls.grade}학년 ${cls.classNo}반`),
              el("span", { class: "count", text: `${cls.studentList.length}명` }),
            ]),
            table,
          ])
        );
      }
    }

    if (unclassified.length > 0) {
      unclassifiedPanel.hidden = false;
      unclassifiedPanel.appendChild(
        el("h2", { text: `분류되지 않은 항목 (${unclassified.length}건)` })
      );
      unclassifiedPanel.appendChild(
        el("p", { text: "학번 형식이 예상과 달라 학급을 자동으로 판별하지 못했습니다." })
      );
      const table = el("table", null, [
        el("thead", null, [
          el("tr", null, [
            el("th", { text: "학번" }),
            el("th", { text: "이름" }),
            el("th", { text: "수업" }),
            el("th", { text: "파일" }),
            el("th", { text: "사유" }),
          ]),
        ]),
        el(
          "tbody",
          null,
          unclassified.map((row) =>
            el("tr", null, [
              el("td", { text: row.studentId }),
              el("td", { text: row.name }),
              el("td", { text: row.subject }),
              el("td", { text: row.sourceFile }),
              el("td", { text: row.reason }),
            ])
          )
        ),
      ]);
      unclassifiedPanel.appendChild(table);
    }
  }
})();
