(function () {
  const fileInput = document.getElementById("fileInput");
  const dropZone = document.getElementById("dropZone");
  const controls = document.getElementById("controls");
  const sheetSelect = document.getElementById("sheetSelect");
  const columnSelect = document.getElementById("columnSelect");
  const filenameInput = document.getElementById("filenameInput");
  const exportButton = document.getElementById("exportButton");
  const resetButton = document.getElementById("resetButton");
  const statusText = document.getElementById("statusText");
  const preview = document.getElementById("preview");
  const previewTable = document.getElementById("previewTable");
  const summaryText = document.getElementById("summaryText");
  const fixedColumnName = "商品重量";
  const insertedColumnName = "快递费";
  const removableRowColumnNames = ["商品名", "成本"];

  let workbook = null;
  let sourceFile = null;

  function setStatus(message, isError) {
    statusText.textContent = message;
    statusText.style.color = isError ? "#b42318" : "";
  }

  function safeName(name) {
    return (name || "导出的表格").replace(/\.[^.]+$/, "");
  }

  function formatError(error) {
    const message = String((error && error.message) || error || "");

    if (/array buffer allocation|invalid array length|out of memory|allocation failed/i.test(message)) {
      return "读取失败：表格有效范围过大，浏览器内存不足。请先在 Excel 中删除多余空白行/空白列，保存后再导入。";
    }

    return message || "操作失败，请确认文件格式是否正确。";
  }

  function getCellDisplayValue(cell) {
    if (!cell) return "";
    if (cell.w !== undefined) return cell.w;
    if (cell.v !== undefined) return cell.v;
    return "";
  }

  function hasCellContent(cell) {
    if (!cell) return false;
    if (cell.f !== undefined && String(cell.f).trim() !== "") return true;
    if (cell.v !== undefined && String(cell.v).trim() !== "") return true;
    if (cell.w !== undefined && String(cell.w).trim() !== "") return true;
    return false;
  }

  function getRows(sheetName) {
    const sheet = workbook.Sheets[sheetName];
    const cells = Object.keys(sheet).filter((key) => key[0] !== "!" && hasCellContent(sheet[key]));

    if (!cells.length) return [[]];

    let maxColumn = 0;
    const rowMap = new Map();
    const positions = cells.map((address) => {
      const position = XLSX.utils.decode_cell(address);
      maxColumn = Math.max(maxColumn, position.c);
      return { address, position };
    });

    positions.forEach(({ address, position }) => {
      if (!rowMap.has(position.r)) {
        rowMap.set(position.r, []);
      }
      const rowsForPosition = rowMap.get(position.r);
      while (rowsForPosition.length <= position.c) {
        rowsForPosition.push("");
      }
      rowsForPosition[position.c] = getCellDisplayValue(sheet[address]);
    });

    return Array.from(rowMap.keys()).sort((left, right) => left - right).map((rowIndex) => {
      const row = rowMap.get(rowIndex);
      while (row.length < maxColumn + 1) {
        row.push("");
      }
      return row;
    });
  }

  function getHeaders(rows) {
    const firstRow = rows[0] || [];
    return firstRow.map((value, index) => {
      const text = String(value || "").trim();
      return text || `未命名列 ${index + 1}`;
    });
  }

  function fillSelect(select, values) {
    select.innerHTML = "";
    values.forEach((value, index) => {
      const option = document.createElement("option");
      option.value = String(index);
      option.textContent = value;
      select.appendChild(option);
    });
  }

  function detectExtension(filename) {
    const match = filename.match(/\.([^.]+)$/);
    return match ? match[1].toLowerCase() : "";
  }

  function findFixedColumnIndex(headers) {
    const exactIndex = headers.findIndex((header) => header.trim() === fixedColumnName);
    if (exactIndex !== -1) return exactIndex;
    return headers.findIndex((header) => header.trim().includes(fixedColumnName));
  }

  function findHeaderIndex(headers, name) {
    const exactIndex = headers.findIndex((header) => header.trim() === name);
    if (exactIndex !== -1) return exactIndex;
    return headers.findIndex((header) => header.trim().includes(name));
  }

  function isBlank(value) {
    return String(value === undefined || value === null ? "" : value).trim() === "";
  }

  function removeBlankProductRows(rows) {
    const headers = getHeaders(rows);
    const indexes = removableRowColumnNames.map((name) => findHeaderIndex(headers, name));

    if (!rows.length || indexes.some((index) => index === -1)) return rows;

    return rows.filter((row, rowIndex) => {
      if (rowIndex === 0) return true;
      return !indexes.every((index) => isBlank(row[index]));
    });
  }

  function parseWeight(value) {
    if (typeof value === "number") return value;
    const text = String(value || "").trim().replace(/,/g, "");
    const match = text.match(/-?\d+(?:\.\d+)?/);
    return match ? Number(match[0]) : NaN;
  }

  function calculateShippingFee(weightValue) {
    const weight = parseWeight(weightValue);

    if (!Number.isFinite(weight) || weight === 0) return "请检查";
    if (weight > 0 && weight < 0.3) return 1.3;
    if (weight >= 0.3 && weight < 0.5) return 1.5;
    if (weight >= 0.5 && weight < 1) return 1.9;
    if (weight >= 1 && weight < 1.5) return 2.3;
    if (weight >= 1.5 && weight < 2) return 2.8;
    if (weight >= 2 && weight < 2.5) return 3.5;
    if (weight >= 2.5 && weight < 3) return 3.9;
    if (weight >= 3 && weight < 4) return 4.3;
    if (weight >= 4 && weight < 5) return 5.6;
    if (weight >= 5) return 13.5;
    return "请检查";
  }

  async function loadFile(file) {
    reset(false);
    sourceFile = file;
    setStatus("正在读取表格...");

    try {
      const data = await file.arrayBuffer();
      workbook = XLSX.read(data, {
        type: "array",
        cellDates: true,
        cellStyles: false
      });

      if (!workbook.SheetNames.length) {
        throw new Error("没有找到可读取的工作表。");
      }

      fillSelect(sheetSelect, workbook.SheetNames);
      filenameInput.value = `${safeName(file.name)}_新增快递费列.xlsx`;
      controls.hidden = false;
      updateColumns();
    } catch (error) {
      reset(false);
      setStatus(formatError(error), true);
    }
  }

  function updateColumns() {
    if (!workbook) return;

    let rows = [];
    const sheetName = workbook.SheetNames[Number(sheetSelect.value) || 0];

    try {
      rows = getRows(sheetName);
    } catch (error) {
      exportButton.disabled = true;
      preview.hidden = true;
      previewTable.innerHTML = "";
      setStatus(formatError(error), true);
      return;
    }

    const headers = getHeaders(rows);

    if (!headers.length) {
      fillSelect(columnSelect, ["未命名列 1"]);
    } else {
      fillSelect(columnSelect, headers);
    }

    const fixedColumnIndex = findFixedColumnIndex(headers);
    columnSelect.disabled = true;

    if (fixedColumnIndex === -1) {
      exportButton.disabled = true;
      preview.hidden = true;
      previewTable.innerHTML = "";
      setStatus(`当前工作表没有找到“${fixedColumnName}”这一列。`, true);
      return;
    }

    columnSelect.value = String(fixedColumnIndex);
    exportButton.disabled = false;
    setStatus(`已固定选择“${fixedColumnName}”列。`);
    renderPreview();
  }

  function getInsertSide() {
    const checked = document.querySelector("input[name='insertSide']:checked");
    return checked ? checked.value : "after";
  }

  function insertBlankColumn(rows, columnIndex, side) {
    const insertIndex = side === "before" ? columnIndex : columnIndex + 1;
    const normalizedRows = rows.length ? rows : [[]];

    return normalizedRows.map((row, rowIndex) => {
      const next = Array.isArray(row) ? row.slice() : [];
      while (next.length < insertIndex) {
        next.push("");
      }
      next.splice(insertIndex, 0, rowIndex === 0 ? insertedColumnName : calculateShippingFee(row[columnIndex]));
      return next;
    });
  }

  function renderPreview() {
    if (!workbook) return;

    const sheetName = workbook.SheetNames[Number(sheetSelect.value) || 0];
    let rows = [];

    try {
      rows = getRows(sheetName);
    } catch (error) {
      exportButton.disabled = true;
      preview.hidden = true;
      previewTable.innerHTML = "";
      setStatus(formatError(error), true);
      return;
    }

    const columnIndex = Number(columnSelect.value) || 0;
    const side = getInsertSide();
    const insertIndex = side === "before" ? columnIndex : columnIndex + 1;
    const previewRows = insertBlankColumn(removeBlankProductRows(rows), columnIndex, side).slice(0, 21);
    const headers = previewRows[0] || [];

    previewTable.innerHTML = "";
    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    headers.forEach((cell, index) => {
      const th = document.createElement("th");
      th.textContent = cell || (index === insertIndex ? insertedColumnName : "");
      if (index === insertIndex) th.className = "inserted";
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    previewTable.appendChild(thead);

    const tbody = document.createElement("tbody");
    previewRows.slice(1).forEach((row) => {
      const tr = document.createElement("tr");
      for (let index = 0; index < headers.length; index += 1) {
        const td = document.createElement("td");
        td.textContent = row[index] || "";
        if (index === insertIndex) td.className = "inserted";
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    });
    previewTable.appendChild(tbody);

    preview.hidden = false;
    summaryText.textContent = `${sheetName}，预览前 ${Math.min(rows.length, 21)} 行`;
  }

  function exportWorkbook() {
    if (!workbook) return;

    try {
      const sheetName = workbook.SheetNames[Number(sheetSelect.value) || 0];
      const columnIndex = Number(columnSelect.value) || 0;
      const side = getInsertSide();
      const rows = getRows(sheetName);
      const updatedRows = insertBlankColumn(removeBlankProductRows(rows), columnIndex, side);
      const outputName = filenameInput.value.trim() || `${safeName(sourceFile && sourceFile.name)}_新增快递费列.xlsx`;
      const finalName = detectExtension(outputName) === "xlsx" ? outputName : `${safeName(outputName)}.xlsx`;
      const nextWorkbook = {
        SheetNames: workbook.SheetNames.slice(),
        Sheets: Object.assign({}, workbook.Sheets),
        Props: Object.assign({}, workbook.Props || {})
      };

      nextWorkbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(updatedRows);
      setStatus("正在导出表格...");
      XLSX.writeFile(nextWorkbook, finalName, { bookType: "xlsx" });
      setStatus(`已导出：${finalName}`);
    } catch (error) {
      setStatus(formatError(error), true);
    }
  }

  function reset(clearInput) {
    workbook = null;
    sourceFile = null;
    controls.hidden = true;
    preview.hidden = true;
    previewTable.innerHTML = "";
    exportButton.disabled = true;
    sheetSelect.innerHTML = "";
    columnSelect.innerHTML = "";
    columnSelect.disabled = false;
    filenameInput.value = "";
    if (clearInput !== false) fileInput.value = "";
    setStatus("");
  }

  fileInput.addEventListener("change", () => {
    const file = fileInput.files && fileInput.files[0];
    if (file) loadFile(file);
  });

  dropZone.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropZone.classList.add("dragging");
  });

  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("dragging");
  });

  dropZone.addEventListener("drop", (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
    const file = event.dataTransfer.files && event.dataTransfer.files[0];
    if (file) loadFile(file);
  });

  sheetSelect.addEventListener("change", updateColumns);
  columnSelect.addEventListener("change", renderPreview);
  document.querySelectorAll("input[name='insertSide']").forEach((radio) => {
    radio.addEventListener("change", renderPreview);
  });
  exportButton.addEventListener("click", exportWorkbook);
  resetButton.addEventListener("click", () => reset(true));
})();
