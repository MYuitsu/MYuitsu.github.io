// State
let sourceWB = null, sourceSheet = null, sourceType = '';
let codeWB = null, codeSheet = null;

// Elements
const fileSource = document.getElementById('file-source');
const fileCode = document.getElementById('file-code');
const sheetSource = document.getElementById('sheet-source');
const sheetCode = document.getElementById('sheet-code');
const dataType = document.getElementById('data-type');
const btnProcess = document.getElementById('btn-process');
const statusArea = document.getElementById('status-area');
const output = document.getElementById('output');
const infoSource = document.getElementById('info-source');
const infoCode = document.getElementById('info-code');

function esc(s) {
  if (s === undefined || s === null) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function showStatus(msg, type = 'info') {
  statusArea.innerHTML = `<div class="status ${type}">${esc(msg)}</div>`;
}

function checkReady() {
  const ready = sourceWB && sourceSheet && sourceType && codeWB && codeSheet;
  btnProcess.disabled = !ready;
}

// File 1: Source
fileSource.addEventListener('change', async (e) => {
  const f = e.target.files && e.target.files[0];
  if (!f) return;
  try {
    const ab = await f.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(ab);
    sourceWB = workbook;
    
    const sheetNames = workbook.worksheets.map(ws => ws.name);
    infoSource.textContent = `Đã đọc: ${f.name} — ${sheetNames.length} sheet(s)`;
    
    sheetSource.innerHTML = '';
    sheetNames.forEach(n => {
      const opt = document.createElement('option');
      opt.value = n; opt.textContent = n;
      sheetSource.appendChild(opt);
    });
    document.getElementById('sheet-source-row').style.display = 'flex';
    document.getElementById('type-row').style.display = 'flex';
    sourceSheet = sheetNames[0];
    checkReady();
  } catch (err) {
    console.error(err);
    infoSource.textContent = 'Lỗi khi đọc file nguồn.';
  }
});

sheetSource.addEventListener('change', () => {
  sourceSheet = sheetSource.value;
  checkReady();
});

dataType.addEventListener('change', () => {
  sourceType = dataType.value;
  checkReady();
});

// File 2: Code
fileCode.addEventListener('change', async (e) => {
  const f = e.target.files && e.target.files[0];
  if (!f) return;
  try {
    const ab = await f.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(ab);
    codeWB = workbook;
    
    const sheetNames = workbook.worksheets.map(ws => ws.name);
    infoCode.textContent = `Đã đọc: ${f.name} — ${sheetNames.length} sheet(s)`;
    
    sheetCode.innerHTML = '';
    sheetNames.forEach(n => {
      const opt = document.createElement('option');
      opt.value = n; opt.textContent = n;
      sheetCode.appendChild(opt);
    });
    document.getElementById('sheet-code-row').style.display = 'flex';
    codeSheet = sheetNames[0];
    checkReady();
  } catch (err) {
    console.error(err);
    infoCode.textContent = 'Lỗi khi đọc file mã.';
  }
});

sheetCode.addEventListener('change', () => {
  codeSheet = sheetCode.value;
  checkReady();
});

// Process
btnProcess.addEventListener('click', async () => {
  output.innerHTML = '';
  statusArea.innerHTML = '';
  
  try {
    // Lấy worksheet cần xử lý
    const worksheet = sourceWB.getWorksheet(sourceSheet);
    const codeWorksheet = codeWB.getWorksheet(codeSheet);
    
    // Hàm helper để lấy giá trị text thực sự từ cell
    function getCellText(cell) {
      if (!cell || cell.value === null || cell.value === undefined) return '';
      
      const val = cell.value;
      
      // Nếu là rich text
      if (val && typeof val === 'object' && val.richText) {
        return val.richText.map(rt => rt.text || '').join('');
      }
      
      // Nếu là formula
      if (val && typeof val === 'object' && val.result !== undefined) {
        return String(val.result);
      }
      
      // Nếu là hyperlink
      if (val && typeof val === 'object' && val.text !== undefined) {
        return String(val.text);
      }
      
      // Nếu là error
      if (val && typeof val === 'object' && val.error) {
        return String(val.error);
      }
      
      // Nếu là date
      if (val instanceof Date) {
        return val.toISOString();
      }
      
      // Plain value (number, string, boolean)
      if (typeof val === 'number' || typeof val === 'string' || typeof val === 'boolean') {
        return String(val);
      }
      
      // Fallback: convert object to string
      if (typeof val === 'object') {
        return JSON.stringify(val);
      }
      
      return String(val);
    }
    
    // Lấy mã từ file code (cột A và B)
    const codeSet = new Set();
    codeWorksheet.eachRow((row, rowNumber) => {
      const a = getCellText(row.getCell(1)).trim().toLowerCase();
      const b = getCellText(row.getCell(2)).trim().toLowerCase();
      if (a) codeSet.add(a);
      if (b) codeSet.add(b);
    });

    showStatus(`Đã tải ${codeSet.size} mã từ file mã (cột A, B). Đang tìm dòng cần xóa...`, 'info');

    // Xác định cột cần kiểm tra dựa trên loại dữ liệu
    let colIndices = [];
    if (sourceType === 'xetnghiem' || sourceType === 'cdha') {
      colIndices = [18, 21]; // R=18, U=21 (ExcelJS dùng 1-based index)
    } else if (sourceType === 'ttpt') {
      colIndices = [14, 15]; // N=14, O=15
    }

    // Tìm các dòng cần XÓA
    const rowsToDelete = [];
    worksheet.eachRow((row, rowNumber) => {
      let match = false;
      for (const colIdx of colIndices) {
        const val = getCellText(row.getCell(colIdx)).trim().toLowerCase();
        if (val && codeSet.has(val)) {
          match = true;
          break;
        }
      }
      
      if (match) {
        rowsToDelete.push(rowNumber);
      }
    });

    showStatus(`Tìm thấy ${rowsToDelete.length} dòng cần xóa. Đang xử lý...`, 'info');

    // === DEBUG: Hiển thị 30 cột và 30 dòng đầu tiên ===
    console.log('=== 📊 DEBUG: 30 DÒNG & 30 CỘT ĐẦU TIÊN ===');
    const maxDebugRows = Math.min(30, worksheet.rowCount);
    const maxDebugCols = 30;
    
    // Tạo header cho các cột
    const colHeaders = [''];
    for (let c = 1; c <= maxDebugCols; c++) {
      const colLetter = getColumnLetter(c);
      colHeaders.push(`[${colLetter}]`);
    }
    console.log(colHeaders.join('\t'));
    console.log('='.repeat(300));
    
    // Hiển thị từng dòng
    let debugRowCount = 0;
    worksheet.eachRow((row, rowNumber) => {
      if (debugRowCount >= maxDebugRows) return;
      
      const rowData = [`[${rowNumber}]`];
      for (let c = 1; c <= maxDebugCols; c++) {
        const cell = row.getCell(c);
        let val = getCellText(cell);
        
        // Xử lý giá trị để tránh ký tự lạ
        if (val) {
          val = val.replace(/\t/g, ' ').replace(/\n/g, ' ').replace(/\r/g, '');
          const displayVal = val.length > 15 ? val.substring(0, 12) + '...' : val;
          rowData.push(displayVal);
        } else {
          rowData.push('');
        }
      }
      
      // Đánh dấu nếu dòng này sẽ bị xóa
      const willDelete = rowsToDelete.includes(rowNumber);
      const prefix = willDelete ? '[XOA] ' : '[GIU]  ';
      console.log(prefix + rowData.join('\t'));
      
      debugRowCount++;
    });
    console.log('='.repeat(300));
    console.log(`Hiển thị ${debugRowCount} dòng đầu tiên × ${maxDebugCols} cột đầu tiên`);
    console.log('=== KẾT THÚC DEBUG ===\n');
    
    // Helper function để convert số cột thành chữ cái
    function getColumnLetter(col) {
      let letter = '';
      while (col > 0) {
        const remainder = (col - 1) % 26;
        letter = String.fromCharCode(65 + remainder) + letter;
        col = Math.floor((col - 1) / 26);
      }
      return letter;
    }

    // Log chi tiết trước khi xóa
    console.log('=== LOG SO SÁNH CHI TIẾT ===');
    console.log(`📊 Tổng số dòng ban đầu: ${worksheet.rowCount}`);
    console.log(`� Tổng số mã từ file Code: ${codeSet.size}`);
    console.log(`�🗑️ Số dòng cần xóa: ${rowsToDelete.length}`);
    console.log(`📋 Danh sách dòng sẽ xóa:`, rowsToDelete);
    
    // Hiển thị mẫu các mã trong codeSet
    console.log('\n🔑 Mẫu 10 mã đầu tiên trong file Code:');
    Array.from(codeSet).slice(0, 10).forEach((code, idx) => {
      console.log(`  ${idx + 1}. "${code}"`);
    });
    
    // Hiển thị chi tiết các dòng sẽ bị xóa với thông tin so sánh
    console.log('\n🔍 Chi tiết 5 dòng đầu tiên sẽ bị xóa (với giá trị so sánh):');
    rowsToDelete.slice(0, 5).forEach(rowNum => {
      const row = worksheet.getRow(rowNum);
      const rowData = [];
      
      // Log các cột đầu tiên
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (colNumber <= 3) {
          const cellText = getCellText(cell);
          rowData.push(`Cột ${colNumber}: "${cellText}"`);
        }
      });
      
      // Log giá trị trong các cột được kiểm tra
      const compareValues = [];
      for (const colIdx of colIndices) {
        const cell = row.getCell(colIdx);
        const val = getCellText(cell);
        const valLower = val.trim().toLowerCase();
        const matched = codeSet.has(valLower);
        compareValues.push(`Cột ${colIdx} (${String.fromCharCode(64 + colIdx)}): "${val}" → lowercase: "${valLower}" → Match: ${matched ? '✅' : '❌'}`);
      }
      
      console.log(`  📍 Dòng ${rowNum}:`);
      console.log(`     ${rowData.join(' | ')}`);
      console.log(`     ${compareValues.join('\n     ')}`);
    });

    // Tạo workbook mới thay vì xóa dòng
    const totalRowsBefore = worksheet.rowCount;
    
    // Tạo workbook mới và copy cấu trúc
    const newWorkbook = new ExcelJS.Workbook();
    const newWorksheet = newWorkbook.addWorksheet(sourceSheet);
    
    // Set các thuộc tính của worksheet
    newWorksheet.properties = { ...worksheet.properties };
    newWorksheet.pageSetup = { ...worksheet.pageSetup };
    newWorksheet.views = worksheet.views ? JSON.parse(JSON.stringify(worksheet.views)) : undefined;
    
    // Copy column widths và properties
    worksheet.columns.forEach((col, idx) => {
      const newCol = newWorksheet.getColumn(idx + 1);
      if (col.width) newCol.width = col.width;
      if (col.hidden) newCol.hidden = col.hidden;
      if (col.outlineLevel) newCol.outlineLevel = col.outlineLevel;
      if (col.style) newCol.style = col.style;
    });
    
    // Copy từng dòng KHÔNG nằm trong danh sách xóa
    const rowsToDeleteSet = new Set(rowsToDelete);
    let newRowIndex = 1;
    
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (!rowsToDeleteSet.has(rowNumber)) {
        const newRow = newWorksheet.getRow(newRowIndex);
        
        // Copy row height
        if (row.height) newRow.height = row.height;
        if (row.hidden) newRow.hidden = row.hidden;
        if (row.outlineLevel) newRow.outlineLevel = row.outlineLevel;
        
        // Copy từng cell với đầy đủ style và value
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const newCell = newRow.getCell(colNumber);
          
          // Copy value (bao gồm cả formula, richText, hyperlink)
          if (cell.value !== undefined && cell.value !== null) {
            if (cell.formula) {
              newCell.value = { formula: cell.formula };
            } else if (cell.value.richText) {
              newCell.value = { richText: JSON.parse(JSON.stringify(cell.value.richText)) };
            } else if (cell.value.hyperlink) {
              newCell.value = {
                text: cell.value.text,
                hyperlink: cell.value.hyperlink
              };
            } else {
              newCell.value = cell.value;
            }
          }
          
          // Copy style và bỏ strikethrough
          if (cell.style) {
            const cellStyle = JSON.parse(JSON.stringify(cell.style));
            
            // Bỏ strikethrough từ font
            if (cellStyle.font && cellStyle.font.strike) {
              cellStyle.font.strike = false;
            }
            
            newCell.style = cellStyle;
          }
          
          // Copy data validation
          if (cell.dataValidation) {
            newCell.dataValidation = JSON.parse(JSON.stringify(cell.dataValidation));
          }
          
          // Copy note/comment
          if (cell.note) {
            newCell.note = cell.note;
          }
        });
        
        newRow.commit();
        newRowIndex++;
      }
    });
    
    // Copy merged cells (adjust row indices)
    const rowMapping = new Map();
    let mappedRow = 1;
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (!rowsToDeleteSet.has(rowNumber)) {
        rowMapping.set(rowNumber, mappedRow);
        mappedRow++;
      }
    });
    
    if (worksheet.model && worksheet.model.merges) {
      worksheet.model.merges.forEach(merge => {
        const match = merge.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
        if (match) {
          const startCol = match[1];
          const startRow = parseInt(match[2]);
          const endCol = match[3];
          const endRow = parseInt(match[4]);
          
          // Chỉ copy merge nếu tất cả các dòng trong merge không bị xóa
          const allRowsPresent = Array.from({ length: endRow - startRow + 1 }, (_, i) => startRow + i)
            .every(r => !rowsToDeleteSet.has(r));
          
          if (allRowsPresent && rowMapping.has(startRow) && rowMapping.has(endRow)) {
            const newStartRow = rowMapping.get(startRow);
            const newEndRow = rowMapping.get(endRow);
            newWorksheet.mergeCells(`${startCol}${newStartRow}:${endCol}${newEndRow}`);
          }
        }
      });
    }
    
    // Thay thế workbook cũ bằng workbook mới
    sourceWB = newWorkbook;
    const totalRowsAfter = newWorksheet.rowCount;
    
    // Log kết quả sau khi xóa
    console.log('\n✅ KẾT QUẢ SAU KHI XÓA:');
    console.log(`📊 Tổng số dòng trước: ${totalRowsBefore}`);
    console.log(`📊 Tổng số dòng sau: ${totalRowsAfter}`);
    console.log(`🗑️ Đã xóa: ${totalRowsBefore - totalRowsAfter} dòng`);
    console.log(`✔️ Còn lại: ${totalRowsAfter} dòng`);
    console.log(`📈 Tỷ lệ xóa: ${((totalRowsBefore - totalRowsAfter) / totalRowsBefore * 100).toFixed(2)}%`);
    console.log('==============================\n');
    
    showStatus(`✅ Đã xóa ${rowsToDelete.length} dòng khỏi sheet "${sourceSheet}". Format, style, và merge cells được giữ nguyên! Xem console để biết chi tiết.`, 'success');
    
    // Tạo nút download
    createDownloadFromWorkbook(sourceWB);

  } catch (err) {
    console.error(err);
    showStatus('❌ Lỗi khi xử lý. Xem console để biết chi tiết.', 'error');
  }
});

async function createDownloadFromWorkbook(wb) {
  const btn = document.createElement('button');
  btn.textContent = '💾 Tải xuống Excel đã lọc';
  btn.style.marginTop = '12px';
  btn.style.marginLeft = '8px';
  btn.addEventListener('click', async () => {
    try {
      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `filtered_${Date.now()}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
      showStatus('✅ File đã được tải xuống!', 'success');
    } catch (err) {
      console.error(err);
      showStatus('❌ Lỗi khi tải file.', 'error');
    }
  });
  output.appendChild(btn);
}
