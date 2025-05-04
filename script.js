// script.js
let rawData = [];
let transformedData = [];

document.getElementById("upload").addEventListener("change", handleFile, false);
document.getElementById("processBtn").addEventListener("click", processData);
document.getElementById("downloadBtn").addEventListener("click", downloadExcel);

function handleFile(e) {
  const file = e.target.files[0];
  const reader = new FileReader();
  const status = document.getElementById("status");

  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    status.textContent = `Loaded ${rawData.length - 1} rows.`;
    document.getElementById("processBtn").disabled = false;
  };
  reader.readAsArrayBuffer(file);
}

function processData() {
  transformedData = [["Brand", "Main Category", "Sub Category", "SKU"]];

  for (let r = 1; r < rawData.length; r++) {
    let row = rawData[r];
    for (let c = 0; c < row.length; c += 3) {
      let brand = row[c] || "";
      let subCategory = row[c + 1] || "";
      let sku = row[c + 2] || "";

      if (!brand && !subCategory && !sku) continue;

      let mainCategory = "";
      if (/Atta|Maida/i.test(subCategory)) mainCategory = "Flour";
      else if (/Semolina/i.test(subCategory)) mainCategory = "Semolina";
      else if (/Lentil/i.test(subCategory)) mainCategory = "Lentil";
      else if (/Mustard Oil/i.test(subCategory)) mainCategory = "Mustard Oil";
      else if (/Puffed Rice/i.test(subCategory)) mainCategory = "Puffed Rice";
      else if (/Rice Chinigura|Rice Others/i.test(subCategory)) mainCategory = "Rice";
      else if (/Salt/i.test(subCategory)) mainCategory = "Salt";

      transformedData.push([brand, mainCategory, subCategory, sku]);
    }
  }

  document.getElementById("status").textContent = `Transformed ${transformedData.length - 1} rows.`;
  document.getElementById("downloadBtn").disabled = false;
  previewTable(transformedData);
}

function previewTable(data) {
  const table = document.getElementById("preview");
  table.innerHTML = "";
  data.forEach((row, i) => {
    let tr = document.createElement("tr");
    row.forEach(cell => {
      let td = document.createElement(i === 0 ? "th" : "td");
      td.textContent = cell;
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
}

function downloadExcel() {
  const ws = XLSX.utils.aoa_to_sheet(transformedData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Transformed");
  XLSX.writeFile(wb, "transformed_data.xlsx");
}
