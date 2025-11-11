sap.ui.define([], function () {
  "use strict";

  return {
    loadFile: function(file) {
        return new Promise((resolve, reject) => {
            fetch(file, {
                method: "GET",
                headers: {
                    "Accept": "application/json;odata.metadata=minimal;IEEE754Compatible=true"
                }
            })
            .then(response => response.arrayBuffer())
            .then(async (f) => {
                resolve(f);
            })
            .catch((error) => {
                reject(error);
            })
        });
    },

    replaceTableVar: function ({
        worksheet,
        tableMarker,
        dataTable
    }) {
        let markerRowNumber;
        worksheet.eachRow((row, rowNumber) => {
            if (row.values.some(v => typeof v === "string" && v.includes(tableMarker))) {
                markerRowNumber = rowNumber;
            }
        });

        if (!markerRowNumber) {
            throw new Error("Không tìm thấy marker row trong template");
        }

        // Xoá dòng marker
        worksheet.spliceRows(markerRowNumber, 1);

        dataTable.forEach(item => {
            worksheet.insertRow(markerRowNumber, Object.values(item));
            markerRowNumber++;
        });

        return worksheet;
    },

    // NOTE:
    // replacements = {
    //     "%SO_ID": "1212123",
    //     "%sold_to_party": "CUST_001",
    //     "%customer_name": "Nguyen Van A"
    // };
    replaceSingleVar: function ({ workbook, worksheet, replacements }) {
        workbook.eachSheet((worksheet) => {
            worksheet.eachRow((row) => {
                row.eachCell((cell) => {
                    if (typeof cell.value === "string") {
                        Object.keys(replacements).forEach(key => {
                            if (cell.value.includes(key)) {
                                cell.value = cell.value.replace(key, replacements[key]);
                            }
                        });
                    }
                });
            });
        });

        return {
            workbook,
            worksheet,
        };
    },

    // MAINTAIN: use worker
    handleExport: function (buffer) {
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });


        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = "TemplateExport.xlsx";
        link.click();
    },

    // For library: XlsxPopulate
    xpReplaceByCoords: function (worksheet, replacements) {
        Object.keys(replacements).forEach((valueKey) => {
            const value = replacements[valueKey]
            worksheet.cell(valueKey).value(value)

        })
    },

    // For library: XlsxPopulate
    // replacements: { "%_SO_ID_%": "5000001234", "%_CUSTOMER_%": "ACME" }
    xpReplaceSingleVar: async function (sheet, replacements) {
            const used = sheet.usedRange();               // vùng thực sự có dữ liệu
            if (!used) return;
            used.cells().forEach(row => {
                row.forEach((cell) => {
                    const v = cell._value;
                    console.log(v);
                    if (typeof v === "string") {
                        let out = v;
                        for (const [k, val] of Object.entries(replacements)) {
                            if (out.includes(k)) out = out.split(k).join(String(val));
                        }
                        if (out !== v) cell.value(out);
                    }
                })
            });
    },

    // For library: XlsxPopulate
        /**
     * sheet: XlsxPopulate.Sheet
     * tableMarker: "%_SALES_ORDER_TS_%"
     * rows: mảng object [{...}, ...]
     * columns: mảng key để xác định thứ tự cột, ví dụ ["soID","customer","soldToParty","saleOrg"]
     */
    xpReplaceTable: function (sheet, { tableMarker, rows, columns }) {
        let startRow = null, startCol = null;

        const used = sheet.usedRange();
        if (!used) throw new Error("Sheet rỗng, không tìm thấy marker");

        used.cells().some(cell => {
            const v = cell.value();
            if (typeof v === "string" && v.includes(tableMarker)) {
            const addr = cell.address(); // { rowNumber, columnNumber }
            startRow = addr.rowNumber;
            startCol = addr.columnNumber;
            cell.value(null);            // xoá marker
            return true;
            }
            return false;
        });

        if (!startRow || !startCol) {
            throw new Error(`Không tìm thấy marker ${tableMarker}`);
        }

        // Chuẩn bị mảng 2D theo thứ tự cột mong muốn
        const data2D = rows.map(r => columns.map(c => r[c]));

        if (data2D.length) {
            const endRow = startRow + data2D.length - 1;
            const endCol = startCol + columns.length - 1;

            // Ghi thẳng 1 lần vào range
            sheet.range(startRow, startCol, endRow, endCol).value(data2D);
        }

        return { startRow, startCol };
    },


  }
});
