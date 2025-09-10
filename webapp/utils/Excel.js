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
    }
  };
});
