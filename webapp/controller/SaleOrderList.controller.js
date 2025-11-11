sap.ui.define([
    "sap/m/MessageToast",
    "sap/ui/core/mvc/Controller",
    "zsalesorder/utils/Excel",
    "zsalesorder/utils/Print",
    "zsalesorder/utils/Loader",
], (MessageToast, Controller, Excel, Print, Loader) => {
    "use strict";

    return Controller.extend("zsalesorder.controller.SaleOrderList", {
        onInit() {
            Loader.init(this.getView());
        },

        async handleExportExcel() {
            const templateListElm = this.byId("templateList");
            const salesOrderTableElm = this.byId("SalesOrderList");

            const selectedTemplate = templateListElm.mProperties.selectedKey;
            const [ aliasTemplate, uuidTemplate ] = selectedTemplate.split(",");

            if (!aliasTemplate || !uuidTemplate) {
                MessageToast.show("Please select export template !");

                return
            }

            const sServiceUrl = "/sap/opu/odata4/sap/z_salesorder__o4_sb/srvd/sap/z_salesorder_1_sd/0001/";

            const templateUrl = sServiceUrl + `TemplateExport(Uuid=${uuidTemplate},TemplateAlias='${aliasTemplate}',IsActiveEntity=true)/Attachment`;

            await import("https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js");

            if (aliasTemplate === 'SALE_ORDER_LIST_TEMPLATE') {
                const tableData = salesOrderTableElm.getItems().map((item) => item.getBindingContext().getObject())

                const soList = tableData.map((item) => ({
                    soID: item.SalesOrder,
                    customer: item.CustomerName,
                    soldToParty: item.SoldToParty,
                    saleOrg: item.SalesOrganization,
                }));

                Excel.loadFile(templateUrl).then(async (f) => {
                    const workbook = new ExcelJS.Workbook();
                    await workbook.xlsx.load(f);

                    const worksheet = workbook.getWorksheet(1);

                    Excel.replaceTableVar({
                        worksheet,
                        tableMarker: '%_SALES_ORDER_TS_%',
                        dataTable: soList,
                    });

                    const buffer = await workbook.xlsx.writeBuffer();

                    Excel.handleExport(buffer);
                })
            }

            if (aliasTemplate === 'SALE_ORDER_DETAIL_TEMPLATE') {
                const selectedItem = salesOrderTableElm.getSelectedItems()
                if (!selectedItem.length) {
                    MessageToast.show("Please select 1 sale order !");

                    return 
                }

                const data = selectedItem[0].getBindingContext().getObject();

                Excel.loadFile(templateUrl).then(async (f) => {
                    const workbook = new ExcelJS.Workbook();
                    await workbook.xlsx.load(f);

                    const worksheet = workbook.getWorksheet(1);

                    Excel.replaceSingleVar({
                        workbook,
                        worksheet,
                        replacements: {
                            "%SO_ID": data.SalesOrder,
                            "%sold_to_party": data.SoldToParty,
                            "%customer_name": data.CustomerName,
                            "%sale_org": data.SalesOrganization,
                        }
                    })

                    const buffer = await workbook.xlsx.writeBuffer();

                    Excel.handleExport(buffer);
                })
            }

            if (aliasTemplate === 'DEMO_CHART') {
                await import("https://cdnjs.cloudflare.com/ajax/libs/xlsx-populate/1.21.0/xlsx-populate.min.js");

                Excel.loadFile(templateUrl).then(async (f) => {

                    const workbook = await XlsxPopulate.fromDataAsync(f);
                    const worksheet = workbook.sheet('Sheet1')

                    Excel.xpReplaceByCoords(    
                        worksheet,
                        {
                            "B2": 1000,
                            "B3": 2000,
                            "B4": 1500,
                            "B5": 1000,
                            "B6": 3000,
                        }
                    );

                    Excel.xpReplaceSingleVar(    
                        worksheet,
                        {
                            "%label_1%": 'Price 1',
                            "%label_2%": 'Price 2',
                            "%label_3%": 'Price 3',
                            "%label_4%": 'Price 4',
                            "%label_5%": 'Price 5',
                        }
                    );

                    const blob = await workbook.outputAsync({ type: "blob" });
                    Excel.handleExport(await blob.arrayBuffer());
                })
            }

            if (aliasTemplate === 'DEMO_MERGE_CELL') {
                Excel.loadFile(templateUrl).then(async (f) => {
                    const workbook = new ExcelJS.Workbook();
                    await workbook.xlsx.load(f);

                    const worksheet = workbook.getWorksheet(1);

                    const data = [
                        {
                            productName: 'Product A',
                            groupData: [
                            { customer: 'Alice', price: 12, itemSold: 3, total: 36 },
                            { customer: 'Bob', price: 10, itemSold: 2, total: 20 }
                            ]
                        },
                        {
                            productName: 'Product B',
                            groupData: [
                            { customer: 'Charlie', price: 8, itemSold: 5, total: 40 },
                            { customer: 'David', price: 15, itemSold: 1, total: 15 },
                            { customer: 'Emma', price: 7, itemSold: 4, total: 28 }
                            ]
                        },
                        {
                            productName: 'Product C',
                            groupData: [
                            { customer: 'Frank', price: 20, itemSold: 1, total: 20 },
                            { customer: 'Grace', price: 25, itemSold: 2, total: 50 }
                            ]
                        },
                        {
                            productName: 'Product D',
                            groupData: [
                            { customer: 'Hannah', price: 18, itemSold: 3, total: 54 },
                            { customer: 'Ian', price: 5, itemSold: 10, total: 50 }
                            ]
                        },
                        {
                            productName: 'Product E',
                            groupData: [
                            { customer: 'Jack', price: 9, itemSold: 7, total: 63 },
                            { customer: 'Karen', price: 14, itemSold: 4, total: 56 },
                            { customer: 'Leo', price: 11, itemSold: 8, total: 88 }
                            ]
                        },
                        {
                            productName: 'Product F',
                            groupData: [
                            { customer: 'Mia', price: 30, itemSold: 1, total: 30 },
                            { customer: 'Nina', price: 6, itemSold: 12, total: 72 }
                            ]
                        },
                        {
                            productName: 'Product G',
                            groupData: [
                            { customer: 'Oscar', price: 16, itemSold: 5, total: 80 },
                            { customer: 'Paul', price: 13, itemSold: 9, total: 117 }
                            ]
                        },
                        {
                            productName: 'Product H',
                            groupData: [
                            { customer: 'Quinn', price: 22, itemSold: 3, total: 66 },
                            { customer: 'Rachel', price: 19, itemSold: 4, total: 76 },
                            { customer: 'Sam', price: 28, itemSold: 2, total: 56 }
                            ]
                        },
                        {
                            productName: 'Product I',
                            groupData: [
                            { customer: 'Tom', price: 17, itemSold: 6, total: 102 },
                            { customer: 'Uma', price: 21, itemSold: 2, total: 42 }
                            ]
                        },
                        {
                            productName: 'Product J',
                            groupData: [
                            { customer: 'Victor', price: 12, itemSold: 7, total: 84 },
                            { customer: 'Wendy', price: 8, itemSold: 6, total: 48 }
                            ]
                        },
                        {
                            productName: 'Product K',
                            groupData: [
                            { customer: 'Xavier', price: 19, itemSold: 3, total: 57 },
                            { customer: 'Yara', price: 15, itemSold: 4, total: 60 }
                            ]
                        },
                        {
                            productName: 'Product L',
                            groupData: [
                            { customer: 'Zack', price: 10, itemSold: 9, total: 90 },
                            { customer: 'Alice', price: 14, itemSold: 5, total: 70 }
                            ]
                        },
                        {
                            productName: 'Product M',
                            groupData: [
                            { customer: 'Bob', price: 18, itemSold: 2, total: 36 },
                            { customer: 'Charlie', price: 9, itemSold: 11, total: 99 },
                            { customer: 'David', price: 25, itemSold: 1, total: 25 }
                            ]
                        },
                        {
                            productName: 'Product N',
                            groupData: [
                            { customer: 'Emma', price: 20, itemSold: 2, total: 40 },
                            { customer: 'Frank', price: 7, itemSold: 8, total: 56 }
                            ]
                        },
                        {
                            productName: 'Product O',
                            groupData: [
                            { customer: 'Grace', price: 16, itemSold: 7, total: 112 },
                            { customer: 'Hannah', price: 12, itemSold: 6, total: 72 }
                            ]
                        },
                        {
                            productName: 'Product P',
                            groupData: [
                            { customer: 'Ian', price: 30, itemSold: 3, total: 90 },
                            { customer: 'Jack', price: 11, itemSold: 10, total: 110 }
                            ]
                        },
                        {
                            productName: 'Product Q',
                            groupData: [
                            { customer: 'Karen', price: 28, itemSold: 2, total: 56 },
                            { customer: 'Leo', price: 13, itemSold: 5, total: 65 },
                            { customer: 'Mia', price: 9, itemSold: 9, total: 81 }
                            ]
                        },
                        {
                            productName: 'Product R',
                            groupData: [
                            { customer: 'Nina', price: 22, itemSold: 3, total: 66 },
                            { customer: 'Oscar', price: 19, itemSold: 4, total: 76 }
                            ]
                        },
                        {
                            productName: 'Product S',
                            groupData: [
                            { customer: 'Paul', price: 15, itemSold: 7, total: 105 },
                            { customer: 'Quinn', price: 26, itemSold: 2, total: 52 }
                            ]
                        },
                        {
                            productName: 'Product T',
                            groupData: [
                            { customer: 'Rachel', price: 17, itemSold: 8, total: 136 },
                            { customer: 'Sam', price: 12, itemSold: 5, total: 60 }
                            ]
                        }
                    ];
                    
                    let startRow = null;
                    let startCol = null;

                    worksheet.eachRow((row, rowNumber) => {
                        row.eachCell((cell, colNumber) => {
                            if (cell.value === '%_PRODUCTS_TS_%') {
                                startRow = rowNumber;
                                startCol = colNumber;
                            }
                        });
                    });

                    if (!startRow || !startCol) {
                        throw new Error("Không tìm thấy marker %_PRODUCTS_TS_% trong file mẫu");
                    }

                    // Xóa marker
                    worksheet.getCell(startRow, startCol).value = null;

                    // Đổ data
                    let currentRow = startRow;
                    data.forEach((product) => {
                        const { productName, groupData } = product;
                        const groupSize = groupData.length;

                        // Merge cell cho productName
                        if (groupSize > 1) {
                            worksheet.mergeCells(currentRow, startCol, currentRow + groupSize - 1, startCol);
                        }
                        worksheet.getCell(currentRow, startCol).value = productName;

                        // Đổ groupData
                        groupData.forEach((g, index) => {
                            const row = worksheet.getRow(currentRow + index);

                            row.getCell(startCol + 1).value = g.customer;
                            row.getCell(startCol + 2).value = g.price;
                            row.getCell(startCol + 3).value = g.itemSold;
                            row.getCell(startCol + 4).value = g.total;

                            // Apply border cho tất cả cell từ A → E
                            for (let col = startCol; col <= startCol + 4; col++) {
                                row.getCell(col).border = {
                                    top: { style: 'thin' },
                                    left: { style: 'thin' },
                                    bottom: { style: 'thin' },
                                    right: { style: 'thin' }
                                };
                            }

                            row.commit();
                        });

                        currentRow += groupSize;
                    });

                    const buffer = await workbook.xlsx.writeBuffer();

                    Excel.handleExport(buffer);
                })
            }

        },

        handleExportPDF() {
            const salesOrderTableElm = this.byId("SalesOrderList");
            const tableData = salesOrderTableElm.getItems().map((item) => item.getBindingContext().getObject())

            const genHtml = (data, index) => {
                return `
                    <!DOCTYPE html>
                        <html lang="en">
                        <head>
                            <meta charset="UTF-8" />
                            <meta name="viewport" content="width=device-width, initial-scale=1.0" />
                            <meta http-equiv="X-UA-Compatible" content="ie=edge" />
                            <title>Sales Order Report</title>
                            <style>
                            h1 {
                                text-align: center;
                                margin-bottom: 24px;
                            }

                            table {
                                border-collapse: collapse;
                            }

                            td,
                            th {
                                padding: 8px 16px;
                                border: 1px solid #333333;
                            }
                            </style>
                        </head>
                        <body>
                            <h1>SALES ORDER REPOST ${index}</h1>
                            <table>
                            <thead>
                                <th>Number</th>
                                <th>Sale to Party</th>
                                <th>Customer</th>
                                <th>Sales Organization</th>
                                <th>Distribution Channel</th>
                            </thead>
                            <tbody>
                                ${ data.map((item) => `
                                    <tr>
                                        <td>${item.SalesOrder}</td>
                                        <td>${item.SoldToParty}</td>
                                        <td>${item.CustomerName}</td>
                                        <td>${item.SalesOrganization}</td>
                                        <td>${item.DistributionChannel}</td>
                                    </tr>    
                                `).join('') }
                            </tbody>
                            </table>
                        </body>
                        </html>

                `
            }


            Print.processPrintQueue(0, [genHtml(tableData, 1), genHtml(tableData, 2)])
        },

        handleExportPDFWithChart: function () {
            const salesOrderTableElm = this.byId("SalesOrderList");
            const tableData = salesOrderTableElm.getItems().map((item) => item.getBindingContext().getObject())

            const genHtml = (data, index) => {
                return `
                    <!DOCTYPE html>
                    <html lang="en">
                    <head>
                        <meta charset="UTF-8" />
                        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
                        <meta http-equiv="X-UA-Compatible" content="ie=edge" />
                        <title>Sales Order Report</title>
                        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
                        <style>
                            h1 {
                                text-align: center;
                                margin-bottom: 24px;
                            }

                            table {
                                border-collapse: collapse;
                                margin-bottom: 40px;
                                width: 100%;
                            }

                            td,
                            th {
                                padding: 8px 16px;
                                border: 1px solid #333333;
                            }

                            .chart-container {
                                width: 600px;
                                height: 400px;
                                margin: 0 auto;
                            }
                        </style>
                    </head>
                    <body>
                        <h1>SALES ORDER REPORT ${index}</h1>

                        <div class="chart-container">
                            <canvas id="myChart"></canvas>
                        </div>

                        <table>
                            <thead>
                                <tr>
                                    <th>Number</th>
                                    <th>Sale to Party</th>
                                    <th>Customer</th>
                                    <th>Sales Organization</th>
                                    <th>Distribution Channel</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${ data.map((item) => `
                                    <tr>
                                        <td>${item.SalesOrder}</td>
                                        <td>${item.SoldToParty}</td>
                                        <td>${item.CustomerName}</td>
                                        <td>${item.SalesOrganization}</td>
                                        <td>${item.DistributionChannel}</td>
                                    </tr>    
                                `).join('') }
                            </tbody>
                        </table>

                        <script>
                            const ctx = document.getElementById('myChart');
                            const chart = new Chart(ctx, {
                                type: 'bar',
                                data: {
                                    labels: ${JSON.stringify(data.map(d => d.CustomerName))},
                                    datasets: [{
                                        label: 'Sales Orders',
                                        data: ${JSON.stringify(data.map(d => d.SalesOrder))},
                                        backgroundColor: 'rgba(54, 162, 235, 0.6)',
                                        borderColor: 'rgba(54, 162, 235, 1)',
                                        borderWidth: 1
                                    }]
                                },
                                options: {
                                    responsive: true,
                                    animation: false,
                                    plugins: {
                                        legend: { position: 'top' },
                                        title: { display: true, text: 'Sales Orders per Customer' }
                                    }
                                }
                            });
                        </script>
                    </body>
                    </html>
                `;
            }

            Print.processPrintQueue(0, [genHtml(tableData, 1)])
        },


        handleExportExcelByRAP() {
            const downloadExcelFromString = (binaryString, fileName) => {
                const dataUri = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${binaryString}`;
                const link = document.createElement('a');
                link.href = dataUri;
                link.setAttribute('download', fileName);
                document.body.appendChild(link);

                // 4. Tự động "click" vào link để trình duyệt bắt đầu tải file
                link.click();

                // 5. Dọn dẹp bằng cách xóa link khỏi trang
                link.parentNode.removeChild(link);
            }


            var sURI = "/sap/opu/odata4/sap/z_salesorder__o4_sb/srvd/sap/z_salesorder_1_sd/0001/";
            var oDataModel = new sap.ui.model.odata.v4.ODataModel({
                serviceUrl: sURI,
                synchronizationMode: "None",   // bắt buộc cho OData V4
                operationMode: "Server",       // server-side paging/filtering/sorting
                autoExpandSelect: true         // tự động $expand, $select
            });
            this.getView().setModel(oDataModel);


            var oListBinding = oDataModel.bindList("/ZCE_SALES_ORDER_REPORT");

            oListBinding.requestContexts().then(function (aContexts) {
                aContexts.forEach(function (oContext) {
                    const data = oContext.getObject()
                    var sFileName = "report.xlsx";
                    const sBase64 = data.fileBase64;

                    downloadExcelFromString(convertBase64UrlToBase64(sBase64), sFileName)
                });
            });
        },

        async handleExportLargeData() {
            const vbapData = [];
            let totalRecord = 0;
            const sServiceUrl = "/sap/opu/odata4/sap/z_salesorder__o4_sb/srvd/sap/z_salesorder_1_sd/0001/";

            const fetchData = async (skip) => {
                const perPage = 50000;

                return new Promise((resolve, reject) => {

                    fetch(sServiceUrl + `ZCE_SALEORDER_LIST?$select=*&$skip=${skip}&$top=${perPage}&$count=true`, {
                        method: "GET",
                        headers: {
                            "Accept": "application/json;odata.metadata=minimal"
                        }
                    })
                    .then(response => response.json())
                    .then(async (res) => {
                        vbapData.push(...res.value);
                        totalRecord = res['@odata.count']
                        console.log(res.value, res.value.length, perPage, totalRecord)
                        if (res.value.length === perPage) {
                            Loader.show(`Đang tải dữ liệu... (${Number(skip).toLocaleString()}/${Number(totalRecord).toLocaleString()})`)
                            resolve(await fetchData(skip + perPage));
                        } else {
                            Loader.show(`Đang tải dữ liệu... (${Number(vbapData.length).toLocaleString()}/${Number(totalRecord).toLocaleString()})`)
                            resolve(vbapData)
                        }
                    })
                    .catch((error) => {
                    })
                })
            }

            const templateListElm = this.byId("templateList");
            await import("https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js");

            const selectedTemplate = templateListElm.mProperties.selectedKey;
            const [ aliasTemplate, uuidTemplate ] = selectedTemplate.split(",");

            if (!aliasTemplate || !uuidTemplate) {
                MessageToast.show("Please select export template !");

                return
            }


            const templateUrl = sServiceUrl + `TemplateExport(Uuid=${uuidTemplate},TemplateAlias='${aliasTemplate}',IsActiveEntity=true)/Attachment`;

            Loader.show('Đang tải dữ liệu...')

            fetchData(0).then((data) => {

                setTimeout(() => Loader.show('Đang tạo file...'), 500)

                Excel.loadFile(templateUrl).then(async (f) => {
                    const workbook = new ExcelJS.Workbook();
                    await workbook.xlsx.load(f);
                    const worksheet = workbook.getWorksheet(1);

                    Excel.replaceTableVar({
                        worksheet,
                        tableMarker: '%_SALES_ORDER_TS_%',
                        dataTable: data,
                    });

                    const buffer = await workbook.xlsx.writeBuffer();

                    Loader.hide()

                    Excel.handleExport(buffer);
                })
            })


        }
    });
});