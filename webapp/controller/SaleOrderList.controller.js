sap.ui.define([
    "sap/m/MessageToast",
    "sap/ui/core/mvc/Controller",
    "zsalesorder/utils/Excel"
], (MessageToast, Controller, Excel) => {
    "use strict";

    return Controller.extend("zsalesorder.controller.SaleOrderList", {
        onInit() {
        },

        async handleExport() {
            const templateListElm = this.byId("templateList");
            const salesOrderTableElm = this.byId("SalesOrderList");
            await import("https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js");

            const selectedTemplate = templateListElm.mProperties.selectedKey;
            const [ aliasTemplate, uuidTemplate ] = selectedTemplate.split(",");

            if (!aliasTemplate || !uuidTemplate) {
                MessageToast.show("Please select export template !");

                return
            }

            const sServiceUrl = "/sap/opu/odata4/sap/z_salesorder__o4_sb/srvd/sap/z_salesorder_1_sd/0001/";

            const templateUrl = sServiceUrl + `TemplateExport(Uuid=${uuidTemplate},TemplateAlias='${aliasTemplate}',IsActiveEntity=true)/Attachment`;

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

        }
    });
});