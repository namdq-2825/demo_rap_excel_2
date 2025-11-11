sap.ui.define([], function () {
  "use strict";

  return {
        /**
         * Hàm xử lý hàng đợi in, in từng tài liệu một cách tuần tự.
         * @param {number} jobIndex - Chỉ số của tài liệu cần in trong mảng printJobsData.
         * @param {string[]} printJobsData - Danh sách các html cần in. 
         * @param printWindow
         */
        processPrintQueue: function (jobIndex, printJobsData, printWindow) {
            const that = this;
            if (jobIndex >= printJobsData.length) {
                console.log("Hoàn thành tất cả các lần in!");
                if (printWindow && !printWindow.closed) {
                    printWindow.close();
                }
                return;
            }

            console.log(`Bắt đầu in tài liệu số ${jobIndex + 1}...`);
            const htmlContent = printJobsData[jobIndex];
            
            // Nếu chưa có window thì mở 1 lần
            if (!printWindow || printWindow.closed) {
                printWindow = window.open("", "PrintWindow", "width=1000px,height=1000px");
            }

            if (printWindow) {
                printWindow.document.open();
                printWindow.document.write(htmlContent);
                printWindow.document.close();

                printWindow.onafterprint = function () {
                    that.processPrintQueue(jobIndex + 1, printJobsData, printWindow);
                };

                setTimeout(function () {
                    printWindow.print();
                }, 500);
            } else {
                alert("Vui lòng cho phép mở cửa sổ popup để tiếp tục quá trình in.");
            }
        }
  };
});
