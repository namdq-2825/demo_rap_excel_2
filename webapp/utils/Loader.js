sap.ui.define([], function () {
  "use strict";

  return {
    init: function(oView) {
      this._oView = oView;
    },
    show: function (text = 'Đang tải dữ liệu...') {
        const oArea = this._oView.byId("customLoading").getDomRef();

        oArea.innerText = text;
    },

    hide: function() {
        const oArea = this._oView.byId("customLoading").getDomRef();

        oArea.innerText = '';
    }
  };
});
