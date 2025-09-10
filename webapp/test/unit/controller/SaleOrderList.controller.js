/*global QUnit*/

sap.ui.define([
	"zsalesorder/controller/SaleOrderList.controller"
], function (Controller) {
	"use strict";

	QUnit.module("SaleOrderList Controller");

	QUnit.test("I should test the SaleOrderList controller", function (assert) {
		var oAppController = new Controller();
		oAppController.onInit();
		assert.ok(oAppController);
	});

});
