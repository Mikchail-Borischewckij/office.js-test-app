/// <reference path="../App.js" />

(function () {
	"use strict";

	var meetDocPattern = /(\[|\()[A-Za-z0-9_=]+(\]|\))/g;
	var meetXUrl = "http://localhost:8080/officejs-app/App/Home.html";
	Office.initialize = function (reason) {
		$(document).ready(function () {
			var docUrl = Office.context.document.url;
			var url = !docUrl ? meetXUrl : meetXUrl + "?docName=" + extractDocumentName(docUrl);
			window.location = url;
		});
	};

	function extractDocumentName(docUrl) {
		var spesificSymbols = ['[', ']'];
		var parts = meetDocPattern.exec(docUrl);
		if (!parts) {
			return null;
		}

		var name = parts[0];
		spesificSymbols.forEach(function (item) {
			name = name.replaceAll(item, '');
		});
		return name;
	}

	String.prototype.replaceAll = function (search, replacement) {
		var target = this;
		return target.split(search).join(replacement);
	};

})();