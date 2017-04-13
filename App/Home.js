/// <reference path="../App.js" />

(function () {
	"use strict";

	var meetDocPattern = /(\[|\()[A-Za-z0-9_=]+(\]|\))/g;
	var meetXUrl = "http://localhost:8080/officejs-app/App/Home.html";
	var documentName;

	Office.initialize = function (reason) {
		$(document).ready(function () {
			var docUrl = Office.context.document.url;
			
			if (docUrl) {
				documentName = extractDocumentName(docUrl);
				if (documentName) {
					var info = getDocumentInfo(documentName);
					window.location = meetXUrl + "?docName=" + docUrl;
					//var url = "https://ec.boardvantage.com/services/officelink/" +
					//	info.token +
					//	"/?action=officeAgendaLink&reqApp=wtask&folder=&locale=";
					//window.location = url;
				}
			}
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

	function getDocumentInfo(parametersString) {
		var decodedString = atob(parametersString);

		var pathSplit = decodedString.split('*');
		var argLength = pathSplit.length + 1;
		if (argLength >= 3) {
			var ecxSessionId = pathSplit[1];
			var serverName = pathSplit[2];
			var lbcookie = pathSplit[3];
			var combinedStr = "B=" + lbcookie + ";" + "j=j" + ";" + "E=" + ecxSessionId;
			var encodedToken = btoa(combinedStr);
			return { server: serverName, token: encodedToken, ecxSessionId: ecxSessionId };
		}
		return null;
	}

})();