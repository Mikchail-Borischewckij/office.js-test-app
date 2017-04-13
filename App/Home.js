/// <reference path="../App.js" />

(function () {
	"use strict";


	var officeApps = [
		{
			name: "WORD"
		},
		{
			name: "Excel"
		}
	];


	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {


			var app = getOfficeApp();

			var setLinkFunc;


			if (app && app.name === "WORD") {
				setLinkFunc = setWordLinkToSelection;
				setTextToWordDocument();
			} else {
				setLinkFunc = setExcelLinkToSelection;
			}

			$("#link").click(function () {
				setLinkFunc(app);
			});

		});
	};

	function setTextToWordDocument() {
		Word.run(function (context) {
			var thisDocument = context.document;
			var docName = getParameterByName("docName", window.location);
			var textSample = 'We insert this text using Word object model. Document Name: ' + docName;

			// Create a range proxy object for the current selection.
			var range = context.document.getSelection();

			// Queue a commmand to insert text at the end of the selection.
			range.insertText(textSample, Word.InsertLocation.end);

			// Synchronize the document state by executing the queued commands, 
			// and return a promise to indicate task completion.
			return context.sync().then(function () {
				console.log('Inserted the text at the end of the selection.');
			});
		})
			.catch(function (error) {
				console.log('Error: ' + JSON.stringify(error));
				if (error instanceof OfficeExtension.Error) {
					console.log('Debug info: ' + JSON.stringify(error.debugInfo));
				}
			});
	}

	function setWordLinkToSelection() {
		Word.run(function (context) {
			// Insert your code here. For example:
			context.document.getSelection().hyperlink = "my custom link";

			return context.sync();
		});
	}

	function setExcelLinkToSelection() {
		Excel.run(function (ctx) {
			var selectedRange = ctx.workbook.getSelectedRange();
			selectedRange.formulas = [['=HYPERLINK("http://www.bing.com","test")']];;
			return ctx.sync();
		});
	}

	function getParameterByName(name, url) {
		if (!url) {
			url = window.location.href;
		}
		name = name.replace(/[\[\]]/g, "\\$&");
		var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
			results = regex.exec(url);
		if (!results) return null;
		if (!results[2]) return '';
		return decodeURIComponent(results[2].replace(/\+/g, " "));
	}

	function getOfficeApp() {
		if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
			return officeApps[0];
		}
		if (Office.context.requirements.isSetSupported('ExcelApi', 1.1)) {
			return officeApps[1];
		}
		return undefined;
	}

})();