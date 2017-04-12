/// <reference path="../App.js" />

(function () {
    "use strict";

	var meetXUrl = "http://localhost:8080/officejs-app/App/Home.html";
	Office.initialize = function (reason) {

		window.location = meetXUrl;
	};
})();