/* eslint-disable no-unused-vars */
/* eslint-disable prettier/prettier */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest 
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global console, document, Excel, Office */

Office.onReady((info) => {
	if (info.host === Office.HostType.Excel) {
		document.getElementById("sideload-msg").style.display = "none";
		document.getElementById("app-body").style.display = "flex";
		document.getElementById("pull-button").onclick = pullData;
	}
});

// loadJSON method to open the JSON file.
function loadJSON(path, success, error) {
	var xhr = new XMLHttpRequest();
	xhr.onreadystatechange = function () {
		if (xhr.readyState === 4) {
			if (xhr.status === 200) {
				success(JSON.parse(xhr.responseText));
			}
			else {
				error(xhr);
			}
		}
	};
	xhr.open('GET', path, true);
	xhr.send();
}

// fetch('http://95.70.184.238:3001/api/wbsApi/j19dc2o4o3kpsyj74pyo')
// 	.then(function (response) {
// 		return response.json();
// 	})
// 	.then(function (myJson) {
// 		console.log(JSON.stringify(myJson));
// 	});

var receivedData;
var dummy;



// http://jsonplaceholder.typicode.com/albums
// https://jsonplaceholder.typicode.com/posts
// http://95.70.184.238:3001/api/wbsApi/j19dc2o4o3kpsyj74pyo
function pullData() {
	Excel.run(async function (context) {
		var url = document.getElementById("url-box").value;
		var sheets = context.workbook.worksheets;
		var sheet = sheets.add("pulled data");
		sheet.activate();

		// fetch(url)
		// 	.then(response => response.json())
		// 	.then(responseJson => {
		// 		let remoteData = responseJson;
		// 		// console.log(data);
		// 		var data = [];
		// 		var row;
		// 		for (let i of remoteData) {
		// 			row = Object.values(i);
		// 			data.push(row);
		// 		}
		// 		var columns = Object.keys(responseJson[0]);
		// 		var table = sheet.tables.add("A1:D1", true /*hasHeaders*/);
		// 		table.name = "pulled_table";
		// 		table.getHeaderRowRange().values = [columns];
		// 		table.rows.add(null /*add rows to the end of the table*/,
		// 			data);
		// 	});

		// eslint-disable-next-line no-undef
		const remoteData_ = await fetch(url);
		const remoteData = await remoteData_.json();
		var data = [];
		var row;
		for (let i of remoteData) {
			row = Object.values(i);
			data.push(row);
		}
		// const alphabet = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];
		var columns = Object.keys(remoteData[0]);
		var table = sheet.tables.add("A1:C1", true /*hasHeaders*/); // TODO alfabeye göre buradaki c harfini programatik yapmak lazım. albüm apisine göre C olmalı. kolon sayısına bağlı bişey. 
		table.name = "pulled_table";
		table.getHeaderRowRange().values = [columns];
		table.rows.add(null /*add rows to the end of the table*/,
			data);

		console.log(data)
		return context.sync().then(function () { })
	})
		.catch(function (error) {
			console.log("Error: " + error);
			// if (error instanceof OfficeExtension.Error) {
			//     console.log("Debug info: " + JSON.stringify(error.debugInfo));
			// }
		});
}