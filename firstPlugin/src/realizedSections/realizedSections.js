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

// print fonksiyonunu console.log yerine kullanmak için yazıyorum. sürekli inspect ile uğraşmak vakit alıyordu. html'in en üstüne koyuyorum gösterilmek yazdırlanı.	  
function print(msg) {
	document.getElementById("print").innerHTML = msg;
};

// elindeki tablo arrayler içinde tutuluyosa header array ile kolon isimlerini vererek objeler içinde tutulmasını sağlayabilirsin.  
export function convertArraysToObjectsInArray(data, header) {
	var result = [];
	for (const row of data) {
		let obj = {}
		for (let i = 0; i < row.length; i++) {
			obj[header[i]] = row[i];
		};
		result.push(obj);
	};
	return result
}

// array içinde objeler ile oluşturulmuş bir tabloyu array içinde arrayler şeklinde düzenler. çıktı olarak iki elementli bi array verir. birincisi kolon başlıklarından oluşan bi array. ikincisi data.
export function convertObjectsToArraysInArray(data) {
	let header = Object.keys(data[0]);
	let arr = [];
	for (let i of data) {
		arr.push(Object.values(i));
	}

	return [header, arr]
}

Office.onReady((info) => {
	if (info.host === Office.HostType.Excel) {
		document.getElementById("sideload-msg").style.display = "none";
		document.getElementById("app-body").style.display = "flex";
		document.getElementById("run-button").onclick = runData;
	}
});
var newRealizedBody;
var sheet_name = "ayiklanmis_gerceklesme";
var resultData = [];
function runData() {
	Excel.run(async function (context) {
		var realizedTableName = document.getElementById("realized-table-name").value;
		var sectionsTableName = document.getElementById("sections-table-name").value;
		var wbsCol = document.getElementById("wbs-column-name").value;
		var dateCol = document.getElementById("date-column-name").value;
		var startCol = document.getElementById("start-column-name").value;
		var endCol = document.getElementById("end-column-name").value;
		var lineCol = document.getElementById("line-column-name").value;
		var idCol = document.getElementById("id-column-name").value;
		var startSecCol = document.getElementById("startsec-column-name").value;
		var endSecCol = document.getElementById("endsec-column-name").value;
		var lineSecCol = document.getElementById("linesec-column-name").value;

		var realizedTable = context.workbook.tables.getItem(realizedTableName);
		var sectionsTable = context.workbook.tables.getItem(sectionsTableName);

		var realizedHeader = realizedTable.getHeaderRowRange().load("values");
		var sectionsHeader = sectionsTable.getHeaderRowRange().load("values");
		var realizedBodyRange = realizedTable.getDataBodyRange().load("values");
		var sectionsBodyRange = sectionsTable.getDataBodyRange().load("values");

		return context.sync().then(function () {
			var realizedHead = realizedHeader.values[0];
			var sectionsHead = sectionsHeader.values[0];
			var realizedData = realizedBodyRange.values; // ! arrays in array
			var sectionsData = sectionsBodyRange.values; // ! arrays in array
			realizedData = convertArraysToObjectsInArray(realizedData, realizedHead); // ! objects in array
			sectionsData = convertArraysToObjectsInArray(sectionsData, sectionsHead).filter(i => i[startSecCol] >= 0 && i[endSecCol] > 0); // ! objects in array
			// var sectionsStartKm = sectionsData.map(i => i[startSecCol]); // array
			// var sectionsEndKm = sectionsData.map(i => i[endSecCol]); // array
			// TODO eğer bi section son km'si X ise başka bi section ilk km'si de X olmalı. hepsi birbirine bağlanmalı. depo öyle olmadığı için patlıyodu. bi kontrol atabiliriz buraya
			realizedData = realizedData.filter(i => i[startCol] >= 0 && i[endCol] > 0 && i[lineCol] !== "") // ! sadece km datası olanların gerekli kolonlarını çekelim.
				.map(i => {
					return {
						imalat: i[wbsCol],
						tarih: i[dateCol],
						bas_km: i[startCol],
						bit_km: i[endCol],
						hat: i[lineCol],
						id: i["ID"]
					}
				});
			// var resultData2 = [];
			var flag;
			for (let row of realizedData) {
				flag = false;
				for (let sec of sectionsData) {
					let control1 = sec[startSecCol] <= row.bas_km;
					let control2 = sec[startSecCol] <= row.bit_km;
					let control3 = sec[endSecCol] >= row.bas_km;
					let control4 = sec[endSecCol] >= row.bit_km;
					let control5 = sec[lineSecCol] == row.hat;
					if (control1 && control2 && control3 && control4 && control5) {
						flag = true;
						resultData.push({
							imalat: row.imalat,
							tarih: row.tarih,
							bas_km: row.bas_km,
							bit_km: row.bit_km,
							bolge: sec[idCol],
							id: row.id
						})
						// continue;
					}
					// else {

				}
				// console.log('flag :', flag);
				if (!flag) {
					let smallerKm = Math.min(...[row.bas_km]);
					let largerKm = Math.max(...[row.bit_km]);
					for (let sec of sectionsData) {
						let control1 = sec[startSecCol] <= smallerKm;
						// console.log('smallerKm :', smallerKm);
						// console.log('startSecCol :', startSecCol);
						let control2 = sec[endSecCol] >= smallerKm;
						let control3 = sec[lineSecCol] == row.hat;
						// console.log(control1);
						// console.log(control2);
						// console.log(control3); 
						// console.log("-------------------")
						if (control1 && control2 && control3) {
							var sectionKMsofSmallerKM = [sec[startSecCol], sec[endSecCol]];
							var sectionofSmallerKM = sec[idCol];
						}
						let control4 = sec[startSecCol] <= largerKm;
						let control5 = sec[endSecCol] >= largerKm;
						// console.log(control4);
						// console.log(control5);
						// console.log(control3);
						// console.log("-------------------")
						if (control4 && control5 && control3) {
							var sectionKMsofLargerKM = [sec[startSecCol], sec[endSecCol]];
							var sectionofLargerKM = sec[idCol];
						}
					}

					// console.log(' :', sectionKMsofSmallerKM);
					// console.log(' :', sectionKMsofLargerKM);
					resultData.push({
						imalat: row.imalat,
						tarih: row.tarih,
						bas_km: Math.min(row.bas_km, row.bit_km),
						bit_km: sectionKMsofSmallerKM[1],
						bolge: sectionofSmallerKM,
						id: row.id
					});
					resultData.push({
						imalat: row.imalat,
						tarih: row.tarih,
						bas_km: Math.max(row.bas_km, row.bit_km),
						bit_km: sectionKMsofLargerKM[1],
						bolge: sectionofLargerKM,
						id: row.id
					});
				}
			}
			// console.log(convertObjectsToArraysInArray(resultData));
			resultData = convertObjectsToArraysInArray(resultData);
			for (let i of resultData[1]) {
				i.push("=+ABS([@[bit_km]]-[@[bas_km]])");
				i.push("=TEXT(RIGHT([@[imalat]],2),\"#\")")
			}
			resultData[0].push("mesafeData");
			resultData[0].push("imalatKodData");
			console.log(resultData);
			// ! verimizi oluşturduk şimdi excelde bir tabloya yazalım.

			// try {
			// 	let sheets = context.workbook.worksheets;
			// 	sheet = sheets.add(sheet_name);
			// 	sheet.load("name, position");

			// } catch {

			// }
			var sheets = context.workbook.worksheets;
			sheets.load("items/name");
			// var sheets = context.workbook.worksheets;
			// var sheet = sheets.add(sheet_name);
			// sheet.load("name, position");

			return context.sync().then(function () {
				var delete_flag = false;
				var new_realized;
				var new_sheet;
				for (let i of sheets.items) {
					if (i.name == sheet_name) {
						delete_flag = true;
					}
				}
				if (delete_flag) {
					new_realized = context.workbook.tables.getItem("new_realized");
					// new_realized.columns.getItem("imalatKod").delete();
					// new_realized.columns.getItem("mesafe").delete();
					new_realized.getDataBodyRange().clear();
					new_sheet = context.workbook.worksheets.getItem(sheet_name);
					// new_realized.rows.add(0, resultData[1]); 
				} else {
					new_sheet = sheets.add(sheet_name);
					new_realized = new_sheet.tables.add("A1:H1", true /*hasHeaders*/);
					new_realized.name = "new_realized";
					new_realized.getHeaderRowRange().values = [resultData[0]];
				}
				new_realized.rows.add(0, resultData[1]);
				new_realized.columns.getItemAt(1).getDataBodyRange().numberFormat = "dd.mm.yyyy"
				new_realized.columns.getItemAt(2).getDataBodyRange().numberFormat = "0+000"
				new_realized.columns.getItemAt(3).getDataBodyRange().numberFormat = "0+000"
				new_sheet.activate();
				// 
				// newRealizedBody = new_realized.getDataBodyRange();
				// newRealizedBody.load("rowCount");
				return context.sync().then(function () {
					// let mesafeData = [];
					// let imalatKodData = [];
					// for (let i = 1; i <= newRealizedBody.rowCount + 1; i++) {
					// 	mesafeData.push(["=+ABS([@[bit_km]]-[@[bas_km]])"]);
					// 	imalatKodData.push(["=TEXT(RIGHT([@[imalat]],2),\"#\")"])
					// }
					// if (delete_flag) {
					// 	new_realized.columns.getItem("mesafe").getDataBodyRange().values = mesafeData;
					// 	new_realized.columns.getItem("imalatKod").getDataBodyRange().values = imalatKodData;
					// } else {
					// 	new_realized.columns.add(-1, mesafeData, "mesafe");
					// 	new_realized.columns.add(-1, imalatKodData, "imalatKod")
					// };
					// new_sheet.getUsedRange().format.autofitColumns();
					// // eslint-disable-next-line no-undef
					location.reload();
				})
			})
		})
	})
		.catch(function (error) {
			console.log("Error: " + error.debugInfo);
			// if (error instanceof OfficeExtension.Error) {
			//     console.log("Debug info: " + JSON.stringify(error.debugInfo));
			// }
		});
}