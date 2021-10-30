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

import { convertArraysToObjectsInArray, convertObjectsToArraysInArray } from '../realizedSections/realizedSections.js';

Office.onReady((info) => {
	if (info.host === Office.HostType.Excel) {
		document.getElementById("sideload-msg").style.display = "none";
		document.getElementById("app-body").style.display = "flex";
		document.getElementById("run-button").onclick = runData;
	}
});
var exportedPlanBody;
var resultData = [];
const sheet_name = "exported_plan";
function runData() {
	Excel.run(function (context) {
		const infoOneCol = document.getElementById("info-one").value;
		const infoTwoCol = document.getElementById("info-two").value;
		const startDateCol = document.getElementById("start-date").value;
		const endDateCol = document.getElementById("end-date").value;
		const amountCol = document.getElementById("amount").value;

		const planTable = context.workbook.tables.getItem(document.getElementById("table-name").value);
		const workDaysTable = context.workbook.tables.getItem("workDays");

		const planHeader = planTable.getHeaderRowRange().load("values");
		const planBodyRange = planTable.getDataBodyRange().load("values");
		const workDaysHeader = workDaysTable.getHeaderRowRange().load("values");
		const workDaysBodyRange = workDaysTable.getDataBodyRange().load("values");

		return context.sync().then(function () {
			var planHead = planHeader.values[0];
			var workDaysHead = workDaysHeader.values[0];
			var planData = planBodyRange.values; // * arrays in array
			var workDaysData = workDaysBodyRange.values; // * arrays in array
			planData = convertArraysToObjectsInArray(planData, planHead) // * objects in array
				.filter(i => i[startDateCol] > 0 && i[endDateCol] > 0 && i[amountCol] > 0)
				.map(i => {
					return {
						info1: i[infoOneCol],
						info2: i[infoTwoCol],
						startDate: i[startDateCol],
						endDate: i[endDateCol],
						amount: i[amountCol],
					}
				})
			workDaysData = convertArraysToObjectsInArray(workDaysData, workDaysHead);
			// console.log('workDaysData :', workDaysData);
			// console.log('planData :', planData);

			for (let row of planData) {
				let days = [];
				// console.log(row.startDate);
				for (let i = row.startDate; i <= row.endDate; i++) {
					if (workDaysData.filter(element => element.date == i)[0].work_day == 1) {
						days.push(i);
					};
				}
				let amountPerDay = Math.round(row.amount / days.length * 10) / 10;
				for (let day of days) {
					resultData.push({
						info1: row.info1,
						info2: row.info2,
						date: day,
						amount: amountPerDay
					})
				}
			};
			resultData = convertObjectsToArraysInArray(resultData);
			for (let i of resultData[1]) {
				i.push("=+XLOOKUP(RIGHT([@info1],2),imalatlar[Kod_str],imalatlar[İsim])");
			}
			resultData[0].push("imalat");
			var sheets = context.workbook.worksheets;
			sheets.load("items/name");
			return context.sync().then(function () {
				var delete_flag = false;
				var exportedPlan;
				var new_sheet;
				for (let i of sheets.items) {
					if (i.name == sheet_name) {
						delete_flag = true;
					}
				}
				if (delete_flag) {
					exportedPlan = context.workbook.tables.getItem("exportedPlan");
					exportedPlan.getDataBodyRange().clear();
					exportedPlan.clearFilters();
					new_sheet = context.workbook.worksheets.getItem(sheet_name);
				} else {
					new_sheet = sheets.add(sheet_name);
					exportedPlan = new_sheet.tables.add("A1:E1", true /*hasHeaders*/);
					exportedPlan.name = "exportedPlan";
					exportedPlan.getHeaderRowRange().values = [resultData[0]];
				}
				new_sheet.activate();
				exportedPlan.rows.add(0, resultData[1]);
				exportedPlan.columns.getItemAt(2).getDataBodyRange().numberFormat = "dd.mm.yyyy"
				return context.sync().then(function () {
					new_sheet.getUsedRange().format.autofitColumns()
					location.reload();

					// * ikinci üretilecek sayfaya girişiyorum. önce data. 
					
				})
			})
		})
	}).catch(function (error) {
		console.log("Error: " + error.debugInfo);
		// if (error instanceof OfficeExtension.Error) {
		//     console.log("Debug info: " + JSON.stringify(error.debugInfo));
		// }
	});
}