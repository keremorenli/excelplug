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
    document.getElementById("calculate").onclick = calculate;
    document.getElementById("reload").onclick = reload;
  }
});

// ACTIVITY TYPE DATA
var ActivityTypeDictionary = {
  "34": "Faz-1 Betonu Dökülmesi",
  "35": "Faz-2 Betonu Dökülmesi",
  "08": "Rayların Hatta Dağıtılması",
  "38": "Panelrayların Serilmesi",
  "41": "Grout Dökülmesi",
  "39": "Panelrayların Kot ve Eksen Ayarının Yapılması",
  "09": "Rayların Hatta Montajı",
  "43": "Plinth Beton Dökümü"
};


// FUNCTIONS
function populateList(listId, arr) {
  const list = document.getElementById(listId);
  for (let i = 0; i <= arr.length - 1; i++) {
    const li = document.createElement("li");
    li.innerHTML = arr[i];
    list.appendChild(li);
  }
}

function findIndex(arr, value) {
  for (let i = 0; i < arr.length; i++) {
    if (arr[i] == value) {
      return i
    }
  }
  return -1
}

function findActivityTypes(columns_arr, data) {
  let key_column_name = "WBSADI"; // ! bu kolon adına göre hareket edecek, önemli.
  let index = findIndex(columns_arr, key_column_name);
  const activityTypes = [];
  for (let i = 0; i < data.length; i++) {
    var wbsValue = data[i][index];
    if (wbsValue == key_column_name) {
      continue; // data'nın ilk satırı kolon başlıklarıysa elemek için.
    }
    var activityCode = wbsValue.slice(wbsValue.length - 2, wbsValue.length);
    if (findIndex(activityTypes, activityCode) == -1) {
      activityTypes.push(activityCode);
    }
  }
  return activityTypes;
}

function getTheLastX(str, x) {
   return str.slice(str.length - x, str.length);
}

// CALCULATE BUTTON
function calculate() {
  Excel.run(function (context) {
    var tableName = "data";
    var table = context.workbook.tables.getItem(tableName);
    var tableRange = table.getRange();
    tableRange.load(["rowCount", "address", "values"]);
    var columnsRow = tableRange.getRow(0);
    columnsRow.load("values");;

    return context.sync().then(
      function () {
        document.getElementById("tableRange").innerHTML = tableRange.address;

        var countOfRows = tableRange.rowCount;
        document.getElementById("number_of_rows").innerHTML = countOfRows;

        var columns = columnsRow.values[0];
        document.getElementById("columns").innerHTML = "";
        populateList("columns", columns);

        var data = tableRange.values;
        var activityTypes = findActivityTypes(columnsRow.values[0], data)
        var tableHTML = document.getElementById("results_table");
        var x = 1;
        for (let i of activityTypes) {
            let row = tableHTML.insertRow(x);
            row.insertCell(0).innerHTML = i;
            row.insertCell(1).innerHTML = ActivityTypeDictionary[i];
            
            // * realized quantity calculation
            let quantity = 0;
            for (let d of data) {
              // ! kolon isimleri önemli
              if (getTheLastX(d[findIndex(columns,"WBSADI")],2) == i) {
                quantity = quantity + Number(d[findIndex(columns,"ILERLEME")]);
              }
            };
            row.insertCell(2).innerHTML = quantity;
            
            // * working hours calculation
            let workingHours = 0;
            for (let d of data) {
              // ! kolon isimleri önemli
              if (getTheLastX(d[findIndex(columns,"WBSADI")],2) == i
               && d[findIndex(columns,"ISDETAYI")] == "Çalışma"
              ) {
                workingHours = workingHours + Number(d[findIndex(columns,"CALISMASAATI")]);
              }
            };
            row.insertCell(3).innerHTML = workingHours;
            
            row.insertCell(4).innerHTML = Math.round(quantity / workingHours * 10);

            x++;
        };

      }
    ).then(context.sync);
  })
    .catch(function (error) {
      console.log("Error: " + error);
      // if (error instanceof OfficeExtension.Error) {
      //     console.log("Debug info: " + JSON.stringify(error.debugInfo));
      // }
    });
}

// var table = document.getElementById("results_table");
// var row = table.insertRow(2);
// var cell1 = row.insertCell(0);
// var cell2 = row.insertCell(1);
// cell1.innerHTML = "new cell";
// cell2.innerHTML = "new cell";


// CALCULATE BUTTON
function reload() {
  Excel.run(function (context) {

    // eslint-disable-next-line no-undef
    location.reload();

    return context.sync()
  })
    .catch(function (error) {
      console.log("Error: " + error);
      // if (error instanceof OfficeExtension.Error) {
      //     console.log("Debug info: " + JSON.stringify(error.debugInfo));
      // }
    });
}