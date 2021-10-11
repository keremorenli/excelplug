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
  }
});

async function calculate() {
  try {
    await Excel.run(async (context) => {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      var range = sheet.getRange("A1:f20");
      range.format.fill.color = "red";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
