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
		document.getElementById("run-button").onclick = runData;
	}
});

function runData() {
	Excel.run(async function (context) {
		
		return context.sync().then(function () { })
	})
		.catch(function (error) {
			console.log("Error: " + error);
			// if (error instanceof OfficeExtension.Error) {
			//     console.log("Debug info: " + JSON.stringify(error.debugInfo));
			// }
		});
}