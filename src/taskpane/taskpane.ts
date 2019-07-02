/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }

});

var myformat = new Intl.NumberFormat('en-US', { 
  minimumIntegerDigits: 1, 
  minimumFractionDigits: 2 
});



export async function run() {
  try {

    const isSupported = Office.context.requirements.isSetSupported("ExcelApi", 1.9);


    if (isSupported) {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();
        let data = await DataProvider.myFunc();
        range.values = data;
        
        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    }
  } catch (error) {
    console.error(error);
  }
}

export class DataProvider {
  static async myFunc(): Promise<number[][]>{
    let x:number[][] = [[9]]; //Edge F12 tool will crash simply by having this line
    return x;
  }
}
