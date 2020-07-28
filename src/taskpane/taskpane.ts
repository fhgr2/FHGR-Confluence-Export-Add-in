/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async context => {
	//
	try {
	document.getElementById("if").style.fontWeight= "bold";
	
    let identifyingFeaturesCCs = context.document.contentControls.getByTag("identifyingFeatures");
    identifyingFeaturesCCs.load("items");
    await context.sync();
    for (let i = 0; i < identifyingFeaturesCCs.items.length; i++)
    {
      let identifyingFeaturesCC = identifyingFeaturesCCs.items[i];
      let ccs = identifyingFeaturesCC.contentControls;
      ccs.load("items");
      await context.sync();
      for (let j = 0; j < ccs.items.length; j++) {
        let cc = ccs.items[j];
        var text = cc.text;
        cc.insertHtml(text, 'Replace');        
      }
    }
	document.getElementById("if").style.fontWeight= "normal";
	await context.sync();
	
	document.getElementById("bf").style.fontWeight= "bold";
    // https://stackoverflow.com/questions/48371446/find-bold-words-in-selection-using-office-addin-javascript-api
    let tsr = context.document.body.getRange("Whole").getTextRanges([" "], true);
    // console.log(tsr);
    // tsr.load("font/bold, font/italic, text, style");
	document.getElementById("log").innerHTML = "a";
    tsr.load("items");
	// Absturz auf der nächsten Zeile! KOmischerweise tut das im 
	await context.sync();
	document.getElementById("log").innerHTML = "b";
	// tsr.load("items");
	
	//await context.sync();
	document.getElementById("log").innerHTML = "c";
	
    for (let i = 0; i < tsr.items.length; i++)
    {
		document.getElementById("log").innerHTML = "d";
		let word = tsr.items[i];
		
      if (word.font.bold) {
        // console.log(word.text);
        // word.font.bold = false;
        // word.insertText("kk", "Replace")
        // Doesn't work because it sets the style of the whole paragraph
        word.style = "Intensive Hervorhebung";
      }
	}
	
	document.getElementById("bf").style.fontWeight= "normal";
	
    await context.sync();
	} catch(e) {
		// document.getElementById("log").innerHTML = e.message;
		// document.getElementById("log").innerHTML = e.stack;
	}
	
  });
}
