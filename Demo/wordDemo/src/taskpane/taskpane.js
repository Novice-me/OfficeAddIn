/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
    
    // change the paragraph color to blue.
    paragraph.font.color = "blue";

  // change the paragraph alignment to center.
  // Use the correct enum member name: 'centered'
  paragraph.alignment = Word.Alignment.centered;

    // change the paragraph to bold.
    paragraph.font.bold = true;

    await context.sync();
  });
}
