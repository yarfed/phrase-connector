/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

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
    /**
     * Insert your Word code here
     */
    var text = $("input[name='contact']:checked").val();
    var values = $('input[name="subscribe"]:checked').map(function() {
      return $( this ).val();
    }).get()
    var text1 = values.filter(Boolean).join(", ")||"все пучком";
    if (values[0]=="отделен от тела"){
      text=(text=="На хуе")?"пенис ":'Нос '
    }
    context.document.body.clear()
    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph(text+" "+text1, Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}
