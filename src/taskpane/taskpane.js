/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = async () => {
      const result = await run();
      if (result) {
        document.getElementById("result").innerHTML = `
          <p>First word bold: ${result.isFirstWordBold}</p>
          <p>Second word underlined: ${result.isSecondWordUnderlined}</p>
          <p>Third word font size: ${result.thirdWordFontSize}</p>
        `;
      } else {
        document.getElementById("result").innerHTML = `
          <p> Not enough word/No paragraphs found.</p>
        `;
      }
    };
  }
});

export async function run() {
  try {
    const result = await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");

      await context.sync();

      if (paragraphs.items.length === 0) {
        console.log("No paragraphs found.");
        return null;
      }

      const firstParagraph = paragraphs.items[0];
      const textRanges = firstParagraph.getTextRanges([" "], true);
      textRanges.load("items");

      await context.sync();

      if (textRanges.items.length < 3) {
        console.log("Not enough words in the first paragraph.");
        return null;
      }

      const firstRange = textRanges.items[0];
      const secondRange = textRanges.items[1];
      const thirdRange = textRanges.items[2];

      firstRange.load("font/bold");
      secondRange.load("font/underline");
      thirdRange.load("font/size");

      await context.sync();

      const isFirstWordBold = firstRange.font.bold;
      const isSecondWordUnderlined = secondRange.font.underline !== "None";
      const thirdWordFontSize = thirdRange.font.size;

      console.log("First word bold:", isFirstWordBold);
      console.log("Second word underlined:", isSecondWordUnderlined);
      console.log("Third word font size:", thirdWordFontSize);

      return {
        isFirstWordBold,
        isSecondWordUnderlined,
        thirdWordFontSize,
      };
    });

    return result;
  } catch (error) {
    console.error("Error in run function:", error);
    return null;
  }
}