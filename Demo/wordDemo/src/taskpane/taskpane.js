/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Add event listeners for buttons
    document.getElementById("run").addEventListener("click", run);
    document.getElementById("insert-date").addEventListener("click", insertDate);
    document.getElementById("get-selected").addEventListener("click", getSelectedText);
    document.getElementById("insert-table").addEventListener("click", insertTable);
    document.getElementById("search-replace").addEventListener("click", searchAndReplace);
    document.getElementById("doc-stats").addEventListener("click", getDocumentStats);
    document.getElementById("insert-link").addEventListener("click", insertLink);
    document.getElementById("change-font").addEventListener("click", changeFont);
    document.getElementById("insert-list").addEventListener("click", insertList);
  }
});

function showMessage(message, isError = false) {
  const messageLabel = document.getElementById("item-subject");
  if (messageLabel) {
    messageLabel.textContent = message;
    messageLabel.style.color = isError ? "#a4262c" : "";
  }

  if (isError) {
    console.error(message);
  } else {
    console.log(message);
  }
}

function sanitizeForHtml(rawText) {
  return String(rawText)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
    // After the paragraph is inserted, queue a command to set its formatting.
    paragraph.load("font,color,alignment,bold");

    await context.sync();
    
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

export async function insertDate() {
  return Word.run(async (context) => {
    // Get the current date
    const currentDate = new Date();
    const dateString = currentDate.toLocaleDateString();

    // Insert the date at the cursor position or end of document
    const range = context.document.getSelection();
    range.insertText(dateString, Word.InsertLocation.replace);

    await context.sync();
  });
}

export async function getSelectedText() {
  console.log("getSelectedText called");
  try {
    return Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load("text");
      await context.sync();

      const selectedText = range.text;
      if (selectedText.trim() === "") {
        showMessage("No text selected. Please select some text first.");
        alert("No text selected. Please select some text first.");
      } else {
        showMessage(`Selected text: "${selectedText}"`);
        alert(`Selected text: "${selectedText}"`);
      }
    });
  } catch (error) {
    console.error("Error in getSelectedText:", error);
    showMessage("Error getting selected text: " + error.message, true);
  }
}

export async function insertTable() {
  console.log("insertTable called");
  try {
    return Word.run(async (context) => {
      // Create a 3x3 table
      const table = context.document.body.insertTable(3, 3, Word.InsertLocation.end);
      
      // Fill the table with sample data
      table.getCell(0, 0).body.insertText("Header 1", Word.InsertLocation.replace);
      table.getCell(0, 1).body.insertText("Header 2", Word.InsertLocation.replace);
      table.getCell(0, 2).body.insertText("Header 3", Word.InsertLocation.replace);
      
      table.getCell(1, 0).body.insertText("Row 1, Col 1", Word.InsertLocation.replace);
      table.getCell(1, 1).body.insertText("Row 1, Col 2", Word.InsertLocation.replace);
      table.getCell(1, 2).body.insertText("Row 1, Col 3", Word.InsertLocation.replace);
      
      table.getCell(2, 0).body.insertText("Row 2, Col 1", Word.InsertLocation.replace);
      table.getCell(2, 1).body.insertText("Row 2, Col 2", Word.InsertLocation.replace);
      table.getCell(2, 2).body.insertText("Row 2, Col 3", Word.InsertLocation.replace);

      await context.sync();
    });
  } catch (error) {
    console.error("Error in insertTable:", error);
    alert("Error inserting table: " + error.message);
  }
}

export async function searchAndReplace() {
  console.log("searchAndReplace called");
  try {
    return Word.run(async (context) => {
      // Prompt user for search and replace text
      const searchText = prompt("Enter text to search for:");
      if (!searchText) return;

      const replaceText = prompt("Enter replacement text:");
      if (replaceText === null) return; // Allow empty string but not cancel

      // Search for the text
      const searchResults = context.document.body.search(searchText);
      context.load(searchResults, "items");
      await context.sync();

      if (searchResults.items.length > 0) {
        // Replace all occurrences
        for (let i = 0; i < searchResults.items.length; i++) {
          searchResults.items[i].insertText(replaceText, Word.InsertLocation.replace);
        }
        await context.sync();
        showMessage(`Replaced ${searchResults.items.length} occurrence(s) of "${searchText}" with "${replaceText}"`);
        alert(`Replaced ${searchResults.items.length} occurrence(s) of "${searchText}" with "${replaceText}"`);
      } else {
        showMessage(`No "${searchText}" found in the document`);
        alert(`No "${searchText}" found in the document`);
      }
    });
  } catch (error) {
    console.error("Error in searchAndReplace:", error);
    showMessage("Error in search and replace: " + error.message, true);
  }
}

export async function getDocumentStats() {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const normalized = body.text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    const words = normalized.match(/\S+/g) || [];
    const wordCount = words.length;
    const charCount = normalized.replace(/\n/g, "").length;
    const paragraphCount = normalized.trim().length === 0 ? 0 : normalized.split(/\n+/).filter(para => para.trim().length > 0).length;

    const statsMessage = `Document Statistics — Words: ${wordCount}, Characters: ${charCount}, Paragraphs: ${paragraphCount}`;
    showMessage(statsMessage);
    alert(statsMessage);
  });
}

export async function insertLink() {
  console.log("insertLink called");
  try {
    return Word.run(async (context) => {
      const linkText = prompt("Enter link text:");
      if (linkText === null) return;

      const linkUrl = prompt("Enter URL:");
      if (!linkUrl) return;

      let normalizedUrl = linkUrl.trim();
      if (!/^https?:\/\//i.test(normalizedUrl) && !/^mailto:/i.test(normalizedUrl)) {
        normalizedUrl = `https://${normalizedUrl}`;
      }

      const range = context.document.getSelection();
      range.load("text");
      await context.sync();

      const displayText = linkText.trim().length > 0 ? linkText.trim() : normalizedUrl;
      const sanitizedText = sanitizeForHtml(displayText);
      const sanitizedUrl = sanitizeForHtml(normalizedUrl);

      if (range.text.trim().length === 0) {
        range.insertHtml(`<a href="${sanitizedUrl}">${sanitizedText}</a>`, Word.InsertLocation.replace);
      } else {
        range.insertHyperlink(normalizedUrl, displayText, Word.InsertLocation.replace);
      }

      await context.sync();
      const linkMsg = `Inserted link: ${displayText} (${normalizedUrl})`;
      showMessage(linkMsg);
      alert(linkMsg);
    });
  } catch (error) {
    console.error("Error in insertLink:", error);
    showMessage("Error inserting link: " + error.message, true);
  }
}

export async function changeFont() {
  console.log("changeFont called");
  try {
    return Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load("text");
      await context.sync();

      if (range.text.trim() === "") {
        showMessage("Please select some text first to change its font.");
        return;
      }

      const fontSize = prompt("Enter font size (e.g., 12, 14, 16):", "12");
      if (fontSize === null) return;

      const parsedFontSize = Number(fontSize);
      if (!Number.isFinite(parsedFontSize) || parsedFontSize <= 0) {
        showMessage("Please provide a valid positive number for the font size.", true);
        return;
      }

      const fontColor = prompt("Enter font color (e.g., red, blue, #FF0000):", "black");
      if (fontColor === null) return;

      const makeBold = confirm("Make text bold?");

      // Apply font changes
      range.font.size = parsedFontSize;
      if (fontColor.trim().length > 0) {
        range.font.color = fontColor.trim();
      }
      range.font.bold = makeBold;

      await context.sync();
      const fontMsg = `Font changed to ${parsedFontSize}pt${fontColor.trim().length > 0 ? `, color ${fontColor.trim()}` : ""}${makeBold ? ", bold" : ""}`;
      showMessage(fontMsg);
      alert(fontMsg);
    });
  } catch (error) {
    console.error("Error in changeFont:", error);
    showMessage("Error changing font: " + error.message, true);
  }
}

export async function insertList() {
  console.log("insertList called");
  try {
    return Word.run(async (context) => {
      const listType = confirm("Create numbered list? (Cancel for bulleted list)");
      const listItems = prompt("Enter list items (separate with commas):", "Item 1, Item 2, Item 3");
      if (!listItems) return;

      const items = listItems.split(',').map(item => item.trim()).filter(item => item.length > 0);

      if (items.length === 0) {
        showMessage("No valid items entered.");
        return;
      }

      const range = context.document.getSelection();
      const listHtml = items.map((item) => `<li>${sanitizeForHtml(item)}</li>`).join("");
      const wrappedList = listType ? `<ol>${listHtml}</ol>` : `<ul>${listHtml}</ul>`;

      range.insertHtml(wrappedList, Word.InsertLocation.replace);
      await context.sync();
      const listMsg = `${listType ? 'Numbered' : 'Bulleted'} list inserted successfully!`;
      showMessage(listMsg);
      alert(listMsg);
    });
  } catch (error) {
    console.error("Error in insertList:", error);
    showMessage("Error inserting list: " + error.message, true);
  }
}
