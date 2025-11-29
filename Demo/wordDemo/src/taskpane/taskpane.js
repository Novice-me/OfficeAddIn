/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    const buttonBindings = [
      ["run", run],
      ["insert-date", insertDate],
      ["get-selected", getSelectedText],
      ["insert-table", insertTable],
      ["doc-stats", getDocumentStats]
    ];

    const formBindings = [
      ["search-replace-form", searchAndReplace],
      ["insert-link-form", insertLink],
      ["change-font-form", changeFont],
      ["insert-list-form", insertList]
    ];

    buttonBindings.forEach(([elementId, handler]) => bindButton(elementId, handler));
    formBindings.forEach(([formId, handler]) => bindForm(formId, handler));

    showMessage("请选择下方一个功能区并按照提示操作。");
    console.log("Office add-in initialized for Word. Event listeners attached.");
  } else {
    console.log("Add-in loaded in non-Word host:", info.host);
  }
});

function bindButton(elementId, handler) {
  const element = document.getElementById(elementId);
  if (!element) {
    console.warn(`Button with id "${elementId}" not found.`);
    return;
  }

  element.addEventListener("click", async (event) => {
    event.preventDefault();
    try {
      await handler();
    } catch (error) {
      console.error(`Error executing handler for ${elementId}:`, error);
      showMessage(`执行操作时出错：${error.message}`, true);
    }
  });
}

function bindForm(formId, handler) {
  const form = document.getElementById(formId);
  if (!form) {
    console.warn(`Form with id "${formId}" not found.`);
    return;
  }

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    try {
      await handler();
    } catch (error) {
      console.error(`Error executing handler for ${formId}:`, error);
      showMessage(`执行操作时出错：${error.message}`, true);
    }
  });
}

function showMessage(message, isError = false) {
  const messageContainer = document.getElementById("status-message");
  if (messageContainer) {
    messageContainer.textContent = message || "";
    messageContainer.classList.remove("error", "success");
    if (message && message.trim()) {
      messageContainer.classList.add(isError ? "error" : "success");
    }
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
    showMessage("已在文档末尾插入带格式的 \"Hello World\" 段落。");
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
    showMessage(`已插入日期 ${dateString}。`);
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
        showMessage("未检测到选中文本，请先在 Word 中选择内容。", true);
      } else {
        showMessage(`选中文本: "${selectedText}"`);
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
      showMessage("已在文档末尾插入 3x3 示例表格。");
    });
  } catch (error) {
    console.error("Error in insertTable:", error);
    showMessage("插入示例表格时出错: " + error.message, true);
  }
}

export async function searchAndReplace() {
  console.log("searchAndReplace called");
  const searchInput = document.getElementById("search-text-input");
  const replaceInput = document.getElementById("replace-text-input");

  if (!searchInput) {
    showMessage("未找到查找输入框，请重新加载任务窗格。", true);
    return;
  }

  const searchText = searchInput.value.trim();
  const replaceText = replaceInput ? replaceInput.value : "";

  if (searchText.length === 0) {
    showMessage("请输入要查找的内容。", true);
    searchInput.focus();
    return;
  }

  try {
    return Word.run(async (context) => {
      const searchResults = context.document.body.search(searchText, { matchWildcards: false });
      context.load(searchResults, "items");
      await context.sync();

      if (searchResults.items.length > 0) {
        // Replace all occurrences
        searchResults.items.forEach((item) => {
          item.insertText(replaceText, Word.InsertLocation.replace);
        });
        await context.sync();
        const replacementDescription = replaceText.trim().length > 0 ? `"${replaceText}"` : "空字符串";
        showMessage(`已将 ${searchResults.items.length} 处 "${searchText}" 替换为 ${replacementDescription}。`);
        if (replaceInput) {
          replaceInput.focus();
        }
      } else {
        showMessage(`未在文档中找到 "${searchText}"，请调整检索内容。`, true);
        searchInput.focus();
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

    const statsMessage = `文档统计 / Document Statistics — 词数: ${wordCount}, 字符数: ${charCount}, 段落: ${paragraphCount}`;
    showMessage(statsMessage);
  });
}

export async function insertLink() {
  console.log("insertLink called");
  const linkTextInput = document.getElementById("link-text-input");
  const linkUrlInput = document.getElementById("link-url-input");

  if (!linkUrlInput) {
    showMessage("未找到链接输入控件，请重新加载任务窗格。", true);
    return;
  }

  const rawLinkText = linkTextInput ? linkTextInput.value.trim() : "";
  const rawUrl = linkUrlInput.value.trim();

  if (rawUrl.length === 0) {
    showMessage("请输入有效的链接地址。", true);
    linkUrlInput.focus();
    return;
  }

  try {
    return Word.run(async (context) => {
      let normalizedUrl = rawUrl;
      if (!/^https?:\/\//i.test(normalizedUrl) && !/^mailto:/i.test(normalizedUrl)) {
        normalizedUrl = `https://${normalizedUrl}`;
      }

      const range = context.document.getSelection();
      range.load("text");
      await context.sync();

      const displayText = rawLinkText.length > 0 ? rawLinkText : normalizedUrl;
      const sanitizedText = sanitizeForHtml(displayText);
      const sanitizedUrl = sanitizeForHtml(normalizedUrl);

      if (range.text.trim().length === 0) {
        range.insertHtml(`<a href="${sanitizedUrl}">${sanitizedText}</a>`, Word.InsertLocation.replace);
      } else {
        range.insertHyperlink(normalizedUrl, displayText, Word.InsertLocation.replace);
      }

      await context.sync();
      const linkMsg = `已插入链接：${displayText} (${normalizedUrl})`;
      showMessage(linkMsg);
      if (linkTextInput) {
        linkTextInput.focus();
      }
    });
  } catch (error) {
    console.error("Error in insertLink:", error);
    showMessage("Error inserting link: " + error.message, true);
  }
}

export async function changeFont() {
  console.log("changeFont called");
  const fontSizeInput = document.getElementById("font-size-input");
  const fontColorInput = document.getElementById("font-color-input");
  const boldCheckbox = document.getElementById("font-bold-checkbox");

  if (!fontSizeInput) {
    showMessage("未找到字号输入控件，请重新加载任务窗格。", true);
    return;
  }

  const fontSizeValue = fontSizeInput.value.trim();
  const parsedFontSize = Number(fontSizeValue);

  if (!fontSizeValue || !Number.isFinite(parsedFontSize) || parsedFontSize <= 0) {
    showMessage("请输入大于 0 的字号数值。", true);
    fontSizeInput.focus();
    return;
  }

  const fontColorValue = fontColorInput ? fontColorInput.value.trim() : "";
  const makeBold = Boolean(boldCheckbox && boldCheckbox.checked);

  try {
    return Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load("text");
      await context.sync();

      if (range.text.trim() === "") {
        showMessage("请先在文档中选择要设置的文本。", true);
        return;
      }

      // Apply font changes
      range.font.size = parsedFontSize;
      if (fontColorValue.length > 0) {
        range.font.color = fontColorValue;
      }
      range.font.bold = makeBold;

      await context.sync();
      const fontMsg = `已应用字号 ${parsedFontSize}pt${fontColorValue.length > 0 ? `，颜色 ${fontColorValue}` : ""}${makeBold ? "，并加粗" : ""}。`;
      showMessage(fontMsg);
    });
  } catch (error) {
    console.error("Error in changeFont:", error);
    showMessage("Error changing font: " + error.message, true);
  }
}

export async function insertList() {
  console.log("insertList called");
  const listTypeSelect = document.getElementById("list-type-select");
  const listItemsTextarea = document.getElementById("list-items-textarea");

  if (!listItemsTextarea) {
    showMessage("未找到列表输入控件，请重新加载任务窗格。", true);
    return;
  }

  const listType = listTypeSelect ? listTypeSelect.value : "bulleted";
  const rawItems = listItemsTextarea.value || "";
  const items = rawItems
    .split(/\r?\n/)
    .map((item) => item.trim())
    .filter((item) => item.length > 0);

  if (items.length === 0) {
    showMessage("请至少输入一行列表项目。", true);
    listItemsTextarea.focus();
    return;
  }

  try {
    return Word.run(async (context) => {
      const range = context.document.getSelection();
      const listHtml = items.map((item) => `<li>${sanitizeForHtml(item)}</li>`).join("");
      const useNumbered = listType === "numbered";
      const wrappedList = useNumbered ? `<ol>${listHtml}</ol>` : `<ul>${listHtml}</ul>`;

      range.insertHtml(wrappedList, Word.InsertLocation.replace);
      await context.sync();
      const listMsg = `${useNumbered ? "已插入编号列表" : "已插入项目符号列表"}，共 ${items.length} 项。`;
      showMessage(listMsg);
      listItemsTextarea.focus();
    });
  } catch (error) {
    console.error("Error in insertList:", error);
    showMessage("Error inserting list: " + error.message, true);
  }
}
