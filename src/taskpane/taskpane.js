/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("translateBtn").onclick = translateAndInsert;
  }
});

export async function translateAndInsert() {
  const text = document.getElementById("inputText").value.trim();
  const resultDiv = document.getElementById("result");

  if (!text) {
    resultDiv.innerText = "⚠️ Please enter English text.";
    return;
  }

  resultDiv.innerText = "⏳ Translating...";
  //console.log("[DEBUG] Start translating:", text);

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const tableName = "TranslationTable";
    let table;

    // 检查翻译表是否存在
    try {
      table = sheet.tables.getItem(tableName);
      table.load("name");
      await context.sync();
    } catch {
      const header = [["编号", "英文", "中文译文", "查询时间"]];
      const range = sheet.getRange("A1:D1");
      range.values = header;
      table = sheet.tables.add(range, true);
      table.name = tableName;
      await context.sync();
    }

    // 调用 Excel TRANSLATE 函数
    const tempCell = sheet.getRange("Z1");
    const formula = `=TRANSLATE("${text}", "en", "zh-chs")`;
    tempCell.formulas = [[formula]];
    await context.sync();
    //console.log("[DEBUG] Formula set:", formula);

    // 等待公式计算结果（轮询等待）
    let translation = "#CALC!";
    const maxTries = 10; // 最多尝试10次（大约10秒）
    for (let i = 0; i < maxTries; i++) {
      await new Promise((resolve) => setTimeout(resolve, 1000)); // 每次等待1秒
      tempCell.load("values");
      await context.sync();

      translation = tempCell.values?.[0]?.[0];
      //console.log(`[DEBUG] Try ${i + 1}: value =`, translation);

      if (translation && translation !== "#CALC!" && translation !== "#NAME?") {
        break;
      }
    }

    // 如果超时
    if (translation === "#CALC!" || translation === "#NAME?") {
      translation = "(翻译超时或函数不支持)";
    }

    const time = new Date().toLocaleString();
    table.load("rows");
    await context.sync();
    const id = table.rows.items.length + 1;

    // 插入表格
    table.rows.add(null, [[id, text, translation, time]]);
    await context.sync();

    resultDiv.innerText = `✅ 翻译结果：${translation}`;
  }).catch((error) => {
    //console.error("[ERROR] Excel.run failed:", error);
    document.getElementById("result").innerText = `❌ Error: ${error.message}`;
  });
}
