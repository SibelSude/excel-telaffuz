Office.onReady(() => {
  Office.actions.associate("reviewNext", reviewNext);
});

function normalizeText(value) {
  return String(value ?? "").trim();
}

function parseTrDate(text) {
  const parts = normalizeText(text).split(".");
  if (parts.length !== 3) return null;

  const day = Number(parts[0]);
  const month = Number(parts[1]) - 1;
  const year = Number(parts[2]);

  if (Number.isNaN(day) || Number.isNaN(month) || Number.isNaN(year)) {
    return null;
  }

  const d = new Date(year, month, day);
  d.setHours(0, 0, 0, 0);
  return d;
}

function formatTrDate(date) {
  return date.toLocaleDateString("tr-TR");
}

async function clearTableRows(table, context) {
  const rows = table.rows;
  rows.load("items");
  await context.sync();

  for (let i = rows.items.length - 1; i >= 0; i--) {
    rows.items[i].delete();
  }

  await context.sync();
}

async function clearReviewScreen(sheet, context) {
  sheet.getRange("J4:J7").clear("Contents");
  sheet.getRange("J9").clear("Contents");
  sheet.getRange("J11:J13").clear("Contents");
  sheet.getRange("I16:K27").clear("Contents");
  sheet.getRange("J29:K29").clear("Contents");
  sheet.getRange("J14").clear("Contents");
  await context.sync();
}

async function reviewNext(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Tekrar");
      const wordsTable = context.workbook.tables.getItem("tblKelimeler");
      const reviewTable = context.workbook.tables.getItem("tblTekrar");

      const activeIdRange = sheet.getRange("J4");
      const resultRange = sheet.getRange("J9");

      activeIdRange.load("text");
      resultRange.load("text");
      await context.sync();

      const activeID = normalizeText(activeIdRange.text[0][0]);
      const result = normalizeText(resultRange.text[0][0]);

      // =====================================================
      // BİLDİM / BİLEMEDİM SEÇİLMEDİYSE
      // =====================================================
      if (result !== "Bildim" && result !== "Bilemedim") {
        const unansweredWordExisted = activeID !== "";
        const oldActiveID = activeID;

        await clearTableRows(reviewTable, context);
        await clearReviewScreen(sheet, context);

        const wordsBody = wordsTable.getDataBodyRange();
        wordsBody.load("values");
        await context.sync();

        const wordValues = wordsBody.values;

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const reviewRows = [];

        let firstID = "";
        let firstWord = "";

        let oldWordID = "";
        let oldWord = "";

        for (let i = 0; i < wordValues.length; i++) {
          const dateText = normalizeText(wordValues[i][8]);
          if (dateText === "") continue;

          const reviewDate = parseTrDate(dateText);
          if (!reviewDate) continue;

          if (reviewDate <= today) {
            const rowID = normalizeText(wordValues[i][0]);
            const rowWord = normalizeText(wordValues[i][1]);

            if (reviewRows.length === 0) {
              firstID = rowID;
              firstWord = rowWord;
            }

            if (unansweredWordExisted && rowID === oldActiveID) {
              oldWordID = rowID;
              oldWord = rowWord;
            }

            reviewRows.push([
              rowID,
              rowWord,
              normalizeText(wordValues[i][2]),
              normalizeText(wordValues[i][3]),
              normalizeText(wordValues[i][7]),
              normalizeText(wordValues[i][8]),
              normalizeText(wordValues[i][9])
            ]);
          }
        }

        if (reviewRows.length === 0) {
          sheet.getRange("J5").values = [
            ["Bugün veya önceki tarihlerden bekleyen tekrar bulunmuyor."]
          ];
          await context.sync();
          return;
        }

        reviewTable.rows.add(null, reviewRows);
        await context.sync();

        if (unansweredWordExisted && oldWordID !== "") {
          sheet.getRange("J4").values = [[oldWordID]];
          sheet.getRange("J5").values = [[oldWord]];
          sheet.getRange("J14").values = [
            ["Bildim/Bilemedim işaretlenmediği için yeniden soruldu."]
          ];
        } else {
          sheet.getRange("J4").values = [[firstID]];
          sheet.getRange("J5").values = [[firstWord]];
        }

        await context.sync();
        return;
      }

      // =====================================================
      // SONUÇ SEÇİLDİYSE
      // =====================================================
      sheet.getRange("J14").clear("Contents");

      if (activeID === "") {
        sheet.getRange("J14").values = [["Aktif kelime bulunamadı."]];
        await context.sync();
        return;
      }

      const searchedID = Number(activeID);

      const wordsBody = wordsTable.getDataBodyRange();
      wordsBody.load("values");
      await context.sync();

      const wordValues = wordsBody.values;
      let foundRow = -1;

      for (let i = 0; i < wordValues.length; i++) {
        if (Number(wordValues[i][0]) === searchedID) {
          foundRow = i;
          break;
        }
      }

      if (foundRow === -1) {
        sheet.getRange("J14").values = [["Kelime ana listede bulunamadı."]];
        await context.sync();
        return;
      }

      const currentStatus = normalizeText(wordValues[foundRow][9]);

      let days = 1;
      let newStatus = "Bilemedim";

      // =====================================================
      // ARALIKLI TEKRAR
      // =====================================================
      if (result === "Bildim") {
        if (currentStatus === "Yeni" || currentStatus === "Bilemedim") {
          days = 3;
          newStatus = "Seviye 1";
        } else if (currentStatus === "Seviye 1") {
          days = 7;
          newStatus = "Seviye 2";
        } else if (currentStatus === "Seviye 2") {
          days = 14;
          newStatus = "Seviye 3";
        } else if (currentStatus === "Seviye 3") {
          days = 30;
          newStatus = "Seviye 4";
        } else {
          days = 60;
          newStatus = "Seviye 5";
        }
      } else {
        days = 1;
        newStatus = "Bilemedim";
      }

      const today = new Date();
      today.setHours(0, 0, 0, 0);

      const nextReview = new Date(today);
      nextReview.setDate(nextReview.getDate() + days);

      const lastReviewCell = wordsBody.getCell(foundRow, 7);
      const nextReviewCell = wordsBody.getCell(foundRow, 8);
      const statusCell = wordsBody.getCell(foundRow, 9);

      lastReviewCell.values = [[formatTrDate(today)]];
      nextReviewCell.values = [[formatTrDate(nextReview)]];
      statusCell.values = [[newStatus]];
      await context.sync();

      // =====================================================
      // SONRAKİ KELİMEYİ BUL
      // =====================================================
      const reviewBody = reviewTable.getDataBodyRange();
      reviewBody.load("values");
      await context.sync();

      const reviewValues = reviewBody.values;

      let reviewRowIndex = -1;
      let nextID = "";
      let nextWord = "";

      for (let i = 0; i < reviewValues.length; i++) {
        if (Number(reviewValues[i][0]) === searchedID) {
          reviewRowIndex = i;

          if (i + 1 < reviewValues.length) {
            nextID = normalizeText(reviewValues[i + 1][0]);
            nextWord = normalizeText(reviewValues[i + 1][1]);
          }

          break;
        }
      }

      if (reviewRowIndex === -1) {
        sheet.getRange("J14").values = [
          ["Aktif kelime tekrar listesinde bulunamadı."]
        ];
        await context.sync();
        return;
      }

      const reviewRows = reviewTable.rows;
      reviewRows.load("items");
      await context.sync();

      reviewRows.items[reviewRowIndex].delete();
      await context.sync();

      // =====================================================
      // ÖNCEKİ KELİMENİN TÜM BİLGİLERİNİ TEMİZLE
      // =====================================================
      await clearReviewScreen(sheet, context);

      // =====================================================
      // SONRAKİ KELİME
      // =====================================================
      if (nextID !== "") {
        sheet.getRange("J4").values = [[nextID]];
        sheet.getRange("J5").values = [[nextWord]];
      } else {
        sheet.getRange("J5").values = [["Bugünkü tekrar tamamlandı."]];
      }

      await context.sync();
    });
  } catch (error) {
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("Tekrar");
        const message =
          error && error.message ? error.message : String(error);
        sheet.getRange("J14").values = [["Hata: " + message]];
        await context.sync();
      });
    } catch (_) {
      // Sessiz geç: komut penceresi açılmaz.
    }
  } finally {
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  }
}
