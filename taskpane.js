Office.onReady(() => {
  document
    .getElementById("speakButton")
    .addEventListener("click", speakWord);

  refreshWord();

  // J5 değişti mi diye her 1 saniyede bir kontrol et
  setInterval(refreshWord, 1000);
});

let lastWord = "";

async function getWord() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Tekrar");
    const range = sheet.getRange("J5");

    range.load("text");
    await context.sync();

    return (range.text[0][0] || "").trim();
  });
}

async function refreshWord() {
  try {
    const word = await getWord();

    if (word !== lastWord) {
      lastWord = word;

      document.getElementById("word").textContent =
        word || "—";

      document.getElementById("status").textContent = "";
    }
  } catch (error) {
    document.getElementById("status").textContent =
      "Hata: " + error.message;
  }
}

async function speakWord() {
  const status = document.getElementById("status");

  try {
    const word = await getWord();

    document.getElementById("word").textContent =
      word || "—";

    lastWord = word;

    if (!word) {
      status.textContent =
        "Tekrar!J5 hücresinde İngilizce kelime yok.";
      return;
    }

    status.textContent = "";

    window.speechSynthesis.cancel();

    const utterance =
      new SpeechSynthesisUtterance(word);

    utterance.lang = "en-US";
    utterance.rate = 0.8;

    const voices =
      window.speechSynthesis.getVoices();

    const englishVoice =
      voices.find((voice) =>
        /^en(-|_)/i.test(voice.lang)
      );

    if (englishVoice) {
      utterance.voice = englishVoice;
    }

    window.speechSynthesis.speak(utterance);

  } catch (error) {
    status.textContent =
      "Hata: " + error.message;
  }
}
