import cors from "cors";
import dotenv from "dotenv";
import express from "express";

dotenv.config({ path: ".env.local" });
dotenv.config();

const app = express();
const port = Number(process.env.API_PORT ?? 3001);
const baseUrl = normalizeBaseUrl(process.env.AI_BASE_URL ?? "http://localhost:20128/v1");
const apiKey = process.env.AI_API_KEY ?? process.env.OPENAI_API_KEY ?? "";
const configuredModel = process.env.AI_MODEL?.trim() || "cx/gpt-5.5";
const WORD_BUDGET_RETRIES = 2;

app.disable("x-powered-by");
app.use(cors({ origin: [/^https:\/\/localhost:3000$/, /^https:\/\/127\.0\.0\.1:3000$/] }));
app.use(express.json({ limit: "80kb" }));
app.use((_request, response, next) => {
  response.setHeader("Cache-Control", "no-store");
  next();
});

app.get("/api/health", (_request, response) => {
  response.json({
    ok: true,
    provider: baseUrl,
    hasKey: Boolean(apiKey),
    model: configuredModel
  });
});

app.post("/api/rewrite", async (request, response) => {
  try {
    const payload = validateRewriteRequest(request.body);
    const model = await resolveModel();
    const rewrittenHtml = await rewriteWithLocalProvider(payload, model);
    const rewrittenText = htmlToPlainText(rewrittenHtml);

    response.json({
      text: rewrittenText,
      html: rewrittenHtml,
      italicTerms: extractItalicTerms(rewrittenHtml),
      model,
      wordCount: countWords(rewrittenText)
    });
  } catch (error) {
    response.status(400).json({
      error: error instanceof Error ? error.message : "Unable to process request."
    });
  }
});

app.listen(port, "127.0.0.1", () => {
  console.log(`Draft Context AI API listening on http://127.0.0.1:${port}`);
});

async function rewriteWithLocalProvider(payload, model) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 60000);
  const budget = getWordBudget(payload.targetWords);

  try {
    let bestText = sanitizeItalicHtml(
      await requestChatCompletion({
        model,
        temperature: 0.28,
        messages: buildMessages(payload, budget),
        signal: controller.signal
      })
    );
    let bestDistance = getBudgetDistance(countWords(bestText), budget);

    for (let attempt = 0; attempt < WORD_BUDGET_RETRIES && bestDistance > 0; attempt += 1) {
      const correctedText = sanitizeItalicHtml(
        await requestChatCompletion({
          model,
          temperature: 0.15,
          messages: buildCorrectionMessages(payload, bestText, budget),
          signal: controller.signal
        })
      );
      const correctedDistance = getBudgetDistance(countWords(correctedText), budget);

      if (correctedDistance <= bestDistance) {
        bestText = correctedText;
        bestDistance = correctedDistance;
      }
    }

    if (countWords(bestText) > budget.max) {
      return trimHtmlToWordLimit(bestText, budget.max);
    }

    return bestText;
  } finally {
    clearTimeout(timeout);
  }
}

async function requestChatCompletion({ model, temperature, messages, signal }) {
  const response = await fetch(`${baseUrl}/chat/completions`, {
    method: "POST",
    signal,
    headers: {
      "Content-Type": "application/json",
      ...(apiKey ? { Authorization: `Bearer ${apiKey}` } : {})
    },
    body: JSON.stringify({
      model,
      temperature,
      messages
    })
  });

  const body = await response.json().catch(() => null);

  if (!response.ok) {
    throw new Error(extractProviderError(body, response.status));
  }

  const text = body?.choices?.[0]?.message?.content?.trim();

  if (!text) {
    throw new Error("Provider returned an empty response.");
  }

  return stripOuterQuotes(text);
}

function buildMessages(payload, budget) {
  const modeInstruction = {
    extend:
      "Perluas teks terpilih secara akademik dan relevan dengan konteks sampai mendekati target final, tetapi jangan melewati batas kata wajib.",
    shrink:
      "Perkecil teks terpilih menjadi lebih padat sampai mendekati target final tanpa menghilangkan makna, istilah teknis, angka, sitasi, dan hubungan logis penting.",
    rewrite:
      "Ganti gaya kalimat/paragraf agar lebih formal, runtut, dan akademik sambil mengikuti target final."
  }[payload.mode];

  const context = payload.snapshot.context;

  return [
    {
      role: "system",
      content:
        "Kamu adalah asisten akademik untuk penulisan skripsi/thesis berbahasa Indonesia. Susun ulang teks dengan diksi yang saintifik, formal, research-oriented, runtut, dan sesuai gaya penulisan karya ilmiah. Fokus pada kejelasan argumen, hubungan sebab-akibat, presisi metodologis, serta koherensi dengan konteks penelitian. Target jumlah kata final adalah aturan utama. Hitung kata sebelum menjawab dan patuhi rentang wajib. Tulis hasil final saja. Jangan memberi pembuka, catatan, bullet, markdown, kutipan kode, atau penjelasan proses. Jangan menambahkan klaim, data, temuan, rujukan, atau angka baru yang tidak didukung konteks. Pertahankan istilah teknis, sitasi, angka, nama metode, variabel, dan konteks penelitian apa adanya jika ada. Tandai istilah asing atau frasa teknis berbahasa asing dengan tag <i>...</i>. Gunakan hanya tag <i>; jangan gunakan tag HTML lain."
    },
    {
      role: "user",
      content: [
        `Mode: ${payload.mode}`,
        `Target final: ${payload.targetWords} kata`,
        `Rentang wajib: ${budget.min}-${budget.max} kata`,
        `Batas keras: jangan lebih dari ${budget.max} kata.`,
        "Tag <i> dan </i> tidak dihitung sebagai kata.",
        "Contoh format istilah asing: metode <i>mutation rate</i> digunakan secara adaptif.",
        "Jika target kata bertentangan dengan mode, prioritaskan target kata.",
        `Instruksi: ${modeInstruction}`,
        "",
        "Konteks terdekat:",
        context.heading ? `Heading: ${context.heading}` : "Heading: -",
        context.before.length ? `Sebelum:\n${context.before.join("\n\n")}` : "Sebelum: -",
        context.after.length ? `Sesudah:\n${context.after.join("\n\n")}` : "Sesudah: -",
        "",
        "Teks terpilih:",
        payload.snapshot.selectedText
      ].join("\n")
    }
  ];
}

function buildCorrectionMessages(payload, previousText, budget) {
  return [
    {
      role: "system",
      content:
        "Kamu adalah editor akademik skripsi/thesis. Koreksi jumlah kata dengan ketat. Output hanya teks final. Gunakan hanya tag <i>...</i> untuk istilah asing. Jangan memakai catatan, bullet, markdown, penjelasan, atau tag HTML lain."
    },
    {
      role: "user",
      content: [
        `Target final: ${payload.targetWords} kata`,
        `Rentang wajib: ${budget.min}-${budget.max} kata`,
        `Jumlah kata respons sebelumnya: ${countWords(previousText)} kata`,
        `Mode awal: ${payload.mode}`,
        "Revisi teks berikut agar tetap ilmiah, formal, dan research-oriented, namun masuk rentang wajib.",
        "Tag <i> dan </i> tidak dihitung sebagai kata.",
        "Jangan menambah klaim, data, sitasi, angka, atau konteks baru.",
        "",
        "Teks yang harus dikoreksi:",
        previousText
      ].join("\n")
    }
  ];
}

async function resolveModel() {
  return configuredModel;
}

function validateRewriteRequest(body) {
  if (!body || typeof body !== "object") {
    throw new Error("Invalid request body.");
  }

  if (!["extend", "shrink", "rewrite"].includes(body.mode)) {
    throw new Error("Invalid mode.");
  }

  const targetWords = Number(body.targetWords);

  if (!Number.isFinite(targetWords) || targetWords < 10 || targetWords > 450) {
    throw new Error("Target words must be between 10 and 450.");
  }

  const selectedText = cleanText(body.snapshot?.selectedText ?? "");

  if (!selectedText) {
    throw new Error("Selection is empty.");
  }

  if (selectedText.length > 12000) {
    throw new Error("Selection is too long. Select a smaller passage.");
  }

  return {
    mode: body.mode,
    targetWords: Math.round(targetWords),
    snapshot: {
      selectedText,
      selectedWordCount: countWords(selectedText),
      context: {
        heading: cleanText(body.snapshot?.context?.heading ?? ""),
        before: cleanTextArray(body.snapshot?.context?.before ?? []),
        after: cleanTextArray(body.snapshot?.context?.after ?? [])
      }
    }
  };
}

function cleanTextArray(items) {
  if (!Array.isArray(items)) {
    return [];
  }

  return items.map((item) => cleanText(item)).filter(Boolean).slice(0, 4);
}

function cleanText(value) {
  return String(value).replace(/\s+/g, " ").trim();
}

function countWords(value) {
  const matches = htmlToPlainText(value).trim().match(/[^\s]+/g);
  return matches ? matches.length : 0;
}

function getWordBudget(targetWords) {
  const tolerance = Math.max(1, Math.ceil(targetWords * 0.1));

  return {
    target: targetWords,
    tolerance,
    min: Math.max(1, targetWords - tolerance),
    max: targetWords + tolerance
  };
}

function getBudgetDistance(wordCount, budget) {
  if (wordCount < budget.min) {
    return budget.min - wordCount;
  }

  if (wordCount > budget.max) {
    return wordCount - budget.max;
  }

  return 0;
}

function trimToWordLimit(value, maxWords) {
  const words = value.trim().split(/\s+/);

  if (words.length <= maxWords) {
    return value;
  }

  const trimmed = words.slice(0, maxWords).join(" ").replace(/[,;:]+$/, "");

  return /[.!?]$/.test(trimmed) ? trimmed : `${trimmed}.`;
}

function sanitizeItalicHtml(value) {
  const normalized = normalizeItalicMarkers(stripOuterQuotes(value));
  const parts = normalized.split(/(<\s*\/?\s*i\s*>)/gi);
  let html = "";
  let isItalicOpen = false;

  for (const part of parts) {
    if (!part) {
      continue;
    }

    if (/^<\s*i\s*>$/i.test(part)) {
      if (!isItalicOpen) {
        html += "<i>";
        isItalicOpen = true;
      }

      continue;
    }

    if (/^<\s*\/\s*i\s*>$/i.test(part)) {
      if (isItalicOpen) {
        html += "</i>";
        isItalicOpen = false;
      }

      continue;
    }

    html += escapeHtml(part);
  }

  if (isItalicOpen) {
    html += "</i>";
  }

  return html
    .replace(/\s+/g, " ")
    .replace(/\s+([.,;:!?])/g, "$1")
    .trim();
}

function normalizeItalicMarkers(value) {
  return value
    .replace(/<\s*em\s*>/gi, "<i>")
    .replace(/<\s*\/\s*em\s*>/gi, "</i>")
    .replace(/\*([^*\n]{2,120})\*/g, "<i>$1</i>");
}

function htmlToPlainText(value) {
  return decodeHtmlEntities(String(value).replace(/<\s*\/?\s*i\s*>/gi, ""))
    .replace(/\s+/g, " ")
    .trim();
}

function extractItalicTerms(value) {
  const terms = [];
  const seen = new Set();
  const pattern = /<i>(.*?)<\/i>/gi;
  let match = pattern.exec(value);

  while (match) {
    const term = htmlToPlainText(match[1]);
    const key = term.toLowerCase();

    if (term && !seen.has(key)) {
      seen.add(key);
      terms.push(term);
    }

    match = pattern.exec(value);
  }

  return terms;
}

function trimHtmlToWordLimit(value, maxWords) {
  const html = sanitizeItalicHtml(value);
  const parts = html.split(/(<\/?i>|\s+)/i);
  let wordCount = 0;
  let output = "";
  let isItalicOpen = false;

  for (const part of parts) {
    if (!part) {
      continue;
    }

    if (/^<i>$/i.test(part)) {
      if (!isItalicOpen) {
        output += "<i>";
        isItalicOpen = true;
      }

      continue;
    }

    if (/^<\/i>$/i.test(part)) {
      if (isItalicOpen) {
        output += "</i>";
        isItalicOpen = false;
      }

      continue;
    }

    if (/^\s+$/.test(part)) {
      if (wordCount > 0 && wordCount < maxWords) {
        output += " ";
      }

      continue;
    }

    if (wordCount >= maxWords) {
      break;
    }

    output += part;
    wordCount += 1;
  }

  output = output.trim().replace(/\s+([.,;:!?])/g, "$1").replace(/[,;:]+$/, "");

  if (isItalicOpen) {
    output += "</i>";
  }

  return /[.!?]$/.test(htmlToPlainText(output)) ? output : `${output}.`;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function decodeHtmlEntities(value) {
  return String(value)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&");
}

function normalizeBaseUrl(value) {
  return value.replace(/\/+$/, "");
}

function stripOuterQuotes(value) {
  const trimmed = value.trim();

  if (
    (trimmed.startsWith('"') && trimmed.endsWith('"')) ||
    (trimmed.startsWith("'") && trimmed.endsWith("'"))
  ) {
    return trimmed.slice(1, -1).trim();
  }

  return trimmed;
}

function extractProviderError(body, status) {
  const message = body?.error?.message ?? body?.message ?? `Provider request failed with status ${status}.`;

  return typeof message === "string" ? message : `Provider request failed with status ${status}.`;
}
