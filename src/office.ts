export type DraftMode = "extend" | "shrink" | "rewrite";
export type CaptionKind = "table" | "figure";
export type CaptionConfidence = "high" | "medium" | "low" | "missing";
export type PageSplitConfidence = "high" | "medium" | "low";
export type ReferenceConfidence = "high" | "medium" | "low";

export interface CaptionCandidate {
  id: string;
  kind: CaptionKind;
  index: number;
  label: "Tabel" | "Gambar";
  targetPosition: "above" | "below";
  rawText: string;
  title: string;
  fullCaption: string;
  confidence: CaptionConfidence;
  isCaptionStyle: boolean;
  note: string;
  syncedReferences?: number;
}

export interface CaptionApplySummary {
  applied: number;
  skipped: number;
  total: number;
  referencesSynced: number;
}

export interface PageSplitCandidate {
  id: string;
  index: number;
  text: string;
  wordCount: number;
  sentenceCount: number;
  confidence: PageSplitConfidence;
  keepTogether: boolean;
  widowControl: boolean;
  isFixed: boolean;
  reason: string;
  beforeBreakText?: string;
  afterBreakText?: string;
  splitSentence?: string;
  splitIndex?: number;
  source?: "rendered" | "fallback";
}

export interface PageSplitApplySummary {
  applied: number;
  skipped: number;
  total: number;
}

export interface ReferenceCandidate {
  id: string;
  label: "Tabel" | "Gambar";
  kind: CaptionKind;
  paragraphIndex: number;
  occurrenceIndex: number;
  mentionText: string;
  oldNumber: string;
  newNumber: string;
  targetTitle: string;
  targetParagraphIndex: number;
  bookmarkName: string;
  confidence: ReferenceConfidence;
  isSynced: boolean;
  contextText: string;
  note: string;
}

export interface ReferenceApplySummary {
  applied: number;
  skipped: number;
  total: number;
}

type AutoCaptionTarget = {
  paragraph: Word.Paragraph;
  caption: CaptionCandidate;
};

export interface AutoCaptionReferenceSyncPlan {
  label: "Tabel" | "Gambar";
  kind: CaptionKind;
  oldNumber: string;
  newNumber: string;
  mentionText: string;
  bookmarkName: string;
}

type AutoCaptionReferenceSyncTarget = AutoCaptionReferenceSyncPlan & {
  paragraph: Word.Paragraph;
};

type PageSplitParagraphData = {
  keepTogether?: boolean;
  widowControl?: boolean;
};

export interface RenderedPageSplitParagraph {
  paragraphIndex: number;
  beforeBreakText: string;
  afterBreakText: string;
  splitSentence: string;
}

export interface FallbackPageSplitParagraph {
  paragraphIndex: number;
  splitSentence: string;
}

export interface ParagraphContext {
  heading: string;
  before: string[];
  after: string[];
}

export interface SelectionSnapshot {
  selectedText: string;
  selectedWordCount: number;
  context: ParagraphContext;
  source: "word" | "preview";
}

const HEADING_HINTS = [
  /^bab\s+[ivxlcdm0-9]+/i,
  /^\d+(\.\d+){0,4}\s+\S+/,
  /^(pendahuluan|metodologi|hasil|pembahasan|kesimpulan|saran)\b/i
];
const CAPTION_FONT_NAME = "Bookman Old Style";
const CAPTION_FONT_SIZE = 10;
const CAPTION_ALIGNMENT = "centered";
const REFERENCE_MENTION_PATTERN = /\b(Tabel|Gambar)\s+((?:\d+|[IVXLCDM]+)(?:\.\d+)*)/gi;

export function countWords(value: string): number {
  const matches = value.trim().match(/[^\s]+/g);
  return matches ? matches.length : 0;
}

export function normalizeText(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

export function getCaptionParagraphFormat() {
  return {
    fontName: CAPTION_FONT_NAME,
    fontSize: CAPTION_FONT_SIZE,
    italic: false,
    alignment: CAPTION_ALIGNMENT
  };
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

export async function waitForOffice(): Promise<boolean> {
  if (typeof Office === "undefined") {
    return false;
  }

  return new Promise((resolve) => {
    const timeout = window.setTimeout(() => {
      resolve(false);
    }, 1200);

    Office.onReady((info) => {
      window.clearTimeout(timeout);
      resolve(info.host === Office.HostType.Word);
    });
  });
}

export async function readSelectionSnapshot(contextRadius: number): Promise<SelectionSnapshot> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    const selectedParagraphs = selection.paragraphs;

    selection.load("text");
    selectedParagraphs.load("items/text");

    await context.sync();

    const selectedText = normalizeText(selection.text);

    if (!selectedText) {
      throw new Error("Select a sentence or paragraph first.");
    }

    const selectedParagraphTexts = selectedParagraphs.items
      .map((paragraph) => normalizeText(paragraph.text))
      .filter(Boolean);
    const firstParagraph = selectedParagraphs.items[0];
    const lastParagraph = selectedParagraphs.items[selectedParagraphs.items.length - 1];
    const before = firstParagraph ? await readAdjacentParagraphs(context, firstParagraph, "previous", contextRadius) : [];
    const after = lastParagraph ? await readAdjacentParagraphs(context, lastParagraph, "next", contextRadius) : [];
    const heading = findNearestHeading(before, before.length);

    return {
      selectedText,
      selectedWordCount: countWords(selectedText),
      context: {
        heading,
        before: selectedParagraphTexts.length ? before : [],
        after
      },
      source: "word"
    };
  });
}

export async function replaceSelectionWith(text: string, formattedHtml = ""): Promise<void> {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const replacementHtml = formattedHtml ? sanitizeInlineHtml(formattedHtml) : escapeHtml(text);
    const insertedRange = selection.insertHtml(replacementHtml, Word.InsertLocation.replace);
    insertedRange.select();
    await context.sync();
  });
}

export async function scanCaptionCandidates(): Promise<CaptionCandidate[]> {
  return Word.run(async (context) => {
    const tables = context.document.body.tables;
    const pictures = context.document.body.inlinePictures;

    tables.load("items");
    pictures.load("items");

    await context.sync();

    const tableCaptionParagraphs = tables.items.map((table) => table.getParagraphBeforeOrNullObject());
    const figureCaptionParagraphs = pictures.items.map((picture) => picture.paragraph.getNextOrNullObject());

    [...tableCaptionParagraphs, ...figureCaptionParagraphs].forEach((paragraph) => {
      paragraph.load("isNullObject,text,styleBuiltIn");
    });

    await context.sync();

    return [
      ...tableCaptionParagraphs.map((paragraph, index) => {
        return buildCaptionCandidate({
          kind: "table",
          index,
          paragraph
        });
      }),
      ...figureCaptionParagraphs.map((paragraph, index) => {
        return buildCaptionCandidate({
          kind: "figure",
          index,
          paragraph
        });
      })
    ];
  });
}

export async function selectCaptionCandidate(candidate: CaptionCandidate): Promise<void> {
  await Word.run(async (context) => {
    if (candidate.kind === "table") {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      const table = tables.items[candidate.index];

      if (!table) {
        throw new Error("Table not found. Scan captions again.");
      }

      table.getRange(Word.RangeLocation.whole).select();
      await context.sync();
      return;
    }

    const pictures = context.document.body.inlinePictures;
    pictures.load("items");
    await context.sync();

    const picture = pictures.items[candidate.index];

    if (!picture) {
      throw new Error("Picture not found. Scan captions again.");
    }

    picture.select();
    await context.sync();
  });
}

export async function applyCaptionCandidate(candidate: CaptionCandidate): Promise<CaptionCandidate> {
  return Word.run(async (context) => {
    const paragraph = await getCaptionParagraph(context, candidate);
    paragraph.load("isNullObject,text,styleBuiltIn");

    await context.sync();

    if (isNullObject(paragraph)) {
      throw new Error("Caption paragraph not found. Scan captions again.");
    }

    const refreshed = buildCaptionCandidate({
      kind: candidate.kind,
      index: candidate.index,
      paragraph
    });

    if (!canApplyCaption(refreshed)) {
      throw new Error("This item does not look like a valid caption.");
    }

    await replaceParagraphsWithAutoCaptions(context, [
      {
        paragraph,
        caption: refreshed
      }
    ]);

    paragraph.load("text,styleBuiltIn");
    await context.sync();

    const converted = buildCaptionCandidate({
      kind: candidate.kind,
      index: candidate.index,
      paragraph
    });

    const referencePlan = buildAutoCaptionReferenceSyncPlan({
      label: refreshed.label,
      kind: refreshed.kind,
      objectIndex: refreshed.index,
      oldCaptionText: refreshed.fullCaption,
      newCaptionText: converted.fullCaption
    });
    const syncedReferences = await syncAutoCaptionReferences(
      context,
      referencePlan ? [{ ...referencePlan, paragraph }] : []
    );

    return {
      ...converted,
      syncedReferences
    };
  });
}

export async function applyAllCaptionCandidates(candidates: CaptionCandidate[]): Promise<CaptionApplySummary> {
  const applicableCandidates = candidates.filter(canApplyCaption);

  if (!applicableCandidates.length) {
    return {
      applied: 0,
      skipped: candidates.length,
      total: candidates.length,
      referencesSynced: 0
    };
  }

  return Word.run(async (context) => {
    let applied = 0;
    let skipped = candidates.length - applicableCandidates.length;

    const tables = context.document.body.tables;
    const pictures = context.document.body.inlinePictures;

    tables.load("items");
    pictures.load("items");

    await context.sync();

    const targets = applicableCandidates.map((candidate) => {
      const paragraph = getLoadedCaptionParagraph(candidate, tables.items, pictures.items);
      return {
        candidate,
        paragraph
      };
    });

    targets.forEach(({ paragraph }) => {
      paragraph?.load("isNullObject,text,styleBuiltIn");
    });

    await context.sync();

    const validTargets: AutoCaptionTarget[] = [];

    targets.forEach(({ candidate, paragraph }) => {
      if (!paragraph || isNullObject(paragraph)) {
        skipped += 1;
        return;
      }

      const refreshed = buildCaptionCandidate({
        kind: candidate.kind,
        index: candidate.index,
        paragraph
      });

      if (!canApplyCaption(refreshed)) {
        skipped += 1;
        return;
      }

      validTargets.push({
        paragraph,
        caption: refreshed
      });
      applied += 1;

    });

    await replaceParagraphsWithAutoCaptions(context, validTargets);

    validTargets.forEach(({ paragraph }) => {
      paragraph.load("text,styleBuiltIn");
    });

    await context.sync();

    const referenceTargets = validTargets
      .map(({ paragraph, caption }) => {
        const converted = buildCaptionCandidate({
          kind: caption.kind,
          index: caption.index,
          paragraph
        });
        const plan = buildAutoCaptionReferenceSyncPlan({
          label: caption.label,
          kind: caption.kind,
          objectIndex: caption.index,
          oldCaptionText: caption.fullCaption,
          newCaptionText: converted.fullCaption
        });

        return plan ? { ...plan, paragraph } : null;
      })
      .filter((target): target is AutoCaptionReferenceSyncTarget => Boolean(target));
    const referencesSynced = await syncAutoCaptionReferences(context, referenceTargets);

    return {
      applied,
      skipped,
      total: candidates.length,
      referencesSynced
    };
  });
}

export async function scanPageSplitCandidates(): Promise<PageSplitCandidate[]> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    const ooxml = context.document.body.getOoxml();
    paragraphs.load("items/text,items/styleBuiltIn,items/keepTogether,items/widowControl");

    await context.sync();

    const renderedSplits = groupRenderedPageSplitsByParagraph(findRenderedPageSplitParagraphs(ooxml.value));
    const candidates: PageSplitCandidate[] = [];

    paragraphs.items.forEach((paragraph, index) => {
      const splitInfos = renderedSplits.get(index) ?? [];

      splitInfos.forEach((splitInfo, splitIndex) => {
        const candidate = buildPageSplitCandidate(paragraph, index, splitInfo, splitIndex);

        if (candidate) {
          candidates.push(candidate);
        }
      });
    });

    if (!candidates.length) {
      const fallbackSplits = groupFallbackPageSplitsByParagraph(
        findFallbackPageSplitParagraphs(paragraphs.items.map((paragraph) => paragraph.text ?? ""))
      );

      paragraphs.items.forEach((paragraph, index) => {
        const splitInfos = fallbackSplits.get(index) ?? [];

        splitInfos.forEach((splitInfo, splitIndex) => {
          const candidate = buildFallbackPageSplitCandidate(paragraph, index, splitInfo, splitIndex);

          if (candidate) {
            candidates.push(candidate);
          }
        });
      });
    }

    return candidates;
  });
}

export async function selectPageSplitCandidate(candidate: PageSplitCandidate): Promise<void> {
  await Word.run(async (context) => {
    const paragraph = await getParagraphByIndex(context, candidate.index);
    paragraph.getRange(Word.RangeLocation.whole).select();
    await context.sync();
  });
}

export async function applyPageSplitFix(candidate: PageSplitCandidate): Promise<PageSplitCandidate> {
  return Word.run(async (context) => {
    const paragraph = await getParagraphByIndex(context, candidate.index);
    applyPageSplitFormatting(paragraph);

    await context.sync();

    paragraph.load("text,styleBuiltIn,keepTogether,widowControl");
    await context.sync();

    return markPageSplitCandidateFixed(candidate, paragraph);
  });
}

export async function applyAllPageSplitFixes(candidates: PageSplitCandidate[]): Promise<PageSplitApplySummary> {
  const applicableCandidates = candidates.filter((candidate) => !candidate.isFixed);
  const pendingParagraphIndexes = Array.from(new Set(applicableCandidates.map((candidate) => candidate.index)));

  if (!pendingParagraphIndexes.length) {
    return {
      applied: 0,
      skipped: candidates.length,
      total: candidates.length
    };
  }

  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");

    await context.sync();

    let applied = 0;
    let skipped = candidates.length - pendingParagraphIndexes.length;

    pendingParagraphIndexes.forEach((paragraphIndex) => {
      const paragraph = paragraphs.items[paragraphIndex];

      if (!paragraph) {
        skipped += 1;
        return;
      }

      applyPageSplitFormatting(paragraph);
      applied += 1;
    });

    await context.sync();

    return {
      applied,
      skipped,
      total: candidates.length
    };
  });
}

export async function scanReferenceCandidates(): Promise<ReferenceCandidate[]> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/text,items/styleBuiltIn");

    await context.sync();

    return buildReferenceCandidatesFromParagraphs(paragraphs.items);
  });
}

export async function selectReferenceCandidate(candidate: ReferenceCandidate): Promise<void> {
  await Word.run(async (context) => {
    const paragraph = await getParagraphByIndex(context, candidate.paragraphIndex);
    const ranges = paragraph.search(candidate.mentionText, {
      matchCase: false,
      matchWholeWord: false
    });

    ranges.load("items");
    await context.sync();

    const range = ranges.items[candidate.occurrenceIndex] ?? ranges.items[0];

    if (!range) {
      paragraph.getRange(Word.RangeLocation.whole).select();
      await context.sync();
      return;
    }

    range.select();
    await context.sync();
  });
}

export async function applyReferenceCandidate(candidate: ReferenceCandidate): Promise<ReferenceCandidate> {
  const summary = await applyReferenceCandidates([candidate]);

  if (!summary.applied) {
    throw new Error("Reference mention could not be applied. Scan references again.");
  }

  return {
    ...candidate,
    isSynced: true,
    oldNumber: candidate.newNumber,
    mentionText: `${candidate.label} ${candidate.newNumber}`,
    note: "Reference field applied."
  };
}

export async function applyAllReferenceCandidates(candidates: ReferenceCandidate[]): Promise<ReferenceApplySummary> {
  return applyReferenceCandidates(candidates);
}

export function createPreviewSnapshot(text: string): SelectionSnapshot {
  const selectedText = normalizeText(text);

  return {
    selectedText,
    selectedWordCount: countWords(selectedText),
    source: "preview",
    context: {
      heading: "Preview mode",
      before: [],
      after: []
    }
  };
}

function buildCaptionCandidate({
  kind,
  index,
  paragraph
}: {
  kind: CaptionKind;
  index: number;
  paragraph: Word.Paragraph;
}): CaptionCandidate {
  const label = kind === "table" ? "Tabel" : "Gambar";
  const targetPosition = kind === "table" ? "above" : "below";

  if (isNullObject(paragraph)) {
    return {
      id: `${kind}-${index}`,
      kind,
      index,
      label,
      targetPosition,
      rawText: "",
      title: "",
      fullCaption: "",
      confidence: "missing",
      isCaptionStyle: false,
      note: kind === "table" ? "No paragraph above this table." : "No paragraph below this picture."
    };
  }

  const rawText = normalizeText(paragraph.text ?? "");
  const parsed = parseCaptionText(rawText, label);

  return {
    id: `${kind}-${index}`,
    kind,
    index,
    label,
    targetPosition,
    rawText,
    title: parsed.title,
    fullCaption: parsed.fullCaption,
    confidence: parsed.confidence,
    isCaptionStyle: paragraph.styleBuiltIn === "Caption",
    note: parsed.note
  };
}

function parseCaptionText(rawText: string, label: "Tabel" | "Gambar") {
  const fullCaption = normalizeText(rawText);

  if (!fullCaption) {
    return {
      title: "",
      fullCaption,
      confidence: "missing" as CaptionConfidence,
      note: `No ${label.toLowerCase()} caption text found.`
    };
  }

  const title = extractCaptionTitle(fullCaption, label);

  if (title) {
    return {
      title,
      fullCaption,
      confidence: "high" as CaptionConfidence,
      note: `Matched ${label} number and title.`
    };
  }

  if (new RegExp(`^${label}\\b`, "i").test(fullCaption)) {
    return {
      title: stripLooseCaptionPrefix(fullCaption, label),
      fullCaption,
      confidence: "medium" as CaptionConfidence,
      note: `Starts with ${label}, but the number pattern is unusual.`
    };
  }

  return {
    title: fullCaption,
    fullCaption,
    confidence: "low" as CaptionConfidence,
    note: `Text exists, but it does not start with ${label}.`
  };
}

function extractCaptionTitle(value: string, label: "Tabel" | "Gambar"): string {
  const escapedLabel = escapeRegExp(label);
  const numberPattern = "(?:\\d+|[IVXLCDM]+)(?:\\.\\d+)*";
  const captionPattern = new RegExp(
    `^\\s*${escapedLabel}\\s+${numberPattern}\\s*[\\.:\\-\\u2013\\u2014\\)]?\\s+(.+?)\\s*$`,
    "i"
  );
  const match = value.match(captionPattern);

  return match ? normalizeText(match[1]) : "";
}

function stripLooseCaptionPrefix(value: string, label: "Tabel" | "Gambar"): string {
  const escapedLabel = escapeRegExp(label);
  const loosePattern = new RegExp(`^\\s*${escapedLabel}\\s+`, "i");

  return normalizeText(value.replace(loosePattern, ""));
}

export function extractCaptionNumber(value: string, label: "Tabel" | "Gambar"): string {
  const escapedLabel = escapeRegExp(label);
  const numberPattern = "((?:\\d+|[IVXLCDM]+)(?:\\.\\d+)*)";
  const match = normalizeText(value).match(new RegExp(`^${escapedLabel}\\s+${numberPattern}\\b`, "i"));

  return match?.[1] ?? "";
}

export function buildAutoCaptionReferenceSyncPlan({
  label,
  kind,
  objectIndex,
  oldCaptionText,
  newCaptionText
}: {
  label: "Tabel" | "Gambar";
  kind: CaptionKind;
  objectIndex: number;
  oldCaptionText: string;
  newCaptionText: string;
}): AutoCaptionReferenceSyncPlan | null {
  const oldNumber = extractCaptionNumber(oldCaptionText, label);
  const newNumber = extractCaptionNumber(newCaptionText, label);

  if (!oldNumber || !newNumber) {
    return null;
  }

  return {
    label,
    kind,
    oldNumber,
    newNumber,
    mentionText: `${label} ${oldNumber}`,
    bookmarkName: makeAutoCaptionReferenceBookmarkName(label, objectIndex, oldNumber)
  };
}

function canApplyCaption(candidate: CaptionCandidate): boolean {
  return candidate.confidence === "high" || candidate.confidence === "medium";
}

async function replaceParagraphsWithAutoCaptions(
  context: Word.RequestContext,
  targets: AutoCaptionTarget[]
): Promise<void> {
  if (!targets.length) {
    return;
  }

  targets.forEach(({ paragraph, caption }) => {
    paragraph.insertText(`${caption.label} `, Word.InsertLocation.replace);
    paragraph.styleBuiltIn = "Caption";
  });

  await context.sync();

  targets.forEach(({ paragraph, caption }) => {
    paragraph
      .getRange(Word.RangeLocation.end)
      .insertField(Word.InsertLocation.before, "Seq", `${caption.label} \\* ARABIC`, true);
  });

  await context.sync();

  targets.forEach(({ paragraph, caption }) => {
    paragraph
      .getRange(Word.RangeLocation.end)
      .insertText(`. ${caption.title}`, Word.InsertLocation.before);
    formatCaptionParagraph(paragraph);
  });

  await context.sync();
}

function formatCaptionParagraph(paragraph: Word.Paragraph): void {
  const format = getCaptionParagraphFormat();

  paragraph.font.name = format.fontName;
  paragraph.font.size = format.fontSize;
  paragraph.font.italic = format.italic;
  paragraph.alignment = Word.Alignment.centered;
}

async function syncAutoCaptionReferences(
  context: Word.RequestContext,
  targets: AutoCaptionReferenceSyncTarget[]
): Promise<number> {
  const unambiguousTargets = filterUnambiguousAutoCaptionReferenceTargets(targets);

  if (!unambiguousTargets.length) {
    return 0;
  }

  const captionSearches = unambiguousTargets.map((target) => {
    const ranges = target.paragraph.search(`${target.label} ${target.newNumber}`, {
      matchCase: false,
      matchWholeWord: false
    });
    ranges.load("items");

    return {
      target,
      ranges
    };
  });

  await context.sync();

  const bookmarkTargets: AutoCaptionReferenceSyncTarget[] = [];

  captionSearches.forEach(({ target, ranges }) => {
    const range = ranges.items[0];

    if (!range) {
      return;
    }

    range.insertBookmark(target.bookmarkName);
    bookmarkTargets.push(target);
  });

  await context.sync();

  if (!bookmarkTargets.length) {
    return 0;
  }

  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items/text,items/styleBuiltIn");

  await context.sync();

  const mentionSearches: Array<{
    target: AutoCaptionReferenceSyncTarget;
    ranges: Word.RangeCollection;
  }> = [];

  paragraphs.items.forEach((paragraph) => {
    const paragraphText = normalizeText(paragraph.text ?? "");

    if (!paragraphText || paragraph.styleBuiltIn === "Caption") {
      return;
    }

    bookmarkTargets.forEach((target) => {
      if (!containsCaptionMention(paragraphText, target)) {
        return;
      }

      const ranges = paragraph.search(target.mentionText, {
        matchCase: false,
        matchWholeWord: false
      });
      ranges.load("items");
      mentionSearches.push({
        target,
        ranges
      });
    });
  });

  await context.sync();

  let synced = 0;

  mentionSearches.forEach(({ target, ranges }) => {
    ranges.items.forEach((range) => {
      range.insertField(Word.InsertLocation.replace, "Ref", `${target.bookmarkName} \\h`, true);
      synced += 1;
    });
  });

  await context.sync();

  return synced;
}

function filterUnambiguousAutoCaptionReferenceTargets(
  targets: AutoCaptionReferenceSyncTarget[]
): AutoCaptionReferenceSyncTarget[] {
  const counts = new Map<string, number>();

  targets.forEach((target) => {
    const key = getAutoCaptionReferenceKey(target);
    counts.set(key, (counts.get(key) ?? 0) + 1);
  });

  return targets.filter((target) => counts.get(getAutoCaptionReferenceKey(target)) === 1);
}

function getAutoCaptionReferenceKey(target: AutoCaptionReferenceSyncPlan): string {
  return `${target.label.toLowerCase()} ${normalizeReferenceNumber(target.oldNumber)}`;
}

function containsCaptionMention(text: string, target: AutoCaptionReferenceSyncPlan): boolean {
  const escapedLabel = escapeRegExp(target.label);
  const escapedNumber = escapeRegExp(target.oldNumber);
  const mentionPattern = new RegExp(`\\b${escapedLabel}\\s+${escapedNumber}(?!\\.?\\d)`, "i");

  return mentionPattern.test(text);
}

function makeAutoCaptionReferenceBookmarkName(
  label: "Tabel" | "Gambar",
  objectIndex: number,
  oldNumber: string
): string {
  const prefix = label === "Tabel" ? "auto_tbl" : "auto_fig";
  const safeNumber = oldNumber.replace(/[^A-Za-z0-9_]+/g, "_").replace(/^_+|_+$/g, "") || "n";

  return `${prefix}_${objectIndex}_${safeNumber}`.slice(0, 40);
}

export function findRenderedPageSplitParagraphs(ooxml: string): RenderedPageSplitParagraph[] {
  const paragraphs = extractOoxmlParagraphs(ooxml);
  const splitParagraphs: RenderedPageSplitParagraph[] = [];

  paragraphs.forEach((paragraphXml, paragraphIndex) => {
    const segments = paragraphXml.split(/<(?:\w+:)?lastRenderedPageBreak\b[^>]*\/?>/gi);

    if (segments.length <= 1) {
      return;
    }

    for (let segmentIndex = 0; segmentIndex < segments.length - 1; segmentIndex += 1) {
      const beforeBreakText = normalizeText(segments.slice(0, segmentIndex + 1).map(extractTextFromOoxmlSegment).join(" "));
      const afterBreakText = normalizeText(segments.slice(segmentIndex + 1).map(extractTextFromOoxmlSegment).join(" "));
      const trailingSentence = getTrailingUnfinishedSentence(beforeBreakText);
      const leadingCompletion = getLeadingSentenceCompletion(afterBreakText);

      if (!trailingSentence || !leadingCompletion) {
        continue;
      }

      splitParagraphs.push({
        paragraphIndex,
        beforeBreakText,
        afterBreakText,
        splitSentence: normalizeText(`${trailingSentence} ${leadingCompletion}`)
      });
    }
  });

  return splitParagraphs;
}

function extractOoxmlParagraphs(ooxml: string): string[] {
  return Array.from(ooxml.matchAll(/<(?:\w+:)?p\b[\s\S]*?<\/(?:\w+:)?p>/gi), (match) => match[0]);
}

function extractTextFromOoxmlSegment(segment: string): string {
  const parts: string[] = [];
  const tokenPattern =
    /<(?:\w+:)?t\b[^>]*>([\s\S]*?)<\/(?:\w+:)?t>|<(?:\w+:)?tab\b[^>]*\/>|<(?:\w+:)?br\b[^>]*\/>|<(?:\w+:)?cr\b[^>]*\/>/gi;
  let match: RegExpExecArray | null;

  while ((match = tokenPattern.exec(segment))) {
    if (match[1] !== undefined) {
      parts.push(decodeXmlEntities(match[1]));
    } else if (/(?:\w+:)?tab/i.test(match[0])) {
      parts.push("\t");
    } else {
      parts.push(" ");
    }
  }

  return parts.join("");
}

function decodeXmlEntities(value: string): string {
  return value.replace(/&(?:#(\d+)|#x([\da-f]+)|amp|lt|gt|quot|apos);/gi, (entity, decimal, hex) => {
    if (decimal) {
      return String.fromCodePoint(Number.parseInt(decimal, 10));
    }

    if (hex) {
      return String.fromCodePoint(Number.parseInt(hex, 16));
    }

    const namedEntities: Record<string, string> = {
      "&amp;": "&",
      "&lt;": "<",
      "&gt;": ">",
      "&quot;": '"',
      "&apos;": "'"
    };

    return namedEntities[entity.toLowerCase()] ?? entity;
  });
}

function getTrailingUnfinishedSentence(value: string): string {
  const text = normalizeText(value);

  if (!text || /[.!?]["')\]]*$/.test(text)) {
    return "";
  }

  const lastBoundary = Math.max(text.lastIndexOf("."), text.lastIndexOf("!"), text.lastIndexOf("?"));

  return normalizeText(lastBoundary >= 0 ? text.slice(lastBoundary + 1) : text);
}

function getLeadingSentenceCompletion(value: string): string {
  const text = normalizeText(value);
  const match = text.match(/^[^.!?]*[.!?]/);

  return normalizeText(match?.[0] ?? "");
}

function groupRenderedPageSplitsByParagraph(
  renderedSplits: RenderedPageSplitParagraph[]
): Map<number, RenderedPageSplitParagraph[]> {
  const grouped = new Map<number, RenderedPageSplitParagraph[]>();

  renderedSplits.forEach((split) => {
    const current = grouped.get(split.paragraphIndex) ?? [];
    current.push(split);
    grouped.set(split.paragraphIndex, current);
  });

  return grouped;
}

export function findFallbackPageSplitParagraphs(paragraphs: string[]): FallbackPageSplitParagraph[] {
  const fallbackSplits: FallbackPageSplitParagraph[] = [];

  paragraphs.forEach((paragraphText, paragraphIndex) => {
    const text = normalizeText(paragraphText);

    if (!text || shouldIgnorePageSplitText(text)) {
      return;
    }

    const longestSentence = getLongestSentence(text);

    if (!longestSentence || !isProbablePageSplitSentence(longestSentence, text)) {
      return;
    }

    fallbackSplits.push({
      paragraphIndex,
      splitSentence: longestSentence
    });
  });

  return fallbackSplits;
}

function groupFallbackPageSplitsByParagraph(
  fallbackSplits: FallbackPageSplitParagraph[]
): Map<number, FallbackPageSplitParagraph[]> {
  const grouped = new Map<number, FallbackPageSplitParagraph[]>();

  fallbackSplits.forEach((split) => {
    const current = grouped.get(split.paragraphIndex) ?? [];
    current.push(split);
    grouped.set(split.paragraphIndex, current);
  });

  return grouped;
}

function getLongestSentence(value: string): string {
  const sentences = splitIntoSentences(value);

  return sentences.reduce((longest, sentence) => {
    return countWords(sentence) > countWords(longest) ? sentence : longest;
  }, "");
}

function splitIntoSentences(value: string): string[] {
  return Array.from(normalizeText(value).matchAll(/[^.!?]+[.!?]+|[^.!?]+$/g), (match) => normalizeText(match[0]))
    .filter(Boolean);
}

function isProbablePageSplitSentence(sentence: string, paragraphText: string): boolean {
  const sentenceWordCount = countWords(sentence);
  const paragraphWordCount = countWords(paragraphText);

  return sentenceWordCount >= 24 || paragraphWordCount >= 42;
}

function shouldIgnorePageSplitText(text: string): boolean {
  if (looksLikeHeading(text)) {
    return true;
  }

  return /^(tabel|gambar)\s+(?:\d+|[ivxlcdm]+)(?:\.\d+)*\s*[.:)\-]/i.test(text);
}

function buildPageSplitCandidate(
  paragraph: Word.Paragraph,
  index: number,
  splitInfo: RenderedPageSplitParagraph,
  splitIndex: number
): PageSplitCandidate | null {
  const paragraphText = normalizeText(paragraph.text ?? "");

  if (!paragraphText || shouldIgnorePageSplitParagraph(paragraph, paragraphText)) {
    return null;
  }

  const paragraphData = paragraph.toJSON() as PageSplitParagraphData;
  const keepTogether = Boolean(paragraphData.keepTogether);
  const widowControl = Boolean(paragraphData.widowControl);
  const isFixed = keepTogether && widowControl;
  const splitSentence = normalizeText(splitInfo.splitSentence);
  const wordCount = countWords(splitSentence);
  const sentenceCount = countSentences(splitSentence);

  return {
    id: `page-split-${index}-${splitIndex}`,
    index,
    text: splitSentence,
    wordCount,
    sentenceCount,
    confidence: getPageSplitConfidence(isFixed),
    keepTogether,
    widowControl,
    isFixed,
    reason: getPageSplitReason(isFixed),
    beforeBreakText: splitInfo.beforeBreakText,
    afterBreakText: splitInfo.afterBreakText,
    splitSentence,
    splitIndex,
    source: "rendered"
  };
}

function buildFallbackPageSplitCandidate(
  paragraph: Word.Paragraph,
  index: number,
  splitInfo: FallbackPageSplitParagraph,
  splitIndex: number
): PageSplitCandidate | null {
  const paragraphText = normalizeText(paragraph.text ?? "");

  if (!paragraphText || shouldIgnorePageSplitParagraph(paragraph, paragraphText)) {
    return null;
  }

  const paragraphData = paragraph.toJSON() as PageSplitParagraphData;
  const keepTogether = Boolean(paragraphData.keepTogether);
  const widowControl = Boolean(paragraphData.widowControl);
  const isFixed = keepTogether && widowControl;
  const splitSentence = normalizeText(splitInfo.splitSentence);

  return {
    id: `page-split-fallback-${index}-${splitIndex}`,
    index,
    text: splitSentence,
    wordCount: countWords(splitSentence),
    sentenceCount: countSentences(splitSentence),
    confidence: isFixed ? "low" : "medium",
    keepTogether,
    widowControl,
    isFixed,
    reason: isFixed
      ? "Already guarded; Word did not expose rendered page-break markers."
      : "Word did not expose rendered page-break markers; this long sentence is a likely split risk.",
    splitSentence,
    splitIndex,
    source: "fallback"
  };
}

function shouldIgnorePageSplitParagraph(paragraph: Word.Paragraph, text: string): boolean {
  if (paragraph.styleBuiltIn === "Caption") {
    return true;
  }

  if (looksLikeHeading(text)) {
    return true;
  }

  return /^(tabel|gambar)\s+(?:\d+|[ivxlcdm]+)(?:\.\d+)*\s*[.:)\-]/i.test(text);
}

function getPageSplitConfidence(isFixed: boolean): PageSplitConfidence {
  if (isFixed) {
    return "low";
  }

  return "high";
}

function getPageSplitReason(isFixed: boolean): string {
  if (isFixed) {
    return "Already guarded, but Word still rendered a page break inside this sentence.";
  }

  return "Sentence crosses a rendered page break before its ending punctuation.";
}

function countSentences(value: string): number {
  const matches = normalizeText(value).match(/[^.!?]+[.!?]+|[^.!?]+$/g);
  return matches ? matches.filter((sentence) => normalizeText(sentence)).length : 0;
}

function markPageSplitCandidateFixed(candidate: PageSplitCandidate, paragraph: Word.Paragraph): PageSplitCandidate {
  const paragraphData = paragraph.toJSON() as PageSplitParagraphData;

  return {
    ...candidate,
    keepTogether: Boolean(paragraphData.keepTogether),
    widowControl: Boolean(paragraphData.widowControl),
    isFixed: Boolean(paragraphData.keepTogether) && Boolean(paragraphData.widowControl),
    confidence: "low",
    reason: "Page split guard applied with keep together and widow control."
  };
}

function applyPageSplitFormatting(paragraph: Word.Paragraph): void {
  paragraph.set({
    keepTogether: true,
    widowControl: true
  } as never);
}

async function getParagraphByIndex(context: Word.RequestContext, index: number): Promise<Word.Paragraph> {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items");

  await context.sync();

  const paragraph = paragraphs.items[index];

  if (!paragraph) {
    throw new Error("Paragraph not found. Scan page splits again.");
  }

  return paragraph;
}

async function applyReferenceCandidates(candidates: ReferenceCandidate[]): Promise<ReferenceApplySummary> {
  const applicableCandidates = candidates.filter(canApplyReference);

  if (!applicableCandidates.length) {
    return {
      applied: 0,
      skipped: candidates.length,
      total: candidates.length
    };
  }

  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/text,items/styleBuiltIn");

    await context.sync();

    const freshCandidates = buildReferenceCandidatesFromParagraphs(paragraphs.items);
    const freshById = new Map(freshCandidates.map((candidate) => [candidate.id, candidate]));
    const targets = new Map<string, ReferenceCandidate>();

    applicableCandidates.forEach((candidate) => {
      const freshCandidate = freshById.get(candidate.id) ?? candidate;

      if (canApplyReference(freshCandidate)) {
        targets.set(freshCandidate.bookmarkName, freshCandidate);
      }
    });

    const targetSearches = Array.from(targets.values()).map((candidate) => {
      const paragraph = paragraphs.items[candidate.targetParagraphIndex];
      const ranges = paragraph?.search(`${candidate.label} ${candidate.newNumber}`, {
        matchCase: false,
        matchWholeWord: false
      });
      ranges?.load("items");

      return {
        candidate,
        ranges
      };
    });

    await context.sync();

    targetSearches.forEach(({ candidate, ranges }) => {
      const range = ranges?.items[0];

      if (range) {
        range.insertBookmark(candidate.bookmarkName);
      }
    });

    await context.sync();

    let applied = 0;
    let skipped = candidates.length - applicableCandidates.length;

    const mentionSearches = applicableCandidates.map((candidate) => {
      const freshCandidate = freshById.get(candidate.id) ?? candidate;
      const paragraph = paragraphs.items[freshCandidate.paragraphIndex];
      const ranges = paragraph?.search(freshCandidate.mentionText, {
        matchCase: false,
        matchWholeWord: false
      });
      ranges?.load("items");

      return {
        candidate: freshCandidate,
        ranges
      };
    });

    await context.sync();

    mentionSearches.forEach(({ candidate, ranges }) => {
      const range = ranges?.items[candidate.occurrenceIndex] ?? ranges?.items[0];

      if (!range) {
        skipped += 1;
        return;
      }

      range.insertField(Word.InsertLocation.replace, "Ref", `${candidate.bookmarkName} \\h`, true);
      applied += 1;
    });

    await context.sync();

    return {
      applied,
      skipped,
      total: candidates.length
    };
  });
}

function buildReferenceCandidatesFromParagraphs(paragraphs: Word.Paragraph[]): ReferenceCandidate[] {
  const captions = paragraphs
    .map((paragraph, paragraphIndex) => buildReferenceCaption(paragraph, paragraphIndex))
    .filter((caption): caption is ReferenceCaption => Boolean(caption));
  const candidates: ReferenceCandidate[] = [];

  paragraphs.forEach((paragraph, paragraphIndex) => {
    const text = normalizeText(paragraph.text ?? "");

    if (!text || paragraph.styleBuiltIn === "Caption") {
      return;
    }

    const matches = Array.from(text.matchAll(REFERENCE_MENTION_PATTERN));
    const occurrenceByMention = new Map<string, number>();

    matches.forEach((match) => {
      const label = normalizeReferenceLabel(match[1]);
      const oldNumber = match[2];
      const mentionText = `${label} ${oldNumber}`;
      const target = findReferenceTarget(captions, label, paragraphIndex);

      if (!target) {
        return;
      }

      const occurrenceIndex = occurrenceByMention.get(mentionText.toLowerCase()) ?? 0;
      occurrenceByMention.set(mentionText.toLowerCase(), occurrenceIndex + 1);

      const isSynced = normalizeReferenceNumber(oldNumber) === normalizeReferenceNumber(target.number);
      const distance = Math.abs(target.paragraphIndex - paragraphIndex);
      const confidence = getReferenceConfidence(distance, isSynced);

      candidates.push({
        id: `ref-${paragraphIndex}-${match.index ?? candidates.length}`,
        label,
        kind: label === "Tabel" ? "table" : "figure",
        paragraphIndex,
        occurrenceIndex,
        mentionText,
        oldNumber,
        newNumber: target.number,
        targetTitle: target.title,
        targetParagraphIndex: target.paragraphIndex,
        bookmarkName: target.bookmarkName,
        confidence,
        isSynced,
        contextText: text,
        note: getReferenceNote(label, oldNumber, target.number, distance, isSynced)
      });
    });
  });

  return candidates;
}

type ReferenceCaption = {
  label: "Tabel" | "Gambar";
  kind: CaptionKind;
  number: string;
  title: string;
  paragraphIndex: number;
  bookmarkName: string;
};

function buildReferenceCaption(paragraph: Word.Paragraph, paragraphIndex: number): ReferenceCaption | null {
  if (paragraph.styleBuiltIn !== "Caption") {
    return null;
  }

  const text = normalizeText(paragraph.text ?? "");
  const match = text.match(/^(Tabel|Gambar)\s+((?:\d+|[IVXLCDM]+)(?:\.\d+)*)\s*[.:)\-]?\s*(.*)$/i);

  if (!match) {
    return null;
  }

  const label = normalizeReferenceLabel(match[1]);
  const title = normalizeText(match[3] ?? "");

  return {
    label,
    kind: label === "Tabel" ? "table" : "figure",
    number: match[2],
    title,
    paragraphIndex,
    bookmarkName: makeReferenceBookmarkName(label, paragraphIndex)
  };
}

function findReferenceTarget(
  captions: ReferenceCaption[],
  label: "Tabel" | "Gambar",
  paragraphIndex: number
): ReferenceCaption | null {
  const sameLabel = captions.filter((caption) => caption.label === label);

  if (!sameLabel.length) {
    return null;
  }

  const previous = sameLabel
    .filter((caption) => caption.paragraphIndex < paragraphIndex)
    .sort((a, b) => b.paragraphIndex - a.paragraphIndex)[0];

  if (previous) {
    return previous;
  }

  return sameLabel
    .filter((caption) => caption.paragraphIndex > paragraphIndex)
    .sort((a, b) => a.paragraphIndex - b.paragraphIndex)[0] ?? null;
}

function normalizeReferenceLabel(value: string): "Tabel" | "Gambar" {
  return /^gambar$/i.test(value) ? "Gambar" : "Tabel";
}

function normalizeReferenceNumber(value: string): string {
  return value.toLowerCase();
}

function getReferenceConfidence(distance: number, isSynced: boolean): ReferenceConfidence {
  if (isSynced || distance <= 3) {
    return "high";
  }

  if (distance <= 12) {
    return "medium";
  }

  return "low";
}

function getReferenceNote(
  label: "Tabel" | "Gambar",
  oldNumber: string,
  newNumber: string,
  distance: number,
  isSynced: boolean
): string {
  if (isSynced) {
    return `${label} reference already matches ${newNumber}.`;
  }

  return `${label} ${oldNumber} can sync to ${label} ${newNumber}; target is ${distance} paragraph${distance === 1 ? "" : "s"} away.`;
}

function makeReferenceBookmarkName(label: "Tabel" | "Gambar", paragraphIndex: number): string {
  const prefix = label === "Tabel" ? "tbl" : "fig";
  return `ref_${prefix}_${paragraphIndex}`.slice(0, 40);
}

function canApplyReference(candidate: ReferenceCandidate): boolean {
  return !candidate.isSynced && (candidate.confidence === "high" || candidate.confidence === "medium");
}

async function readAdjacentParagraphs(
  context: Word.RequestContext,
  startParagraph: Word.Paragraph,
  direction: "previous" | "next",
  count: number
): Promise<string[]> {
  const texts: string[] = [];
  let cursor = startParagraph;
  let attempts = 0;

  while (texts.length < count && attempts < count + 6) {
    attempts += 1;

    const adjacent =
      direction === "previous" ? cursor.getPreviousOrNullObject() : cursor.getNextOrNullObject();

    adjacent.load("isNullObject,text");
    await context.sync();

    if (isNullObject(adjacent)) {
      break;
    }

    const text = normalizeText(adjacent.text ?? "");

    if (text) {
      if (direction === "previous") {
        texts.unshift(text);
      } else {
        texts.push(text);
      }
    }

    cursor = adjacent;
  }

  return texts;
}

function getLoadedCaptionParagraph(
  candidate: CaptionCandidate,
  tables: Word.Table[],
  pictures: Word.InlinePicture[]
): Word.Paragraph | null {
  if (candidate.kind === "table") {
    return tables[candidate.index]?.getParagraphBeforeOrNullObject() ?? null;
  }

  const picture = pictures[candidate.index];

  return picture ? picture.paragraph.getNextOrNullObject() : null;
}

async function getCaptionParagraph(
  context: Word.RequestContext,
  candidate: CaptionCandidate
): Promise<Word.Paragraph> {
  if (candidate.kind === "table") {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();

    const table = tables.items[candidate.index];

    if (!table) {
      throw new Error("Table not found. Scan captions again.");
    }

    return table.getParagraphBeforeOrNullObject();
  }

  const pictures = context.document.body.inlinePictures;
  pictures.load("items");
  await context.sync();

  const picture = pictures.items[candidate.index];

  if (!picture) {
    throw new Error("Picture not found. Scan captions again.");
  }

  return picture.paragraph.getNextOrNullObject();
}

function isNullObject(value: Word.Paragraph): boolean {
  return Boolean((value as Word.Paragraph & { isNullObject?: boolean }).isNullObject);
}

function sanitizeInlineHtml(value: string): string {
  const parts = value.split(/(<\s*\/?\s*i\s*>)/gi);
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

    html += escapeHtml(decodeHtmlEntities(part));
  }

  if (isItalicOpen) {
    html += "</i>";
  }

  return html.replace(/\s+/g, " ").trim();
}

function escapeHtml(value: string): string {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function decodeHtmlEntities(value: string): string {
  return String(value)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&");
}

function findSelectedParagraphIndex(paragraphs: string[], selectedText: string): number {
  const normalizedSelection = normalizeText(selectedText).toLowerCase();

  if (!normalizedSelection) {
    return -1;
  }

  const exactIndex = paragraphs.findIndex((paragraph) => paragraph.toLowerCase() === normalizedSelection);

  if (exactIndex >= 0) {
    return exactIndex;
  }

  return paragraphs.findIndex((paragraph) => paragraph.toLowerCase().includes(normalizedSelection));
}

function buildContext(paragraphs: string[], selectedIndex: number, radius: number): ParagraphContext {
  if (selectedIndex < 0) {
    return {
      heading: "",
      before: paragraphs.slice(0, Math.min(radius, paragraphs.length)),
      after: []
    };
  }

  const beforeStart = Math.max(0, selectedIndex - radius);
  const afterEnd = Math.min(paragraphs.length, selectedIndex + radius + 1);

  return {
    heading: findNearestHeading(paragraphs, selectedIndex),
    before: paragraphs.slice(beforeStart, selectedIndex),
    after: paragraphs.slice(selectedIndex + 1, afterEnd)
  };
}

function findNearestHeading(paragraphs: string[], selectedIndex: number): string {
  for (let index = selectedIndex - 1; index >= 0; index -= 1) {
    const text = paragraphs[index];

    if (looksLikeHeading(text)) {
      return text;
    }
  }

  return "";
}

function looksLikeHeading(text: string): boolean {
  if (text.length > 90 || text.endsWith(".")) {
    return false;
  }

  return HEADING_HINTS.some((pattern) => pattern.test(text));
}
