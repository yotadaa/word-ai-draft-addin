import type { DraftMode, SelectionSnapshot } from "./office";

export interface RewriteRequest {
  mode: DraftMode;
  targetWords: number;
  snapshot: SelectionSnapshot;
}

export interface RewriteResponse {
  text: string;
  html: string;
  italicTerms: string[];
  model: string;
  wordCount: number;
}

export async function requestRewrite(payload: RewriteRequest): Promise<RewriteResponse> {
  const response = await fetch("/api/rewrite", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  const body = await response.json().catch(() => null);

  if (!response.ok) {
    const message = body?.error ?? "AI request failed";
    throw new Error(message);
  }

  return body as RewriteResponse;
}
