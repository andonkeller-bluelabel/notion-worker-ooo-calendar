/**
 * Minimal best-effort Slack notifier (chat.postMessage) for operational
 * alerts — an approved request we couldn't put on the calendar, a Graph 403
 * that means the Application Access Policy drifted. Logs on failure and never
 * throws into the caller. Modeled on the Kantata and Read.ai workers' slack.ts.
 * Requires SLACK_BOT_TOKEN (chat:write, invited to the channel).
 */

import { slackBotToken, opsChannel } from "./env.js";

/** Escapes the three characters Slack treats as markup in message text. */
export function slackEscape(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** `<url|label>` with the label escaped — a bare `|` in a label breaks the link. */
export function slackLink(label: string, url: string | undefined | null): string {
  const safe = slackEscape(label).replace(/\|/g, "-");
  return url ? `<${url}|${safe}>` : safe;
}

export async function postSlackMessage(channel: string, text: string): Promise<void> {
  const token = slackBotToken();
  if (!token || !channel) {
    console.error("[slack] no token/channel — skipping message:", text);
    return;
  }
  try {
    const res = await fetch("https://slack.com/api/chat.postMessage", {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json; charset=utf-8" },
      body: JSON.stringify({ channel, text }),
    });
    const body = (await res.json()) as { ok?: boolean; error?: string };
    if (!body.ok) console.error(`[slack] chat.postMessage failed: ${body.error ?? res.status}`);
  } catch (err) {
    console.error("[slack] chat.postMessage threw:", err);
  }
}

/** Posts to the configured ops channel if one is set (no-op otherwise). */
export async function notifyOps(text: string): Promise<void> {
  const channel = opsChannel();
  if (channel) await postSlackMessage(channel, text);
}
