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

/**
 * Posts to Slack. Returns whether it actually landed.
 *
 * The RETURN VALUE matters: callers that record "this was announced" must not
 * do so when the post failed, or the message is lost for good. A misconfigured
 * token silently dropped a notification before this returned anything.
 */
export async function postSlackMessage(channel: string, text: string): Promise<boolean> {
  const token = slackBotToken();
  if (!token || !channel) {
    console.error("[slack] no token/channel — skipping message:", text);
    return false;
  }
  try {
    const res = await fetch("https://slack.com/api/chat.postMessage", {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json; charset=utf-8" },
      body: JSON.stringify({ channel, text }),
    });
    const body = (await res.json()) as { ok?: boolean; error?: string };
    if (!body.ok) {
      console.error(`[slack] chat.postMessage failed: ${body.error ?? res.status}`);
      return false;
    }
    return true;
  } catch (err) {
    console.error("[slack] chat.postMessage threw:", err);
    return false;
  }
}

/** Posts to the configured ops channel if one is set (no-op otherwise). */
export async function notifyOps(text: string): Promise<void> {
  const channel = opsChannel();
  if (channel) await postSlackMessage(channel, text);
}

/**
 * Resolves a Slack user id from an email address, or null.
 *
 * Needs the `users:read.email` scope. Best-effort: a lookup failure is logged
 * and returns null, so a DM is skipped rather than the whole reconcile failing.
 * Cached for five minutes — the same handful of approvers come up repeatedly.
 */
let slackIdCache: { at: number; byEmail: Map<string, string | null> } | null = null;
const SLACK_ID_CACHE_MS = 5 * 60 * 1000;

export async function lookupSlackUserId(email: string): Promise<string | null> {
  const token = slackBotToken();
  const key = email.trim().toLowerCase();
  if (!token || !key) return null;
  if (!slackIdCache || Date.now() - slackIdCache.at > SLACK_ID_CACHE_MS) {
    slackIdCache = { at: Date.now(), byEmail: new Map() };
  }
  const cached = slackIdCache.byEmail.get(key);
  if (cached !== undefined) return cached;

  try {
    const res = await fetch(`https://slack.com/api/users.lookupByEmail?email=${encodeURIComponent(key)}`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const body = (await res.json()) as { ok?: boolean; error?: string; user?: { id?: string } };
    if (!body.ok || !body.user?.id) {
      console.error(`[slack] users.lookupByEmail failed for ${key}: ${body.error ?? res.status}`);
      slackIdCache.byEmail.set(key, null);
      return null;
    }
    slackIdCache.byEmail.set(key, body.user.id);
    return body.user.id;
  } catch (err) {
    console.error(`[slack] users.lookupByEmail threw for ${key}:`, err);
    return null;
  }
}

/**
 * DMs a person by email address. Returns whether it landed.
 *
 * Posting to a user id opens the direct message conversation, so no separate
 * conversations.open is needed.
 */
export async function dmByEmail(email: string, text: string): Promise<boolean> {
  const userId = await lookupSlackUserId(email);
  if (!userId) return false;
  return postSlackMessage(userId, text);
}
