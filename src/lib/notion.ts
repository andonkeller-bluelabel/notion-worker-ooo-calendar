/**
 * Thin, loosely-typed wrappers around the Notion client (`@notionhq/client`
 * v5). The generated request types are enormous, so — like the Kantata and
 * Read.ai workers — we build plain request objects and cast at the call
 * boundary, keeping higher-level code readable.
 *
 * Webhooks get an authenticated client as `context.notion` from the platform.
 * Scheduled syncs get one too, but only if NOTION_API_TOKEN is set — the
 * platform builds the sync's client from that internal-integration token and
 * throws without it.
 */

import type { Client } from "@notionhq/client";

export type Notion = Client;

export interface NotionPage {
  id: string;
  url?: string;
  /** Page-level creator; only an id, with no email. Prefer the property. */
  createdById?: string;
  /** True when the page is in the trash — the `page.deleted` case. */
  inTrash: boolean;
  properties: Record<string, any>;
}

const MAX_TEXT = 2000; // Notion rich_text content cap per item.

function truncate(s: string, max = MAX_TEXT): string {
  return s.length > max ? s.slice(0, max - 1) + "…" : s;
}

// --- property VALUE builders (Public API request shape) ---
export const P = {
  title: (s: string) => ({ title: [{ text: { content: truncate(s) } }] }),
  richText: (s: string) => ({ rich_text: s ? [{ text: { content: truncate(s) } }] : [] }),
  /** Clears a rich_text property. Notion represents "empty" as an empty array. */
  clearRichText: () => ({ rich_text: [] as unknown[] }),
};

export type Props = Record<string, unknown>;

// --- writes ---

export async function updateProps(notion: Notion, pageId: string, properties: Props): Promise<void> {
  await notion.pages.update({ page_id: pageId, properties } as Parameters<Notion["pages"]["update"]>[0]);
}

// --- reads ---

/**
 * Retrieves a single page with its full properties. Notion returns trashed
 * pages from `pages.retrieve` (with `in_trash: true`) rather than 404ing, so
 * this is how the sweep tells "deleted" apart from "never existed".
 */
export async function retrievePage(notion: Notion, pageId: string): Promise<NotionPage> {
  const res = (await notion.pages.retrieve({ page_id: pageId } as Parameters<Notion["pages"]["retrieve"]>[0])) as {
    id: string;
    url?: string;
    in_trash?: boolean;
    archived?: boolean;
    created_by?: { id?: string };
    properties?: Record<string, any>;
  };
  return {
    id: res.id,
    url: res.url,
    createdById: res.created_by?.id,
    inTrash: Boolean(res.in_trash ?? res.archived),
    properties: res.properties ?? {},
  };
}

/** Fetches every (non-trashed) page in a data source, following pagination. */
export async function queryAll(notion: Notion, dataSourceId: string, filter?: unknown): Promise<NotionPage[]> {
  const out: NotionPage[] = [];
  let cursor: string | undefined;
  do {
    const res = (await notion.dataSources.query({
      data_source_id: dataSourceId,
      ...(filter ? { filter } : {}),
      start_cursor: cursor,
      page_size: 100,
    } as Parameters<Notion["dataSources"]["query"]>[0])) as {
      results: Array<{
        id: string;
        url?: string;
        in_trash?: boolean;
        archived?: boolean;
        created_by?: { id?: string };
        properties?: Record<string, any>;
      }>;
      has_more: boolean;
      next_cursor: string | null;
    };
    for (const p of res.results ?? []) {
      if (p && p.id) {
        out.push({
          id: p.id,
          url: p.url,
          createdById: p.created_by?.id,
          inTrash: Boolean(p.in_trash ?? p.archived),
          properties: p.properties ?? {},
        });
      }
    }
    cursor = res.has_more ? (res.next_cursor ?? undefined) : undefined;
  } while (cursor);
  return out;
}

function joinRich(arr: unknown): string {
  if (!Array.isArray(arr)) return "";
  return arr
    .map((x) => (x && typeof x === "object" ? String((x as Record<string, unknown>).plain_text ?? "") : ""))
    .join("")
    .trim();
}

/** Reads a property as a single string (title/rich_text/select/status/email/url). */
export function readString(page: NotionPage, prop: string): string | null {
  const v = page.properties[prop] as Record<string, any> | undefined;
  if (!v) return null;
  switch (v.type) {
    case "title":
      return joinRich(v.title) || null;
    case "rich_text":
      return joinRich(v.rich_text) || null;
    case "select":
      return v.select?.name ?? null;
    case "status":
      return v.status?.name ?? null;
    case "email":
      return v.email ?? null;
    case "url":
      return v.url ?? null;
    default:
      return null;
  }
}

/** Reads a date property's `start` as the raw Notion string (YYYY-MM-DD or ISO). */
export function readDateStart(page: NotionPage, prop: string): string | null {
  const v = page.properties[prop] as Record<string, any> | undefined;
  if (!v || v.type !== "date") return null;
  const start = v.date?.start;
  return typeof start === "string" && start ? start : null;
}

/**
 * Reads a date property's `end` — the far side of a Notion date *range*.
 * Distinct from the separate "End Date" column; see oooRequest.ts, which
 * accepts either modeling.
 */
export function readDateEnd(page: NotionPage, prop: string): string | null {
  const v = page.properties[prop] as Record<string, any> | undefined;
  if (!v || v.type !== "date") return null;
  const end = v.date?.end;
  return typeof end === "string" && end ? end : null;
}

export interface NotionPerson {
  id: string;
  name: string | null;
  email: string | null;
}

/** Reads a `people` property. Bots and guests may lack a name and/or email. */
export function readPeople(page: NotionPage, prop: string): NotionPerson[] {
  const v = page.properties[prop] as Record<string, any> | undefined;
  if (!v || v.type !== "people" || !Array.isArray(v.people)) return [];
  return v.people
    .filter((p: any) => p && typeof p.id === "string")
    .map((p: any) => ({
      id: String(p.id),
      name: typeof p.name === "string" && p.name ? p.name : null,
      email: typeof p.person?.email === "string" && p.person.email ? p.person.email : null,
    }));
}

/**
 * Page ids come back from the API as dashed UUIDs but appear undashed in
 * Notion URLs. Normalize before comparing ids from different sources.
 */
export function normalizeId(id: string): string {
  return id.replace(/-/g, "").toLowerCase();
}

/** Canonical Notion URL for a page id, for when the API didn't return one. */
export function pageUrl(pageId: string, url?: string): string {
  return url ?? `https://www.notion.so/${pageId.replace(/-/g, "")}`;
}

/**
 * Reads a `created_by` property. Unlike the page-level `created_by`, the
 * PROPERTY comes back expanded with the person's name and verified email, so a
 * Core Request needs no extra user lookup.
 *
 * `isPerson` is false for integrations and bots. A request must never be
 * attributed to one — the backfill rows, for instance, were created by an
 * integration.
 */
export function readCreatedBy(page: NotionPage, prop: string): (NotionPerson & { isPerson: boolean }) | null {
  const v = page.properties[prop] as Record<string, any> | undefined;
  const u = v?.type === "created_by" ? v.created_by : undefined;
  if (!u || typeof u.id !== "string") return null;
  return {
    id: u.id,
    name: typeof u.name === "string" && u.name ? u.name : null,
    email: typeof u.person?.email === "string" && u.person.email ? u.person.email : null,
    isPerson: u.type === "person",
  };
}

/** People-property value builder. */
export function peopleValue(userIds: string[]) {
  return { people: userIds.map((id) => ({ object: "user", id })) };
}
