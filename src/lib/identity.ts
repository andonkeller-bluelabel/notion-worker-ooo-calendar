/**
 * Decides how to fill in who a request belongs to.
 *
 * Two forms feed OOO Entries and they establish identity differently. A Core
 * Request is submitted by a signed-in person, so `Created by` is authoritative
 * and carries a verified email. A Flex Request may be anonymous, so the typed
 * email is the only trustworthy signal and `Created by` must be ignored —
 * trusting it there would attribute someone's leave to whoever Notion records
 * as the creator of an anonymous submission.
 *
 * Pure and unit-tested. The Notion user lookup happens in the caller.
 */

import { SourceForm } from "./schema.js";

export interface CreatedBy {
  id: string;
  name: string | null;
  email: string | null;
  /** False for integrations and bots — never attribute a request to one. */
  isPerson: boolean;
}

export interface IdentityInputs {
  /** The row's `Source`, or null on rows predating the property. */
  source: string | null;
  /** Whether BlueLabeler already holds someone. */
  hasBlueLabeler: boolean;
  /** The row's `Your Email`, or null. */
  email: string | null;
  createdBy: CreatedBy | null;
}

export interface IdentityFill {
  blueLabelerId?: string;
  email?: string;
}

/**
 * What to write, given the row and (for the email path) the Notion user whose
 * address matches `email`. Returns an empty object when nothing should change.
 *
 * Only ever fills a blank. A value a human put there — filing on a colleague's
 * behalf, or correcting a bad attribution — is never overwritten, which also
 * means the sweep can't undo a correction ten minutes later.
 */
export function identityFill(inputs: IdentityInputs, matchedUserId: string | null): IdentityFill {
  const { source, hasBlueLabeler, email, createdBy } = inputs;
  const fill: IdentityFill = {};
  // `Created by` is authoritative only for the signed-in form, and only when
  // it resolves to a real person rather than an integration.
  const trustedCreator = source === SourceForm.CORE && createdBy?.isPerson ? createdBy : null;

  if (!email && trustedCreator?.email) fill.email = trustedCreator.email;

  if (!hasBlueLabeler) {
    if (trustedCreator) fill.blueLabelerId = trustedCreator.id;
    else if (matchedUserId) fill.blueLabelerId = matchedUserId;
  }
  return fill;
}

/**
 * Whether this row needs an email-to-user lookup at all. Keeps the caller from
 * listing the workspace's users on every reconcile: only an unattributed row
 * with an address to match on is worth the call.
 */
export function needsEmailLookup(inputs: IdentityInputs): boolean {
  if (inputs.hasBlueLabeler) return false;
  if (inputs.source === SourceForm.CORE && inputs.createdBy?.isPerson) return false;
  return Boolean(inputs.email);
}
