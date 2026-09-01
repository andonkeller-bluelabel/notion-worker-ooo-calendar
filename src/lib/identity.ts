/**
 * Decides how to fill in who a request belongs to.
 *
 * Two forms feed OOO Entries and they differ structurally: the Flex Request
 * REQUIRES an email because the responder may be anonymous, while the Core
 * Request asks for none because the responder is signed in. That difference is
 * the discriminator, so no hidden marker property is needed.
 *
 *   email present  -> the typed address is the identity; `Created by` is IGNORED
 *   email absent   -> `Created by` is the identity, and supplies the email too
 *
 * Ignoring `Created by` whenever an email exists is the load-bearing part.
 * Notion does not guarantee what it records as the creator of an anonymous form
 * submission — it could be a guest, the form's owner, or an integration — and
 * trusting it there would file a contractor's leave under whoever that is, then
 * put their name on a calendar the whole company reads.
 *
 * Pure and unit-tested. The Notion user lookup happens in the caller.
 */

export interface CreatedBy {
  id: string;
  name: string | null;
  email: string | null;
  /** False for integrations and bots — never attribute a request to one. */
  isPerson: boolean;
}

export interface IdentityInputs {
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
 * account address matches `email`. Empty object means leave the row alone.
 *
 * Only ever fills a blank. A value a human set — filing on a colleague's
 * behalf, or correcting a bad attribution — is never overwritten, so the sweep
 * cannot undo a correction ten minutes later.
 */
export function identityFill(inputs: IdentityInputs, matchedUserId: string | null): IdentityFill {
  const { hasBlueLabeler, email, createdBy } = inputs;
  const fill: IdentityFill = {};

  if (email) {
    // Flex-shaped: the address is the identity. Link a person only on an exact
    // match; no match is fine, since the name still derives from the address.
    if (!hasBlueLabeler && matchedUserId) fill.blueLabelerId = matchedUserId;
    return fill;
  }

  // Core-shaped: nobody typed an address, so the signed-in creator is who this
  // is for — provided it resolves to a real person rather than an integration.
  if (!createdBy?.isPerson) return fill;
  if (createdBy.email) fill.email = createdBy.email;
  if (!hasBlueLabeler) fill.blueLabelerId = createdBy.id;
  return fill;
}

/**
 * Whether this row needs an email-to-user lookup at all. Keeps the caller from
 * listing the workspace's users on every reconcile: only an unattributed row
 * with an address to match on is worth the call.
 */
export function needsEmailLookup(inputs: IdentityInputs): boolean {
  return !inputs.hasBlueLabeler && Boolean(inputs.email);
}

/**
 * The display name to attach to a BlueLabeler we just filled in, or null.
 *
 * Only the source we ACTUALLY filled from may supply the name, and it becomes
 * the calendar subject, so a wrong one is visible company-wide. Getting this
 * backwards once put "✈️ Anonymous" on a row whose BlueLabeler had been
 * correctly resolved to a real person: an anonymous form submission's creator
 * is a sentinel user literally named "Anonymous", and it was being read on the
 * email-match path where it has no bearing.
 *
 * Null is safe — the name then derives from the email address instead.
 */
export function filledName(
  blueLabelerId: string | undefined,
  matched: { id: string; name: string | null } | null,
  creator: CreatedBy | null,
): string | null {
  if (!blueLabelerId) return null;
  if (matched?.id === blueLabelerId) return matched.name;
  if (creator?.isPerson && creator.id === blueLabelerId) return creator.name;
  return null;
}
