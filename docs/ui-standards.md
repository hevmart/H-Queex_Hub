# H-Queex Hub — UI Standards

This document is the standing set of UI rules for `templates/index.html`. It exists so
basic UX quality doesn't depend on being re-requested every time.

**Process requirement:**
- Before making any UI change: check this document.
- After making any UI change: verify the change against every rule below.
- If a design decision requires deviating from a rule here, flag the deviation explicitly
  in your response and explain why — never deviate silently.

## Brand palette (reference)

| Token | Hex | Use |
|---|---|---|
| Navy | `#16294A` | Primary brand colour — headers, primary buttons, nav |
| Gold | `#B08D57` | Accent — primary button border, active/on states |
| Steel blue | `#618096` | Muted/secondary text, hint text, optional field labels |
| Charcoal | `#37373A` | Primary body text |
| Dark grey | `#535356` | Secondary/muted body text |
| Border | `#E2E2E3` | Field and card borders |
| Bg light | `#F7F7F7` | Input backgrounds, alternating table rows |
| Bg deep | `#F2F2F2` | Disabled/readonly backgrounds |
| White | `#FFFFFF` | Base background, text on navy/gold |
| Error/red | `#C62828` | Errors, destructive actions, non-deductible/blocked states |

Font: Segoe UI throughout. No other typeface, no other colours.

## Form fields

- Placeholder text must always be shorter than the field width — **max 25 characters**
  for standard fields, **max 40** for wide/full-row fields.
- Placeholder text is never a sentence describing what the field does — **one or two
  words only** (e.g. `"Search suppliers..."`, not `"Search or type a new supplier /
  payee"`).
- **Exception: a primary Description field may use a realistic worked example as its
  placeholder** (e.g. `"e.g. Adobe Creative Cloud subscription August 2026"`), exceeding
  the normal character/word-count limits above. This applies only to the single primary
  Description field on a form where the description itself is the record's main
  identifying content (such as Expenses, where Title has been removed) — it does not
  apply to Notes/Details fields or to standard single-line inputs. Pair it with a short
  steel blue `.mini` hint line beneath explaining why specificity matters, rather than
  relying on the placeholder alone.
- Every input must be wide enough to display its expected content without horizontal
  scrolling (e.g. calculated currency fields sized for at least 8 characters including
  `€` and two decimals; free-text fields like Supplier / Payee spanning the full row).
- Labels sit above the field. Never rely on placeholder text as the only label.
- Mandatory fields show a visible asterisk (`.field.required label::after`).
- Locked/read-only fields must never look editable: no dropdown arrow, no input border —
  render as a badge or plain text instead.
- **Description fields must always be textareas with a minimum of 3 rows, never
  single-line inputs.** Any field labelled Description, Notes, or Details anywhere in
  the app must follow this rule — full width, resizable vertically, same font size and
  Segoe UI as every other field, with comfortable padding for reading and editing.
- **Every currency amount field shows a `€` prefix inside the field and always displays
  two decimal places** (`€18.00`, not `18`), applied consistently across Income,
  Expenses, Invoices, and Payroll. Values are formatted to two decimals on blur.

## Dropdowns

- Primary options (most likely to be picked) are visually prominent — regular weight,
  full-size, standard text colour.
- Secondary actions (add new, exceptions) are visually subdued — smaller text, steel
  blue `#618096`, separated from primary options by a divider line.
- Maximum 5–6 visible options before the list scrolls internally.
- Never show a dropdown arrow on a field the user cannot interact with.

## Warning messages

- Tell the user exactly what to do next, not just what the problem is.
- Never use a "Learn more" control that replaces the current message in place — if more
  detail is needed, use a proper modal or an expand-below, never an in-place swap.
- Severity meaning is fixed:
  - **Red** = blocked or non-deductible — requires action or acknowledgement.
  - **Amber** = needs attention but the user can proceed.
  - **Green/info** = informational only.
- Plain English only, no accounting/tax jargon without explanation.

## Buttons

- Primary action: navy `#16294A` background, white text, gold `#B08D57` bottom border.
- Secondary/cancel: white background, navy `#16294A` text and border.
- Destructive action: red `#C62828`, always requires confirmation before firing.
- Labels describe the outcome, not the mechanical action — `"Confirm and Save"` not
  `"Submit"`, `"Go Back to Form"` not `"Cancel"`, `"I Understand — Save Record"` not
  `"Save Anyway"`.
- Never use ALL CAPS on a button label or any interactive element.
- **Destructive actions must never appear as prominent buttons in list views — they
  belong inside the edit context only.** Archive/Cancel is never a button on a table row
  or a list card. Edit is the sole row-level action everywhere (Services, Clients,
  Suppliers, Income, Expenses, Invoices, Payroll, Subscriptions, Leads, Proposals, SOPs,
  Delivery Log, Documents, Projects). The archive/cancel action lives at the bottom of the
  record's edit form or edit modal, inside its own **Danger Zone**: a bordered, faint-red
  box (`rgba(198, 40, 40, 0.03)` background, `rgba(198, 40, 40, 0.18)` 1px border, 8px
  radius, ~28px margin above so it never crowds the field/button above it) with a small
  steel-blue uppercase "Danger zone" label at the top, a thin internal divider, then a
  small red `#C62828` text link with a trash icon ("Archive this record" — "Cancel
  invoice" for Invoices) using the shared `archive_in_modal()` macro (see Component
  Patterns) so every instance across the app is identical. The box itself is what
  separates this from whatever section follows (e.g. a list table) — don't rely on the
  parent container's default spacing.
- **A Danger Zone can hold more than one action** (e.g. Invoices: Cancel + Bad Debt, or
  Reverse Payment when a payment is recorded). Use the `danger_zone_item()` macro for each
  individual action and wrap them together in one shared zone box — `archive_in_modal()` is
  just `danger_zone_item()` pre-wrapped for the common single-action case. Every action still
  gets its own inline confirm step; a `blocked_message` on any one item replaces just that
  item with an explanatory line, it doesn't have to block the whole zone.
- Archive/cancel confirmation is an inline step below the link ("Are you sure? This will
  archive the record. Yes, archive / Cancel"), never a full modal on top of a modal and
  never a floating tooltip.
- **Every add/edit form must have a Cancel button.** Secondary style (white background,
  navy `#16294A` text and border), positioned immediately to the left of the primary
  Save/Update button. Clicking it discards any entered changes and returns the form to a
  blank/clean state; if the form has been modified from its initial state, confirm before
  discarding. This is distinct from the archive/cancel-record link above — this Cancel
  button abandons an in-progress edit, it never deletes or archives a saved record.
- If archiving a record is blocked by a dependency (e.g. an Income entry linked to a paid
  invoice), the archive link does not render at all — show a plain explanatory line
  instead (steel blue `#618096`) telling the user where to make the change.

## Breadcrumbs

- Every page reached from a parent section (e.g. a Finance module reached from the
  Finance hub) shows a breadcrumb above its page title: small steel blue `#618096` text,
  format `Parent › Current Page`, with the large navy bold page title directly beneath
  it. This is how a user knows where they are and how to get back — never rely on the
  main nav alone to convey hierarchy.

## Layout

- Fields in the same logical group are visually grouped together.
- Related fields sharing a row have equal visual weight unless one is clearly more
  important.
- No element may be truncated or overflow its container on a standard 1280px-wide
  screen.
- Binary yes/no options use a toggle switch, not a checkbox styled to look like a legal
  disclaimer.

## Typography

- Never use ALL CAPS for user-facing labels. Title Case for labels, sentence case for
  hint text and messages.
- Hint text is always smaller than label text and always steel blue `#618096`.
- Error messages are always red `#C62828` — never orange or yellow.

## Consistency

- Every form follows the same field order convention: Date → Title/Name → Description →
  Category → Amounts → secondary fields.
- Every list/table uses the same column styling: header in navy, alternating rows in
  white and `#F7F7F7`. Numeric columns (price, amount) right-aligned; status columns
  centred; name/description columns left-aligned; action columns (Edit/Archive)
  right-aligned — consistently, in every table, not decided per-page.
- Every page uses the same header structure: module title in navy, subtitle in steel
  blue.
- **The same interactive component must never be hand-rolled twice.** The in-modal
  archive/cancel action, Edit button, and the supplier/client smart-search field are each
  defined exactly once as a Jinja2 macro at the top of `templates/index.html`
  (`archive_in_modal()`, `edit_button()`, `supplier_search_field()`) and called everywhere
  they're needed. If a page needs a
  variant (different label, different confirm text), extend the macro's parameters —
  never copy-paste the markup and tweak it in place. A second hand-written copy of an
  existing component is a bug, not a new feature.
- **System status information never belongs in the main KPI/dashboard area.** Backup
  status, server/sync status, and similar operational information belong in a dedicated
  status bar or footer, in small steel blue `#618096` text — not as a KPI card competing
  with business metrics like Income or Net Cashflow.

## Component Patterns

General rule: **before building any interactive UI component, identify the standard,
universally understood pattern for that component type and follow it exactly.** Do not
invent a custom layout for something that already has a convention users recognise on
sight.

### Autocomplete / search fields

- The search results dropdown must always appear immediately below the input field with
  **zero gap**, so input and dropdown read as a single unified component.
- The dropdown must use `position: absolute` so it floats above the form layout and
  never pushes other elements down.
- The input field's border becomes gold `#B08D57` while the dropdown is open.
- The dropdown connects seamlessly to the bottom of the input field — no visible
  separation or double border between them (drop the dropdown's top border/radius so it
  reads as a continuation of the field, not a separate box).
- Secondary options or toggles related to the field always appear **below the closed
  dropdown**, never between the input and the results.
- Maximum 5 results visible before scrolling.
- This is a standard browser autocomplete pattern — follow it exactly, do not invent a
  new layout.

## Data Standards

- **Every record in every JSON file must have a stable `id` field**, generated once at
  creation and never modified afterward — not even when the record's name/title changes.
- Format is `PREFIX-NNN`: three digits minimum, zero-padded, sequential per prefix
  (`CLT-001`, `CLT-002`, …, `CLT-010`, `CLT-100`, …). Current prefixes:

  | Prefix | Entity |
  |---|---|
  | `CLT` | Clients |
  | `SUP` | Suppliers |
  | `SVC` | Services |
  | `PRJ` | Projects (internal id — see note below) |
  | `HQ-PRJ` | Projects (human-readable project number) |
  | `DLV` | Delivery log entries |
  | `SOP` | SOPs |
  | `DOC` | Company documents |
  | `CMP` | Compliance calendar manual entries |
  | `VAT` | VAT return periods (`VAT-2026-01` … `VAT-2026-06`, deterministic from the
    bi-monthly period — not persisted, since it's derived from a date rather than
    created by a user action) |
  | `HQ-LEAD` | CRM leads |
  | `HQ-PROP` | CRM proposals |
  | `EXP` | Expenses (needed so a subscription record can link back to the exact
    expense that paid a given billing period — see Subscription ↔ Expense Linking
    below) |

  Projects carry two ids by design: the internal `id` (plain UUID) is the actual join
  key used by Delivery Log and SOPs; `project_number` (`HQ-PRJ-2026-001`) is the
  human-readable label shown to users. Never conflate the two — cross-module references
  always use the internal `id`.
- **Join keys always use the stable id, never a name.** A record referencing a client,
  supplier, service, project, or document stores `client_id` / `supplier_id` /
  `service_id` / `project_id` / `linked_document_id` — the corresponding `*_name` /
  `*_title` field is a display convenience only, refreshed from the source record, and
  must never be used to look up or match the related record.
- **Display names are always resolved live from the source record via its id**, not read
  from a stored snapshot — so renaming a client, supplier, or service updates every page
  that references it immediately, with no migration needed for that rename itself.
- Migrations that assign missing or legacy-format ids must be **idempotent**: safe to run
  on every load, a no-op once every record already has a correctly-formatted id, and
  responsible for repointing any known cross-references if an id changes shape (e.g. an
  old UUID being replaced by its `PREFIX-NNN` equivalent).

### Invoice lifecycle

- **Invoices are never deleted.** The only removal path is Cancel, which flips `Status`
  to `Cancelled`, stamps `Cancellation Date`, and keeps the row in the register (hidden
  from the default register view, visible via the "Show cancelled" toggle) — audit trail
  first.
- **Only a `Draft` invoice can be edited.** Once `Status` is anything else (`Issued`,
  `Paid`, `Partially Paid`, `Overdue`, `Bad Debt`, `Cancelled`), the edit form renders with
  every field disabled (wrapped in a `<fieldset disabled>`) and a warning banner: "This
  invoice has been issued. To make changes you must cancel it and create a new invoice."
  The backend enforces this too (`update_invoice` rejects any update where the invoice's
  current status isn't Draft) — the disabled form is a UX courtesy, not the real guard.
- **Status has two tiers: manual and system-computed.** `Draft`, `Issued`, `Bad Debt` are
  the only values a user can pick from a dropdown (`INVOICE_MANUAL_STATUS_OPTIONS`).
  `Paid`, `Partially Paid`, `Overdue`, `Cancelled` are computed by the system (a recorded
  payment, a due date passing, the Cancel action) and rendered as a read-only status pill
  instead of a dropdown once reached. The same manual/auto split applies to Income:
  `Received` is never manually selected — logging a non-invoiced income entry means the
  money was already received, so it's set automatically and shown as a badge; only
  `Pending`/`Cancelled` remain a dropdown choice.
- **Paid/Partially Paid invoices must have their payment reversed before Cancel or Bad
  Debt.** Both actions check this and, if blocked, show "This invoice has payments
  recorded. Reverse the payments before cancelling." with a Reverse Payment action in the
  Danger Zone instead — it resets the balance to the full total, clears the payment
  record, and reopens the invoice as `Issued` (or `Overdue` if past due date).

### Subscription ↔ Expense linking

- Each subscription record carries a `periods` dict keyed by `YYYY-MM`
  (`{"2026-08": {"expense_id", "receipt_attached", "paid", "linked_at"}}`) — this is the
  billing-period ledger, not the rolling `next_charge_date`/`last_posted_date` fields
  (which only ever describe the single next/most-recent charge).
- Saving an expense whose supplier matches an active subscription (`_match_subscription_by_supplier`)
  tags it `Expense Type: Subscription` + `Subscription ID`, and links it into that
  period's `periods` entry.
- If the matched period already has an expense linked with `Status == "Auto-posted"`
  (i.e. the subscription auto-sync job created a placeholder expense before a real
  receipt arrived), the new save **updates that same row in place** rather than creating
  a duplicate — the stable `EXP-NNN` id is what makes this possible.
- `next_charge_date` only advances when the period being linked is the one currently
  due. Relinking a past period (e.g. attaching a receipt after the fact) must never push
  future billing dates forward.

## Client-side state: single source of truth

- **A value must live in exactly one real form element — never a hidden field mirrored
  by a separate visible control.** If a value needs to be both edited compactly (e.g. a
  button group) and shown/edited elsewhere (e.g. a review card), both UIs must read from
  and write to the *same* underlying `<input>`/`<select>`, not two elements kept in sync
  by JS. Syncing two copies of the same state is exactly the kind of bug that is easy to
  introduce and hard to spot — "it's set in the button group but the other view still
  shows the old value" is a symptom of this, not a one-off bug to patch around.
- Concretely: on the expense form, `expense_type_field` is a single real `<select>`
  (visually hidden inside the full form, visibly rendered on the OCR review card via
  `buildReviewRow(..., "select", ...)`, which clones its `<option>`s and wires `change`
  back onto the same element). The type-selector buttons in the full form and the OCR
  auto-detection logic both write to this one element; neither reads or writes a second,
  parallel field. Form submission always reads this same element regardless of which
  view was last visible, so there is nothing to "propagate" between views.
- When adding a field that needs to appear in more than one place (a compact form and a
  review/summary card, a table row and its edit modal, etc.), reach for this pattern
  first rather than inventing a hidden-field-plus-sync approach.

## Component Patterns — OCR Review Card

The expense-add flow is receipt-first: Upload → OCR Review Card → (edit in) Full Form.
The review card (`#expense_review_step` in `templates/index.html`) has its own
established look, distinct from a standard form section:

- White card, 1px `#E2E2E3` border, gold `#B08D57` left accent border, rounded corners.
- Navy `#16294A` header bar across the top with the card title in white and a confidence
  badge on the right (gold pill ≥ 80%, muted gold 60–79%, red < 60% — there is no green
  in the brand palette, so "good" confidence is gold, not green).
- Body uses a two-column row layout (`review-card-row`): steel-blue uppercase label on
  the left, value on the right, thin bottom border separating rows. Edit pencils (gold)
  only appear on hover, keeping the default view uncluttered.
- A field OCR could not read shows an inline amber pill ("Not found — please add") that
  expands into an editable input on click — never a large red banner for a merely-missing
  optional field. Reserve red exclusively for values that are genuinely blocking (e.g. a
  field the record cannot be saved without).
- Total Amount is the one deliberately oversized value (gold, bold, 24px) — it is the
  number the user is most likely to sanity-check before saving.
- Footer is a distinct light-grey band below a divider: Cancel (left) / Edit in full form
  (left, next to Cancel) / Save (right, primary navy+gold, disabled with an explanatory
  steel-blue hint until required fields — e.g. Payment Method — are set).
- Every dropdown or editable value on the card is bound to the *same* real form field the
  full form uses (see Single Source of Truth above) — the card is a view, not a second
  copy of the data.
