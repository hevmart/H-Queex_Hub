# H-Queex Hub — Session Handover

Last updated: 2026-08-11. Use this to brief a new chat session cold — it should not
need to re-derive any of this from git history.

## What H-Queex Hub is

A Flask + JSON-file business management app for H-Queex, an Irish sole-trader/limited
company consultancy. No database — every entity is a flat JSON file (or a "sheet" JSON
file for spreadsheet-style entities), read/written directly. Single-user, runs locally
via `Launch-HQueex-Hub.ps1` on `http://127.0.0.1:5000`.

**Test suite: 206 tests passing** (`tests/test_writeback.py`, run via
`.venv/Scripts/python.exe -m pytest -q`). Every feature session so far has ended with
the full suite green before committing — keep that bar.

**App runs with `debug=False`** (`app.run(debug=False, use_reloader=False, ...)`), so
**Jinja template auto-reload is off** — editing `templates/index.html` or `app.py` while
the dev server is running requires a full process restart before the change is visible,
even though the file on disk is correct immediately. This caused real confusion across
this session (screenshots looking "unchanged" after a fix) — always restart after
template/app edits before re-checking in a browser. Also watch for **stale duplicate
`python app.py` processes** accumulating from repeated dev-server restarts (only one
ever actually owns port 5000 — `Get-NetTCPConnection -LocalPort 5000 | Select
OwningProcess` finds the real one; kill the rest via `taskkill /PID <n> /F`).

## Modules built (chronological)

1. **Company** — Documents, Compliance Calendar (auto-calculated Irish filing deadlines),
   Business Profile, Settings.
2. **Operations** — Projects (Kanban, filterable by status/tier/client), DMAIC Tracker
   (phase-sequenced, transitions to Complete are server-enforced in order), Delivery Log
   (with Clarity Partner invoice generation pulling all uninvoiced entries for a billing
   period into a Draft invoice), SOP Library (version-controlled, Draft → Review →
   Approved workflow is server-enforced, cannot skip a step).
3. **CRM** — Lead pipeline (Kanban, cards colour-coded green/amber/red by next-action-date
   severity), Lead detail page (`/crm/leads/<id>` — activity timeline, linked proposals,
   Convert-to-Client button on Won leads), Proposals (line items + service catalogue,
   expiry warnings amber <7 days / red expired), Convert-to-Invoice, public API for
   website enquiry intake.
4. **Stable IDs foundation** — every entity given a `PREFIX-NNN` id (see
   `docs/ui-standards.md` → Data Standards for the full prefix table and the rules
   around join keys / idempotent migrations). Most recently added: `EXP` for Expenses.
5. **Finance core** — Income, Expenses (filter bar + live summary panel + bulk
   archive/CSV export, all client-side JS, no page reload), Invoices (branded print
   preview at `/invoices/preview/<row_number>`, configurable `invoice_number_offset` on
   the business profile), Payroll, Subscriptions, VAT3 export, Ledger/journal, bank
   reconciliation, capital allowances, **Year-end pack generator** (`/finance/year-end-pack/generate`
   — downloads a branded ZIP: P&L by category, VAT summary per bi-monthly period,
   pre-trading expense schedule, capital allowances schedule, full transaction CSV; each
   generation is logged to `filing-log.json` and the Finance/Ledger pages warn if
   transactions have changed since the last pack for that year).
6. **Receipt OCR** — Claude vision (`claude-sonnet-4-6`) reads uploaded receipts/invoices
   (PDF via `pdf2image` + a locally vendored Poppler binary, or image formats directly)
   and drives a receipt-first Add Expense flow: **Upload → OCR Review Card → (optionally)
   Full Form**. See Architectural Decisions below for the details worth knowing before
   touching this code.
7. **Dashboard** — Upcoming Actions (overdue/due-soon invoices, next 3 compliance
   deadlines, Clarity Partner unbilled delivery, project deadlines, lead follow-ups — all
   clickable), Year-to-date summary (income/expenses/net profit/estimated tax using the
   existing phase-aware `PHASE_POLICY` rate, current VAT position).

## Architectural decisions worth knowing

- **Stable ids, idempotent migrations.** Every entity gets a `PREFIX-NNN` id assigned
  once, never reassigned. Migrations that backfill missing/legacy ids run on every
  `load_finance_data()` call and must be no-ops once complete. Follow the existing
  pattern (`_migrate_prefixed_ids`, `_generate_prefixed_id`) rather than inventing a new
  one for any new entity.
- **Sheet vs JSON-list entities.** Income/Expenses/Invoices/Clients/Suppliers are
  "sheets" (`SHEET_JSON_PATHS`, loaded via `_load_sheet_records_raw` /
  `_save_sheet_records_raw`, positionally indexed by `__row_number` for
  edit/delete but carrying a stable `id` too now). Everything else (Projects, SOPs,
  Subscriptions, Documents, Leads, Proposals, Services, etc.) is a plain JSON list file
  loaded via dedicated `_load_*`/`_save_*` functions. Know which kind you're touching —
  the helper functions are not interchangeable.
- **Single source of truth for client-side state** (newly documented in
  `ui-standards.md`, born from a real bug this session): a value must live in exactly
  one real form element. Never mirror a value into a hidden field kept "in sync" with a
  separate visible control by JS — bind every UI that needs to show/edit that value to
  the *same* DOM element. The expense-type selector bug (button group silently
  disagreeing with a hidden field) took three attempts to fix properly because of this
  anti-pattern; the fix converted the hidden `<input>` into a real (visually hidden)
  `<select>` that both the button group and the OCR review card's dropdown read/write
  directly.
- **Receipt-first OCR flow**, all within one `<form id="expense_manual_form">`:
  - Step 1 (`#expense_upload_step`): big dropzone + mobile "Take Photo", "Enter manually
    instead" link. Shown by default for new expenses only (not when editing).
  - Step 2a (`#expense_review_step`): the OCR review card — H-Queex-branded (navy header,
    gold left border), two-column rows, missing-field amber pills (never large red
    banners for merely-optional-but-missing fields), inline supplier search/quick-add,
    Expense Type and Status as real visible dropdowns, Save disabled until Payment
    Method is set. Full component spec is in `ui-standards.md`.
  - Step 2b: the full form (all sections), reachable via "Edit in full form" from the
    review card, or directly via "Enter manually instead" from Step 1.
  - Backend: `POST /expenses/ocr` runs the vision call and also does two pieces of
    business logic client can't: `_match_subscription_by_supplier()` (subscription
    auto-detection) and `_find_duplicate_expense_for_period()` (duplicate warning).
    Both are returned in the JSON (`subscription_match`, `duplicate_warning`) for the
    client to render — the client never re-derives this logic itself.
- **Subscription ↔ Expense linking.** Subscriptions carry a `periods` dict
  (`{"2026-08": {expense_id, receipt_attached, paid, linked_at}}`) as the real
  billing-period ledger — `next_charge_date`/`last_posted_date` only describe the single
  next/most-recent charge, they are not a history. Saving an expense that matches an
  active subscription links it into that period; if an `Auto-posted` placeholder expense
  already exists for that period (created by the subscription sync job), the real save
  **updates that row in place** instead of creating a duplicate. `next_charge_date` only
  advances when the period being linked is the one currently due.
- **Backend-authoritative business logic.** Where a UI needs a computed value that has
  real business rules behind it (Phase Tag from a date, subscription matching, duplicate
  detection), the server computes it and the client fetches/reads it — the client does
  not reimplement the rule in JS. (Phase Tag was briefly shown on the review card via a
  small `/api/expenses/phase-tag` endpoint; removed from the UI per explicit request,
  but the pattern — and the endpoint — still stands as the template for anything similar.)
- **Two data-model gaps accepted, not solved**: subscriptions have no `currency` or
  `phase_tag` field in the schema. When a real Anthropic "Claude Pro" subscription was
  added this session, those two values were put in the `notes` field as text rather than
  silently dropped or the schema being expanded ad hoc — flagged to the user, not yet
  decided whether to promote them to real fields.
- **Modal width follows field count, not a fixed default.** `ui-standards.md` now
  requires modals with more than 6 fields to be ≥800px wide with a multi-column layout;
  narrow (`.modal-box`/`.form-modal-box`, 380/560px) modals are for simple confirmations
  only. The wide pattern is the `.wide-modal-box` modifier (900px, `min(900px, 85vw)`)
  combined with `.service-modal-box` (tightened field/row spacing, 90vh scroll safety
  net) — used by Lead, Proposal, and Service. Other wide forms (Business Profile,
  Project, Delivery, SOP, Client) reuse `.wide-modal-box` on top of the existing
  `.form-modal-box` base rather than switching base classes, since `.form-modal-box` sets
  nothing but `max-width`. When adding a new modal with >6 fields, add `wide-modal-box` to
  its class list rather than inventing a new width value.
- **`fmt_date` Jinja filter** (`app.py`, registered via `app.jinja_env.filters["fmt_date"]`)
  converts stored ISO (`YYYY-MM-DD`) dates to `DD/MM/YYYY` for display. Applied to every
  list-view/detail-view date display across the app. Never apply it to `<input
  type="date">` `value=` bindings — the native date picker requires ISO format. Field key
  names for "the date" are **not consistent across sheets** — Income uses `Date`,
  Expenses uses `Date (Registered)` — check the actual JSON before assuming a key name
  (this caused a real bug: the dashboard's YTD expense total silently read the wrong key
  and always showed €0.00 until caught and fixed).
- **Duplicate section headings are a recurring failure mode, not a one-off bug.** A
  generic fallback block in `index.html` (`{% if active_tab not in (finance_tabs +
  company_tabs + operations_tabs + crm_tabs + [...]) %}<h2>{{ page_title }}</h2>{% endif %}`)
  renders a page-title `<h2>` for any `active_tab` not explicitly excluded. Every new
  page/tab that already renders its own `<h1>`/title elsewhere (a new PDF/print preview
  tab, a new sub-view) **must** be added to that exclusion list, or it silently gets a
  duplicate heading. This bug recurred three times this session (CRM hub, invoice
  preview, proposal PDF) before the pattern was recognized — check this list first
  whenever adding a new `active_tab`.
- **Year-end pack** (`_build_pnl_by_category`, `_build_year_vat_period_summaries`,
  `_build_pretrading_schedule`, `_render_branded_report_html` in `app.py`) generates a
  ZIP in-memory (`zipfile.ZipFile` + `BytesIO`) with the logo embedded as base64 in each
  HTML report so the files are portable outside the running server. VAT period math
  reuses `_compute_vat_summary_for_bounds` (extracted from `_compute_vat_control_summary`,
  which now just calls it for the current period) — use this shared function for any
  future per-period VAT reporting rather than re-deriving the T1/T2/T3 loop.
- **`filing-log.json` is a new top-level data file** — added to the `isolated_subscription_file`
  autouse test fixture's path-redirection list (`tests/test_writeback.py`). Any new
  top-level `*_PATH` constant added to `app.py` **must** be added to that fixture's
  save/redirect/restore triple, or tests will read/write the real project data file
  instead of an isolated tmp copy (this happened once this session before being caught).

## Known issues / things a new session should be aware of

- **OCR requires local setup per machine**: `ANTHROPIC_API_KEY` in `.env` (gitignored,
  never committed — a placeholder-only `.env.example` is committed instead), and a
  vendored Poppler binary at `vendor/poppler/` (gitignored, downloaded manually as a
  portable zip because `choco install poppler` failed on this machine due to permissions
  — see any prior session transcript for the exact download URL if it needs redoing on a
  new machine).
- **No test exercises the real Claude API** — all OCR tests mock `_run_receipt_ocr`.
  This is intentional (cost, determinism, no network dependency in CI) but means a
  prompt regression that changes Claude's actual output shape wouldn't be caught by the
  suite — only by live testing.
- **Travel/Mileage and "Other" expense types are built but lightly tested.** The
  Subscription and Receipt/Invoice paths have been exercised heavily (including live,
  with a real Anthropic invoice); the mileage calculator (km × €0.43 Revenue rate,
  journey log, "actual costs instead" toggle) has unit-level test coverage but hasn't
  had the same live back-and-forth scrutiny.
- **Real `subscriptions.json` currently has 3 entries**: two old test/sample records
  ("HQ Sample Subscription", "HQ Subscription Form", supplier "HQ Supplier") and one real
  one added this session ("Claude Pro", supplier "Anthropic, PBC", used to prove the
  subscription-auto-detect flow end-to-end). The two sample ones are probably safe to
  archive/delete once no longer needed for testing — nobody has asked for that yet.
- **`docs/ui-standards.md` now has two explicitly-flagged deviations from its own base
  rules** (both intentional, both documented inline where they occur): the Description
  field's placeholder exceeds the normal character-limit rule (Expenses only, since
  Title was removed and Description became the sole primary identifying field); the
  review card's confidence badge uses gold rather than green for "good" (brand palette
  has no green).

## Next priorities (not yet requested, but logical follow-ons)

- Decide whether `currency` and `phase_tag` become real subscription fields, or stay as
  notes-text conventions.
- Broaden live/manual testing of the Travel and Subsistence and Other expense-type paths
  the same way Subscription/Receipt were tested.
- Consider whether the "receipt-first OCR" pattern (Upload → Review Card → Full Form)
  and its single-source-of-truth lesson should be applied anywhere else in the app
  (Income? Invoices?) — no request for this yet, but the pattern is now proven out.
- Decide what to do with the two sample subscription records in production data.
- The Task 10 code-quality audit added a single global `@app.errorhandler(500)` (branded
  friendly error page + server-side logging) rather than auditing and wrapping every
  individual POST route in its own try/except — this satisfies "no unhandled 500 without
  a friendly message" but a route-by-route audit for *silent* failure modes (e.g. a
  write that partially succeeds) hasn't been done.
- `_export_ledger_journal_csv` filtered by year is currently the "all transactions" file
  in the year-end pack. If a future accountant workflow wants separate Income/Expense/
  Invoice CSVs instead of one ledger-journal-shaped CSV, that's a small addition to
  `generate_year_end_pack()`, not a redesign.
- Real-world verification of the year-end pack's VAT/P&L numbers against an actual filed
  return hasn't happened yet — the calculations reuse existing, already-tested helpers
  but the *assembled* pack itself has only been smoke-tested (`test_year_end_pack_generates_zip_with_expected_files`),
  not reviewed by an accountant.
