# Store screenshots

AppSource listing images. **Specs:** PNG, **1366×768** or 1280×720, 1–5 images,
no excessive whitespace, show the add-in actually doing something.

The two files here (`01-overview.png`, `02-taskpane.png`) are landing-page renders
— fine as placeholders, but replace them with **real in-Excel captures** before a
public listing (they convert far better and reviewers expect them).

## Shot list (capture these in Excel with the add-in loaded)

1. **Functions in the grid** — a small table with `=LLM.CLASSIFY`, `=LLM.FIELDS`,
   and `=LLM.SUMMARIZE` visibly returning values in adjacent columns. The formula
   bar should show one `=LLM.…` call. *This is the money shot — make it #1.*
2. **Task pane / settings** — the **LLM Settings** pane open, a provider selected,
   model filled. Shows the BYO-provider story.
3. **Agent** — the Agent box with a plain-English instruction typed and the
   **Apply N changes** button visible (approve-before-apply). Shows the agent
   editing the sheet safely.
4. *(optional)* **MAP over a column** — `=LLM.MAP(A2:A20, "translate to German")`
   spilling results down a column.

Keep filenames zero-padded and ordered (`01-…`, `02-…`); AppSource shows them in
that order. Update the paths in `APPSOURCE.md` Step 3 only if you rename them.

## Regenerate the placeholder renders

The current placeholders are produced from `../../src/site/index.html` at
1366×768. Any headless-Chrome screenshot of that page at that viewport reproduces
them; there is no committed generator script (kept out to avoid a Puppeteer dep).
