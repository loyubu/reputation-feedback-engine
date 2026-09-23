# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A course capstone: a **Reputation & Feedback Intelligence Engine** for a multi-location car
care and auto repair chain, to be built on **n8n**. Group project, graded on a working
end-to-end demo rather than production hardening.

**Nothing has been built yet.** There are no workflows, no datastore, no credentials, no n8n
instance. This directory currently holds the client brief, a design review, and the sources
that generate it. Do not assume code exists.

`Feedback-Engine-Design.pdf` is the agreed design; read it before proposing changes to the
system's shape.

**The client brief is deliberately not in this repository.** `Capstone Project Information.txt`
is the final authority on requirements and is kept locally, outside version control, because
this repo is public and the brief is the course's material. It sits beside this file in the
working directory. If it is absent, the design PDF carries the requirements it was derived from
— but for anything contested, ask for the brief rather than inferring it.

## Commands

The only build here produces the two review documents from `docs-src/`.

```bash
cd docs-src
npm install                     # once — pptxgenjs 4.0.1

node deck.js                    # deck  -> ../Feedback-Engine-Design.pptx
```

PDF (headless Chrome; there is no LibreOffice on this machine):

```bash
"/c/Program Files/Google/Chrome/Application/chrome.exe" --headless=new --disable-gpu \
  --no-pdf-header-footer --print-to-pdf="<abs>/Feedback-Engine-Design.pdf" \
  "file:///<abs>/docs-src/design.html"
```

Both documents are generated. **Edit `docs-src/design.html` or `docs-src/deck.js` and rebuild
— never hand-edit the PDF or PPTX**, or the next rebuild silently discards the change.

### Verifying document changes

There is no LibreOffice or `pdftoppm`, so the usual render paths do not work. What does:

```bash
# rasterise any PDF for visual inspection (pip install pymupdf)
python -c "import pymupdf;d=pymupdf.open('x.pdf');[p.get_pixmap(dpi=110).save(f'pg-{i+1}.png') for i,p in enumerate(d)]"
```

For the deck, **PowerPoint itself is installed and is the most faithful renderer available** —
export via COM (`$pp = New-Object -ComObject PowerPoint.Application`, `SaveAs($pdf, 32)`) then
rasterise as above. This is worth doing: it has already caught shape-overlap and
arrow-misalignment defects that passed structural validation.

## The design, and the parts that are easy to get wrong

These are the points a reader of the brief alone reliably misses. They are decided; do not
quietly re-litigate them.

**Four routes, not three.** The brief's flow section lists three, but point 8 adds a fourth:
anything that cannot be confidently scored is flagged for a human rather than guessed at. It is
a first-class route with its own queue, not an error handler.

**Scoring is two axes.** Polarity *and* severity, plus a confidence value. Collapsing them into
one number is precisely the flattening the brief objects to.

**Customer history is an override, not an input.** A repeat complainer escalates "regardless of
how mild this particular message reads" — history bypasses the severity threshold rather than
nudging a score.

**Rule order in the decision node is load-bearing.** First match wins, and the history test sits
*above* the severity test. Reverse them and a mild repeat complaint falls through to the ordinary
queue, which collapses the third demo case into the second.

**The standing rule.** The system never publishes publicly and never sends a *response* to a
customer. It notices, scores, routes and prepares; a human does the rest. Any feature that
auto-posts a review or auto-sends a reply contradicts the brief.

The one exception is deliberate: the **feedback request** that goes out when a visit is marked
complete *is* automatic — the brief requires it (flow steps 1–2). Everything after the customer
replies — drafted responses, testimonials — waits for a human. Do not read the rule as forbidding
the request, or the first workflow never gets built.

### Structural decisions

- Routing logic belongs in **one Code node** emitting a `route` field, followed by a single
  Switch — not a tree of IF nodes.
- Use a **Basic LLM Chain with a structured output parser, not an AI Agent**. One classification
  call; an Agent adds nondeterminism to a demo that must behave identically three times.
- **Two model calls, not one** — score first, draft second. Keeps drafting failures from taking
  the routing with them. Drafting runs on the negative route (manager response) and, once the
  polish node lands, on the positive route (testimonial polish) — a *separate* node with its own
  prompt, not one node switching behaviour on a field. Both continue on error, so a drafting
  failure still writes the record.
- The **polish prompt is constrained to facts present in the original**. The live failure mode is
  not clumsy phrasing, it is the model inventing a branch, a technician or a service the customer
  never mentioned.
- **Every threshold in one Config node** (severity cut-off, confidence floor, history lookback,
  form base URL). None buried in expressions.
- The demo needs a **seeder workflow**, because the repeat-negative case only behaves as a repeat
  if that customer's first complaint was processed first. Ordering cannot be done by hand live.

## The four decisions, settled 2026-09-22

These blocked implementation and are now answered by the group. They are decided; do not reopen
them unilaterally.

1. **Datastore + dashboard** — **Airtable is authoritative.** It holds every processed record and
   is the only store the history override reads. **Sheets is a downstream mirror**, written after
   the Airtable write and never read back by a workflow — it exists for the team view and charts.
   Routing logic reading Sheets is a defect: if the two stores can disagree, the history override
   becomes a coin flip and demo case three collapses into case two.
2. **Request and reply channel** — **n8n Form link** (Form Trigger), URL built from the Config
   node's base URL.
3. **Manager alert channel** — **Slack and email in parallel** off the escalation route. The
   recipient comes from the Config node, **never from the feedback payload** — resolving it from
   the customer's submitted address would email the customer and break the standing rule.
4. **"Ready to post" content** — **both the raw words and a lightly polished draft**, in separate
   fields (`raw_text`, `polished_draft`, plus `variant_used` recording what the approver chose).
   Showing both makes the edit auditable, and gives the positive route a real human decision
   rather than a rubber stamp. **The polish node is deferred**: get the four routes working end to
   end on raw text, then add it. If the trial runs out first, shipping raw-only loses nothing.

## Build constraints specific to this project

Built on **n8n Cloud first, then exported to self-hosted Docker on Oracle Cloud**. That move
shapes decisions now:

- **Credentials do not survive an export** — they are stripped from workflow JSON. The list to
  re-connect: Anthropic, Airtable, Google Sheets, Slack, and SMTP or Gmail. Slack's OAuth is the
  fiddliest to redo.
- **Build the feedback form link from the Config node's base URL**, never hardcoded. The domain
  is already owned, so that value can point at its permanent home from the start.
- **Keep state in the external datastore, not n8n Data Tables**, so the migration moves workflows
  only.
- **Pin the Docker image to the cloud version** — `typeVersion` drift shows up as parameters that
  quietly stop applying.
- Oracle Always Free instances are **arm64**; any sidecar container needs an arm64 image.
- The **n8n Cloud trial is 14 days** — the real schedule constraint, not cost. Whole build runs to
  roughly $5 of Anthropic credit.

n8n's Anthropic node keeps its own model list, which can lag releases. If the intended model is
absent, an HTTP Request node against the Messages API does the same job.
