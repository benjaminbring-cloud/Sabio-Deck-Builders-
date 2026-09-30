# Weekly proactive strategy decks

Each week, scan Ben Bring's work email for live client and agency openings, pick the best one or two, and build proactive decks modeled on the Sabio × Walmart Sparky deck.

## Where things live

| What | Where |
|---|---|
| Work email | Outlook, `bbring@sabio.inc` (Microsoft 365 connector). The Gmail account is personal newsletters, so skip it. |
| Reference deck | OneDrive `2026/Clients/Publicis/Walmart/Sabio_Walmart_Sparky_AI_Audiences_10.2.26.pptx` |
| Renderer | `lib/sabio-deck.js`: a JSON spec in, a `.pptx` out |
| Weekly output | `weeks/<Monday date>/`, holding `strategy-memo.md`, `specs/*.json` and `decks/*.pptx` |
| OneDrive copies | `2026/Clients/<Agency>/<Client>/` for decks, `2026/Clients/Weekly Proactive Strategy/` for the memo |

## Steps

1. **Scan the last 7 days** of Outlook Inbox and Sent Items. Skip newsletters, vendor pitches, internal wins, calendar accepts and automated reports. Keep threads where a client or agency contact:
   - offered access (intros, seat IDs, a meeting), or
   - named a need, a brief or a timeline, or
   - is someone Ben is actively trying to get in front of.
2. **Read the full thread** for each candidate, plus older history on the same client (search the client name).
3. **Search OneDrive/SharePoint** for existing decks on that client or agency. Don't duplicate work in progress. If a deck exists, list it in the memo instead of building a new one.
4. **Rank the openings.** Weigh how warm the contact is, the timing (a season, a launch, a due date), and whether a deck is the right next artifact. Build decks for the top 1–2. The rest go in the memo with one next action each.
5. **Write a spec per deck** in `weeks/<date>/specs/<client>.json`. Copy the slide order from last week's spec: cover, contents, table, then four sections (audience, activation, measurement, plan), next steps, thanks. Rules:
   - Build audiences from app ownership (four signal buckets, three core audiences: Prime / Conquest / Prove or equivalent).
   - Tag every reach or device count that isn't from a real build `[CONFIRM]`.
   - Never carry another client's pricing or non-public numbers. Use anonymized or published case studies only.
   - Voice moments (short-answer table, "why now", next steps, thanks) follow the `ben-bring-voice` skill. Use short sentences, plain first person, and avoid words like "leverage", "unlock" and "seamless".
   - For pharma/health, don't build condition-level audiences. Say so on the privacy slide.
6. **Render and QA**:
   ```bash
   npm install            # first run only
   node lib/sabio-deck.js weeks/<date>/specs/<client>.json weeks/<date>/decks/Sabio_<Client>_<Topic>_<M.D.YY>.pptx
   python3 <pptx skill>/scripts/office/validate.py <deck>
   soffice --headless --convert-to pdf <deck> && pdftoppm -jpeg -r 60 <deck>.pdf slide
   ```
   Look at every slide image for overflow, overlap and empty cards. Fix the spec or the renderer and re-render.
7. **Write `strategy-memo.md`** with the decks built, other openings ranked with next actions, and a checklist of what to confirm before sending.
8. **Deliver**:
   - Commit `weeks/<date>/` and push.
   - Upload the decks and memo to OneDrive (paths above).
   - Email Ben a short summary with the OneDrive links. **Never email clients.** Everything is a draft for Ben.
