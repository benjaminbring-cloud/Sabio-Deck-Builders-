# Weekly proactive strategy decks

Each week, scan Ben Bring's work email for live client and agency openings, pick the best one or two, and build proactive decks modeled on the Sabio × Walmart Sparky deck.

## Where things live

| What | Where |
|---|---|
| Work email | Outlook, `bbring@sabio.inc` (Microsoft 365 connector). The Gmail account is personal newsletters, so skip it. |
| Reference deck | OneDrive `2026/Clients/Publicis/Walmart/Sabio_Walmart_Sparky_AI_Audiences_10.2.26.pptx` |
| Renderer | `lib/sabio-deck.js`: a JSON spec in, a Sabio-branded `.pptx` out |
| Brand | Sabio Brand Guideline (SharePoint `Branding Materials/.../Sabio_Styleguide`): dark grey theme `#1A1A1A`, Sabio yellow `#FFC63E`, Poppins, official logos in `assets/` |
| House style | Ben's 2026 decks: `Sabio x APEX 9.17.26`, `Sabio_Kinesso 9.15.26 V2`, `Sabio_Walmart_Sparky_AI_Audiences_10.2.26` |
| Weekly output | `weeks/<Monday date>/`, holding `strategy-memo.md`, `specs/*.json` and `decks/*.pptx` |
| OneDrive copies | `2026/Clients/<Agency>/<Client>/` for decks, `2026/Clients/Weekly Proactive Strategy/` for the memo. Needs write access; see step 8 |

## Steps

1. **Scan the last 7 days** of Outlook Inbox and Sent Items. Skip newsletters, vendor pitches, internal wins, calendar accepts and automated reports. Keep threads where a client or agency contact:
   - offered access (intros, seat IDs, a meeting), or
   - named a need, a brief or a timeline, or
   - is someone Ben is actively trying to get in front of.
2. **Read the full thread** for each candidate, plus older history on the same client (search the client name).
3. **Research the client's market news** (web search) for the last ~12 months: launches, approvals, earnings, competitive moves, regulation, calendar moments. The deck has to show how Sabio helps them *today*, before they've asked. Every market claim gets a number in the deck that points to the Sources slide.
4. **Search OneDrive/SharePoint** for existing decks on that client or agency. Don't duplicate work in progress. If a deck exists, list it in the memo instead of building a new one.
5. **Rank the openings.** Weigh how warm the contact is, the timing (a season, a launch, a due date), and whether a deck is the right next artifact. Build decks for the top 1–2. The rest go in the memo with one next action each.
6. **Write a spec per deck** in `weeks/<date>/specs/<client>.json`. Copy the slide order from last week's spec: cover (with the one timely stat) → agenda → what's happening at the client (dated news cards, each with "what it means") → the opening (the point of view) → Sabio at a glance → three audiences → reach build → activation → screen-to-action flow → measurement → proof → guardrails → where we go next (three moves with owners and dates) → thank you → sources. Rules:
   - Build audiences from app ownership (four signal buckets, three core audiences: Prime / Conquest / Prove or equivalent).
   - Tag every reach or device count that isn't from a real build `[CONFIRM]`.
   - Never carry another client's pricing or non-public numbers. Use anonymized or published case studies only.
   - Voice moments (short-answer table, "why now", next steps, thanks) follow the `ben-bring-voice` skill. Use short sentences, plain first person, and avoid words like "leverage", "unlock" and "seamless".
   - For pharma/health, don't build condition-level audiences. Say so on the privacy slide.
7. **Render and QA**:
   ```bash
   npm install            # first run only
   # Poppins for true-to-width QA renders: npm pack @fontsource/poppins@4.5.10, convert the latin .woff files to .ttf into ~/.fonts
   node lib/sabio-deck.js weeks/<date>/specs/<client>.json weeks/<date>/decks/Sabio_<Client>_<Topic>_<M.D.YY>.pptx
   python3 <pptx skill>/scripts/office/validate.py <deck>
   soffice --headless --convert-to pdf <deck> && pdftoppm -jpeg -r 60 <deck>.pdf slide
   ```
   Look at every slide image for overflow, overlap and empty cards. Fix the spec or the renderer and re-render.
8. **Write `strategy-memo.md`** with the decks built, other openings ranked with next actions, and a checklist of what to confirm before sending.
9. **Deliver**:
   - Commit `weeks/<date>/` and push.
   - End the run with a short summary: the decks built, their GitHub links, and the ranked openings.
   - Upload the decks and memo to OneDrive only if the Microsoft 365 connector has `Files.ReadWrite.All`. As of 9/30/26 it is read-only (`Files.Read`, `Mail.Read`), so the repo is the delivery point.
   - **Never email or contact clients.** Everything is a draft for Ben.
