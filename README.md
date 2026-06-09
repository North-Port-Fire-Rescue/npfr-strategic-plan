# NPFR Strategic Plan — `/splan`

A static, single-file web application for viewing, navigating, and exporting the North Port Fire Rescue Strategic Plan. Hosted on GitHub Pages. No backend, no build process, no dependencies to install.

---

## Repository structure

```
/splan
├── index.html        # The entire application — UI, data, and logic in one file
└── README.md         # This file
```

That's intentionally it. The data lives inside `index.html` as a JavaScript object near the top of the file. There is no separate database, API, or CMS.

---

## Hosting (GitHub Pages)

1. Create a repository named `splan` under your GitHub account or the NPFR organization.
2. Push `index.html` and `README.md` to the `main` branch.
3. Go to **Settings → Pages → Source** and set it to `main` branch, `/ (root)`.
4. The plan will be live at:
   ```
   https://<your-username>.github.io/splan
   ```
5. Every `git push` to `main` redeploys automatically within ~30 seconds.

---

## Editing the data

All plan content lives in a single JavaScript object at the top of `index.html`, clearly marked with a comment block:

```
/* ============================================================
   STRATEGIC PLAN DATA — edit here
   ============================================================ */
const PLAN = { ... };
```

### Adding or editing a goal

Find the relevant pillar array inside `PLAN.pillars`. Each goal follows this shape:

```javascript
{
  id: "P1.G3",
  year: "FY27",
  statement: "Develop and implement a major incident response planning program...",
  predecessor: "P1.G1",   // Goal ID this depends on, or null
  owner: "Scott Lane",
  tasks: [
    "Gather incident data from the last 36 months",
    "Separate and rank incident types",
    "Review current response plans for gaps",
    "Create a committee and formulate an implementation plan"
  ],
  metric: "All identified incident types have a documented response plan reviewed by command staff"
}
```

Change any field value and save. That's the entire edit process.

### Adding a new pillar

Add a new object to the `PLAN.pillars` array following the same shape as the existing six pillars. The UI will pick it up automatically.

---

## Live meeting workflow

For in-meeting updates without pushing to GitHub:

1. Open `index.html` directly in Chrome or Firefox (`File → Open`, or drag the file into the browser).
2. Edit the data object in any text editor — VS Code recommended.
3. Save the file and refresh the browser tab. Changes appear immediately.
4. When the meeting is done, `git push` to publish the updated version publicly.

**VS Code Live Server** (optional but useful): Install the Live Server extension. Right-click `index.html` → *Open with Live Server*. The browser will auto-refresh on every save — no manual refresh needed during a meeting.

```bash
# Install Live Server via VS Code Extensions panel, or via CLI:
code --install-extension ritwickdey.LiveServer
```

---

## Excel export

A **Download Excel** button in the application header generates a formatted `.xlsx` workbook from the current data state using SheetJS (loaded from CDN, no install required). The workbook contains:

- A **Summary** sheet with the mission, vision, values, and pillar overview
- One sheet per pillar with all goals, predecessors, tasks, owners, and metrics
- A **Dependencies** sheet listing all goal-to-goal dependency chains

The export reflects whatever data is currently loaded — if you have local edits that haven't been pushed, the download captures those edits.

---

## Dependency chain logic

Each goal carries a `predecessor` field containing the Goal ID of the gate goal that must be completed first, or `null` if the goal has no dependency. The UI uses this field to:

- Display a **"Gates: P#.G#, P#.G#"** badge on gate goals showing what they unlock
- Display a **"Requires: P#.G#"** badge on dependent goals
- Render a dependency diagram view showing the full chain across all pillars

---

## Making a structural change (adding a field, changing layout)

The data schema and the UI rendering are both in `index.html`. Search for the section marked:

```
/* ============================================================
   UI RENDERING — edit below here for layout/style changes
   ============================================================ */
```

Changes above that line are data. Changes below are presentation. Keeping them mentally separate makes edits easier to reason about.

---

## No internet? No problem

The application loads SheetJS for Excel export from a CDN. If you are presenting from a location with no internet access:

- All viewing and navigation works fully offline — the data and UI are self-contained in `index.html`.
- The Excel export button will fail silently if the CDN is unreachable.
- To make export work offline, download SheetJS (`xlsx.full.min.js`) and update the `<script src>` tag to point to the local file path.

```html
<!-- Replace this: -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

<!-- With this (file placed alongside index.html): -->
<script src="xlsx.full.min.js"></script>
```

---

## Git workflow reference

```bash
# Clone the repo
git clone https://github.com/<your-username>/splan.git
cd splan

# Make edits to index.html, then publish
git add index.html
git commit -m "Update P2.G3 owner and tasks"
git push origin main

# The live site updates within ~30 seconds
```

---

## Maintainer

North Port Fire Rescue — Quality Assurance / Quality Improvement  
Questions about the data: contact the QA/QI Officer.  
Questions about the code: the entire application is in `index.html` — open it in any text editor.