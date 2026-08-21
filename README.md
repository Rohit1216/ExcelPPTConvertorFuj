# LinkedIn Team → PowerPoint Generator

Upload an Excel workbook (one sheet per team) and get back a branded
PowerPoint deck: one slide grid of "cards" per person, name hyperlinked to
their LinkedIn profile, and photo pulled from a `PhotoUrl` column
(falls back to a generated placeholder avatar if the URL isn't a real image
or is missing).

## Expected Excel format

Each **sheet** = one team/section (its tab name becomes the slide title).
Header row, any order, case-insensitive:

| Name | Phone | Email | Designation | LinkedinUrl | Location | PhotoUrl |
|---|---|---|---|---|---|---|

Only `Name` and `Designation` are required — everything else is optional.
See `sample_input/NATO_example.xlsx` for a working example.

## Repo layout

```
app.py                     Streamlit UI
utils/ppt_generator.py     All the PPTX-building logic (reusable outside Streamlit too)
requirements.txt
sample_input/               Example workbook to try the app with
```

## Run it locally

```bash
git clone https://github.com/<your-username>/<your-repo>.git
cd <your-repo>
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

This opens `http://localhost:8501`. Upload an `.xlsx`, click **Generate
PPTX**, and download the result.

## Step-by-step: put this on GitHub

1. Create a new empty repository on GitHub (no README/license — you already
   have files), e.g. `linkedin-deck-generator`.
2. On your machine, in this folder:
   ```bash
   git init
   git add .
   git commit -m "Initial commit: LinkedIn team deck generator"
   git branch -M main
   git remote add origin https://github.com/<your-username>/linkedin-deck-generator.git
   git push -u origin main
   ```
3. Refresh the GitHub page — your files should now be there.

## Step-by-step: deploy for free on Streamlit Community Cloud

1. Go to **https://share.streamlit.io** and sign in with your GitHub account.
2. Click **"Create app"** → **"Deploy a public app from GitHub"**.
3. Pick your repository, branch `main`, and main file path `app.py`.
4. Click **Deploy**. The first build takes a couple of minutes (installing
   `requirements.txt`).
5. You'll get a public URL like `https://your-app-name.streamlit.app` —
   share it with anyone; they can upload their own Excel file and download
   the generated deck, no installation needed.

Whenever you `git push` new commits to `main`, Streamlit Cloud
auto-redeploys.

## Customizing the look

Open `utils/ppt_generator.py` — the constants at the top control the whole
theme:

- `HEADER_COLOR_LEFT` / `HEADER_COLOR_RIGHT` — header gradient
- `CARD_FILL` / `CARD_BORDER` — regular person-card colors
- `FEATURED_FILL` — the big highlighted card (e.g. for a CEO/exec)
- `GRID_COLS`, `GRID_ROWS_NORMAL` — grid density per slide (default 3×3 = 9
  people/slide; extra people automatically flow onto "(2/2)", "(3/3)" ...
  continuation slides)

If you want the output to match an exact company template (fonts, logo,
exact positions), upload that `.pptx` in the sidebar's "base template"
field — the generator will build slides using its theme/master. For a
pixel-perfect clone of a specific template's card layout beyond colors,
send that template file and this generator logic to Claude and ask for the
positions to be matched exactly.

## Notes on photos

The generator now tries harder before giving up on a person's photo:

1. **`PhotoUrl` as a direct image** — plain HTTP GET, works for any
   `.jpg`/`.png`/CDN media link (e.g. `media.licdn.com/...`).
2. **If `PhotoUrl` (or `LinkedinUrl`) is a `linkedin.com` link**, that step
   above will fail (a profile page isn't an image file), so the app fetches
   that page's HTML and pulls out its `og:image` tag — the same photo
   LinkedIn itself uses for link-preview cards — and downloads *that*.
3. **If everything fails**, a clean generated placeholder avatar is used
   instead. The app never crashes because of a bad or blocked photo link.

**Caveat:** step 2 is best-effort. LinkedIn increasingly serves a
login-wall page to anonymous/non-browser requests, in which case
`og:image` may be LinkedIn's own logo rather than the person's actual
photo, or the request may be blocked/rate-limited outright — especially if
you run this many times in a row. There's no reliable way around this
without logging in (which this app deliberately does not do, to stay
within LinkedIn's terms of use). For guaranteed real photos, point
`PhotoUrl` directly at an image file (company headshot CDN, Gravatar,
etc.) rather than relying on the LinkedIn fallback.

Downloaded images are cropped to a circle automatically, and results are
cached per-URL for the duration of a single generation run so the same
person's photo isn't fetched twice.
