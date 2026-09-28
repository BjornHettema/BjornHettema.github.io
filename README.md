# bjornhettema.github.io

Personal site for Bjorn Hettema — spatial planner working across urban water,
mobility and urban data in the Netherlands and Southeast Asia.

`index.html` is the entire site: one self-contained file, no build step, no
dependencies. Fonts load from Google Fonts; everything else is inline.

## Structure

```
index.html          the site
404.html            not-found page
me.jpg              portrait
favicon.svg         site icon
robots.txt          crawler rules
sitemap.xml         one entry, the home page
cv/                 two CV variants (.docx), web copies without a phone number
scripts/build-cv.js generates all four CV files from one content model
```

## Publishing

**GitHub Pages:** this repo is named `BjornHettema.github.io`, so it is a GitHub
*user site* and serves from the root of `main` at https://bjornhettema.github.io.
Settings → Pages → deploy from branch `main`, folder `/ (root)`.

A user site must live in a **public** repository on a GitHub Free account;
Pages from a private repo requires GitHub Pro.

**Custom domain:** add a `CNAME` file at the root containing the bare domain,
then point a CNAME record at `<username>.github.io`.

**Anywhere else:** upload `index.html` and `cv/`. That is the whole deployment.

## The two CVs

- `cv/CV-BjornHettema-urban-water-research.docx` — leads with research and
  fieldwork. For water, climate and consultancy roles.
- `cv/CV-BjornHettema-urban-data-geospatial.docx` — leads with the technical
  work. For geospatial and urban-tech roles.

Same evidence in both, different emphasis and ordering.

The files in `cv/` are the **web copies** and carry no phone number, because this
repository and the site are public. `scripts/build-cv.js` also writes
`*-full.docx` versions that do include it — those are the ones to attach to an
email, and they are deliberately not committed here.

Regenerate all four:

```bash
npm install docx
node scripts/build-cv.js
```

## Changing the portrait

`me.jpg` sits beside `index.html` and is referenced from `<div class="avatar">`.
Swap the file, keeping it square, and nothing else needs touching.

## Note on credentials

Nothing in this repository should ever hold a token. `.gitignore` covers the
usual suspects; the real rule is that credentials live in a password manager,
never in a file or a commit.
