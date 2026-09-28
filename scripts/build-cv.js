const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink,
  AlignmentType, BorderStyle, TabStopType, TabStopPosition, convertInchesToTwip
} = require("docx");

const ACCENT = "0E6E7D";
const INK = "16232B";
const MUTED = "5E7180";
const FONT = "Arial";

/* ---------------- shared content ---------------- */

const CONTACT_FULL = "contact.hettema@gmail.com  ·  +31 6 46521042  ·  Bangkok, Thailand (Dutch national)";
const CONTACT_WEB  = "contact.hettema@gmail.com  ·  Bangkok, Thailand (Dutch national)";
const LINKEDIN = "linkedin.com/in/bjorn-hettema-02286515b";
const SITE = "bjornhettema.github.io";
const SITE_URL = "https://bjornhettema.github.io/";

const PROFILE_TECH =
  "Spatial planner with a three-year foundation in electrical and computer engineering, working where urban systems meet software. Built a sensor-driven dynamic routing system in Python for a commercial IoT waste product, and map and analyse urban networks in QGIS, Overpass/OpenStreetMap and R. Fifteen months of applied research and fieldwork across Vietnam and Indonesia, documented publicly by both host universities.";

const PROFILE_RESEARCH =
  "Spatial planner specialising in urban water systems, climate adaptation and smart-city applications, with a technical foundation in computer and electrical engineering. Sole author of a comparative study of historical canal and quay-wall systems in Semarang, Jakarta and Surabaya, conducted within the Dutch City Deal “Timeless Canals” framework and graded 4.0. Fifteen months of international fieldwork, published by both host universities.";

const SKILLS_TECH = [
  ["Geospatial", "QGIS · ArcGIS · Overpass Turbo / OpenStreetMap · spatial analysis · suitability & network mapping"],
  ["Programming & data", "Python (routing, data processing, APIs) · R (demographic analysis) · LaTeX"],
  ["Methods", "Vehicle routing & graph algorithms · multi-criteria scoring models · comparative case study · field survey design"],
  ["Domain", "Smart cities & urban IoT · urban mobility · climate adaptation & urban water · Dutch Omgevingswet and CROW standards"],
  ["Certified", "UAS A1/A3 drone operator · TEFL/TESOL Level 5 · CEFR C1 English"],
];

const SKILLS_RESEARCH = [
  ["Research", "Comparative case study design · data triangulation · field survey & scoring instruments · governance mapping"],
  ["Geospatial", "QGIS · ArcGIS · Overpass Turbo / OpenStreetMap · spatial and suitability analysis"],
  ["Domain", "Urban water & climate adaptation · urban mobility · smart cities · Dutch Omgevingswet, CROW and City Deal frameworks"],
  ["Programming & data", "Python · R · LaTeX"],
  ["Certified", "UAS A1/A3 drone operator · TEFL/TESOL Level 5 · CEFR C1 English"],
];

const PUBLIC_RECORD = [
  {
    text: "Ton Duc Thang University, Faculty of Civil Engineering — feature article on the internship, with certificate of completion",
    url: "https://civil.tdtu.edu.vn/en/news-events/2026/strengthening-global-academic-ties-dutch-student-bjorn-hettemas-internship-urban",
  },
  {
    text: "Universitas Diponegoro, Postgraduate School — named lead researcher on the Saxion–UNDIP canal research collaboration",
    url: "https://pasca.undip.ac.id/en/collaboration-between-sps-undip-and-saxion-university-on-researching-the-history-of-canal-transformation-of-the-javanese-shore/",
  },
  {
    text: "FIP-AM@UT, University of Twente — “Vollebak”, the smart waste project behind the internship",
    url: "https://fip.utwente.nl/nl/project/vollebak/",
  },
  {
    text: "s.v. Watt — Wall of Fame, listed as chairman of the 18th, 19th and 20th boards",
    url: "https://svwatt.com/en/association/board/wall-of-fame/",
  },
];

const HOSOCO_TECH = {
  role: "Spatial planning & smart mobility intern",
  org: "Hosoco (Holland Software Co.) · Ton Duc Thang University (VN) · Saxion (NL)",
  dates: "02/2025 – 08/2025",
  bullets: [
    "Built a Python program generating sensor-driven, dynamic waste-collection routes, integrating live bin fill-level data with the Google Routes API",
    "Modelled the road network as a weighted graph of intersections, bins and depot, using cached Dijkstra shortest paths, a nearest-insertion construction heuristic, and event-triggered re-routing when a bin crosses its fill threshold",
    "Designed and compared four bin-prioritisation models: rule-based thresholds, weighted arithmetic mean, exponential weighting, and a prize-collecting formulation driven by overflow probability",
    "Evaluated the Google Routes, NextBillion.ai and HERE routing APIs against fleet requirements, input structure, pricing and rate limits",
    "Presented the advisory framework to academic and industry stakeholders; the work was published by TDTU’s Faculty of Civil Engineering",
  ],
};

const HOSOCO_RESEARCH = {
  role: "Spatial planning & smart mobility intern",
  org: "Hosoco (Holland Software Co.) · Ton Duc Thang University (VN) · Saxion (NL)",
  dates: "02/2025 – 08/2025",
  bullets: [
    "Researched smart-city applications for municipal waste management across social, spatial, regulatory and market dimensions, covering Dutch municipal regulation, the EU Waste Framework Directive and SDG 11/12",
    "Built a Python program integrating live sensor data with the Google Routes API to generate dynamic collection routes, and specified the prioritisation logic behind it",
    "Presented the advisory framework to academic and industry stakeholders; the work was published by TDTU’s Faculty of Civil Engineering",
  ],
};

const ROLES_OTHER = [
  {
    role: "Chairman, academic committee Urban Studies",
    org: "Saxion University of Applied Sciences (NL)",
    dates: "09/2023 – 10/2025",
    bullets: [
      "Represented student interests in curriculum development and quality of education; advised on the Education and Examination Regulations",
      "Led biweekly meetings with academic supervisors and represented the academy in the six-yearly national accreditation evaluation",
    ],
  },
  {
    role: "Chairman, 18th, 19th and 20th boards",
    org: "Study association s.v. Watt (NL)",
    dates: "02/2021 – 10/2022",
    bullets: [
      "First representative towards 500+ members, 15 academies and 5 external sponsors across three consecutive terms",
      "Coordinated educational initiatives in partnership with university faculty",
    ],
  },
  {
    role: "Operations employee",
    org: "Albert Heijn, Deventer (NL)",
    dates: "10/2017 – 09/2026",
    bullets: [
      "Nine years of continuous part-time employment alongside full-time study",
    ],
  },
];

const P_THESIS = {
  role: "Graduation research — “Adaptive transformation of historical canals and quay walls in Java, Indonesia”",
  org: "Professorship of Sustainable Areas & Soil Transitions, Saxion (NL) · Universitas Diponegoro (ID) · graded 4.0",
  dates: "02/2026 – 08/2026",
  bullets: [
    "Sole author of a comparative study across Semarang, Jakarta and Surabaya assessing technical integrity, functional use and spatial opportunity of colonial-era canal systems",
    "Designed a segment-level 1–5 scoring instrument applied identically across all three cities, triangulated with historical document analysis, QGIS mapping and photographic field validation",
    "Mapped urban water-governance structures per city and produced adaptive-transformation recommendations framed by the Dutch City Deal “Timeless Canals”",
    "Named lead researcher by Universitas Diponegoro, coordinating one lecturer and three master’s students",
  ],
};

const P_DEVENTER = {
  role: "“Deventer by foot” — pedestrian network study",
  org: "Client: Goudappel (NL) · Saxion StadsLAB",
  dates: "09/2025 – 02/2026",
  bullets: [
    "Built the digital analysis layer: Overpass Turbo queries against OpenStreetMap, node filtering against defined inclusion criteria, and four QGIS network maps to the CROW national pedestrian-network standard",
    "Worked with Goudappel’s Nederlands Verplaatsingspanel mobility data alongside own pedestrian counts and survey instruments",
    "Contributed to the methodological framework (STOMP) and the physical barrier analysis across the IJssel, the canals and the rail network",
  ],
};

const P_ERASMUS = {
  role: "Erasmus+ “Regenerative Urbanism” — blended intensive programme",
  org: "International multidisciplinary team (NL)",
  dates: "09/2025 – 10/2025",
  bullets: [
    "Developed a regenerative urban vision transforming a station-area parking site into a climate-resilient, multifunctional public space",
    "Presented recommendations to Witteveen+Bos",
  ],
};

const P_TWENTE = {
  role: "Minor — “Engineering the future (under)world”",
  org: "Saxion (NL) · municipalities across the Twente region",
  dates: "09/2023 – 10/2025",
  bullets: [
    "Translated soil and subsurface policy objectives into an operational work plan with Twente municipalities, within the funded national programme “Bodemkracht van Twente”",
    "Produced water- and soil-driven area-development recommendations using demographic analysis in R and GIS-based suitability mapping",
  ],
};

const P_ALBANIA = {
  role: "“Inclusive accessibility in the public domain”",
  org: "Universiteti Politeknik i Tiranës (AL) · Saxion (NL)",
  dates: "02/2024 – 06/2024",
  bullets: [
    "Comparative spatial analysis of accessibility and inclusiveness across public spaces, transport hubs and parks in the Netherlands and Albania",
    "Developed spatial and policy recommendations for accessibility, walkability and social inclusion",
  ],
};

const EDUCATION = [
  ["B.Sc. Spatial Development", "Saxion University of Applied Sciences, Deventer (NL)", "2022 – 2026"],
  ["Environmental Science — collaborative partnership", "Universitas Diponegoro, Semarang (ID)", "2026"],
  ["City, Urban & Regional Planning — collaborative partnership", "Ton Duc Thang University (VN)", "2025"],
  ["Electrical & Information Engineering — foundation, not completed", "Saxion (NL)", "2019 – 2022"],
  ["Technical Informatics / Computer Engineering — foundation, not completed", "Saxion (NL)", "2018 – 2019"],
];

const LANGUAGES = "Dutch (native)  ·  English (native/bilingual, CEFR C1)  ·  German (professional working)  ·  Vietnamese (elementary)  ·  Russian (elementary)  ·  French (basic)";

/* ---------------- builders ---------------- */

const RIGHT_TAB = [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }];

function sectionHead(text) {
  return new Paragraph({
    spacing: { before: 260, after: 90 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 3 } },
    children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 17, font: FONT, color: ACCENT, characterSpacing: 24 })],
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after == null ? 60 : opts.after, line: 250 },
    children: [new TextRun({ text, size: 19, font: FONT, color: INK })],
  });
}

function bullet(text) {
  return new Paragraph({
    bullet: { level: 0 },
    spacing: { after: 40, line: 240 },
    indent: { left: convertInchesToTwip(0.25), hanging: convertInchesToTwip(0.15) },
    children: [new TextRun({ text, size: 18, font: FONT, color: INK })],
  });
}

function entry(e) {
  const out = [
    new Paragraph({
      spacing: { before: 140, after: 0 },
      tabStops: RIGHT_TAB,
      children: [
        new TextRun({ text: e.role, bold: true, size: 19, font: FONT, color: INK }),
        new TextRun({ text: "\t" + e.dates, size: 17, font: FONT, color: MUTED }),
      ],
    }),
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: e.org, size: 17, font: FONT, color: MUTED, italics: true })],
    }),
  ];
  e.bullets.forEach((b) => out.push(bullet(b)));
  return out;
}

function skillLines(rows) {
  return rows.map(([label, val]) =>
    new Paragraph({
      spacing: { after: 46, line: 240 },
      indent: { left: convertInchesToTwip(1.15), hanging: convertInchesToTwip(1.15) },
      children: [
        new TextRun({ text: label + "   ", bold: true, size: 18, font: FONT, color: ACCENT }),
        new TextRun({ text: val, size: 18, font: FONT, color: INK }),
      ],
    })
  );
}

function recordLines() {
  return PUBLIC_RECORD.map((r) =>
    new Paragraph({
      spacing: { after: 56, line: 240 },
      indent: { left: convertInchesToTwip(0.25), hanging: convertInchesToTwip(0.15) },
      bullet: { level: 0 },
      children: [
        new TextRun({ text: r.text + "  ", size: 18, font: FONT, color: INK }),
        new ExternalHyperlink({
          link: r.url,
          children: [new TextRun({ text: "link", size: 17, font: FONT, color: ACCENT, underline: {} })],
        }),
      ],
    })
  );
}

function educationLines() {
  return EDUCATION.map(([deg, inst, yrs]) =>
    new Paragraph({
      spacing: { after: 58 },
      tabStops: RIGHT_TAB,
      children: [
        new TextRun({ text: deg, bold: true, size: 18, font: FONT, color: INK }),
        new TextRun({ text: "  —  " + inst, size: 17, font: FONT, color: MUTED }),
        new TextRun({ text: "\t" + yrs, size: 17, font: FONT, color: MUTED }),
      ],
    })
  );
}

function header(tagline, web) {
  return [
    new Paragraph({
      spacing: { after: 30 },
      children: [new TextRun({ text: "Bjorn Hettema", bold: true, size: 40, font: FONT, color: INK, characterSpacing: -10 })],
    }),
    new Paragraph({
      spacing: { after: 70 },
      children: [new TextRun({ text: tagline, size: 19, font: FONT, color: ACCENT })],
    }),
    new Paragraph({
      spacing: { after: 20 },
      children: [new TextRun({ text: web ? CONTACT_WEB : CONTACT_FULL, size: 17, font: FONT, color: MUTED })],
    }),
    new Paragraph({
      spacing: { after: 40 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 6 } },
      children: [
        new ExternalHyperlink({
          link: SITE_URL,
          children: [new TextRun({ text: SITE, size: 17, font: FONT, color: ACCENT, underline: {} })],
        }),
        new TextRun({ text: "   \u00b7   ", size: 17, font: FONT, color: MUTED }),
        new ExternalHyperlink({
          link: "https://www." + LINKEDIN,
          children: [new TextRun({ text: LINKEDIN, size: 17, font: FONT, color: ACCENT, underline: {} })],
        }),
      ],
    }),
  ];
}

function buildDoc(variant, web) {
  const tech = variant === "tech";
  const kids = [];

  kids.push(...header(
    tech
      ? "B.Sc. Spatial Development  ·  Urban data, geospatial analysis & routing systems"
      : "B.Sc. Spatial Development  ·  Urban water, climate adaptation & smart cities",
    web
  ));

  kids.push(sectionHead("Profile"));
  kids.push(body(tech ? PROFILE_TECH : PROFILE_RESEARCH, { after: 40 }));

  kids.push(sectionHead(tech ? "Technical skills" : "Core competencies"));
  kids.push(...skillLines(tech ? SKILLS_TECH : SKILLS_RESEARCH));

  if (tech) {
    kids.push(sectionHead("Work experience"));
    [HOSOCO_TECH, ...ROLES_OTHER].forEach((e) => kids.push(...entry(e)));

    kids.push(sectionHead("Selected projects"));
    [P_DEVENTER, P_THESIS, P_TWENTE, P_ERASMUS, P_ALBANIA].forEach((e) => kids.push(...entry(e)));
  } else {
    kids.push(sectionHead("Research & project experience"));
    [P_THESIS, P_DEVENTER, P_TWENTE, P_ERASMUS, P_ALBANIA].forEach((e) => kids.push(...entry(e)));

    kids.push(sectionHead("Work experience"));
    [HOSOCO_RESEARCH, ...ROLES_OTHER].forEach((e) => kids.push(...entry(e)));
  }

  kids.push(sectionHead("Published record"));
  kids.push(...recordLines());

  kids.push(sectionHead("Education"));
  kids.push(...educationLines());

  kids.push(sectionHead("Languages"));
  kids.push(body(LANGUAGES, { after: 0 }));

  return new Document({
    creator: "Bjorn Hettema",
    lastModifiedBy: "Bjorn Hettema",
    title: "Bjorn Hettema — CV",
    description: "Curriculum vitae",
    styles: { default: { document: { run: { font: FONT, size: 19, color: INK } } } },
    sections: [{
      properties: {
        page: { margin: { top: 850, bottom: 700, left: 950, right: 950 } },
      },
      children: kids,
    }],
  });
}

const targets = [
  ["tech",     false, "/mnt/user-data/outputs/CV-BjornHettema-urban-data-geospatial-full.docx"],
  ["research", false, "/mnt/user-data/outputs/CV-BjornHettema-urban-water-research-full.docx"],
  ["tech",     true,  "/mnt/user-data/outputs/CV-BjornHettema-urban-data-geospatial.docx"],
  ["research", true,  "/mnt/user-data/outputs/CV-BjornHettema-urban-water-research.docx"],
];

(async () => {
  for (const [variant, web, path] of targets) {
    const buf = await Packer.toBuffer(buildDoc(variant, web));
    fs.writeFileSync(path, buf);
    console.log("wrote", path, buf.length, "bytes");
  }
})();
