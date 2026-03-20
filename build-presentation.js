const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Vansh Rana";
pres.title = "Vansh Rana – Portfolio Presentation";

// ── Color palette (no # prefix) ──
const C = {
  bg:        "07111F",
  bgCard:    "0D1A2B",
  bgMid:     "0D2137",
  accent:    "7DF9C6",
  blue:      "77A8FF",
  purple:    "B482FF",
  white:     "FFFFFF",
  muted:     "94A3B8",
  dark:      "1E3A5F",
};

const makeShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.25 });

// ════════════════════════════════════════════
// SLIDE 1 — Title / Hero
// ════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  // Left accent bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.06, h: 5.625,
    fill: { color: C.accent }, line: { color: C.accent }
  });

  // Large name
  s.addText("Vansh Rana", {
    x: 0.5, y: 0.7, w: 6.5, h: 1.3,
    fontSize: 52, bold: true, color: C.white,
    fontFace: "Calibri", margin: 0
  });

  // Title tagline
  s.addText("Aspiring Data Analyst & Software Developer", {
    x: 0.5, y: 1.95, w: 6.5, h: 0.55,
    fontSize: 18, color: C.accent, bold: false,
    fontFace: "Calibri", margin: 0
  });

  // Description
  s.addText(
    "Computer Science undergraduate with hands-on projects in data analytics, machine learning, and software development. Available for internships and entry-level roles.",
    {
      x: 0.5, y: 2.6, w: 6.0, h: 1.0,
      fontSize: 13, color: C.muted, fontFace: "Calibri",
      lineSpacingMultiple: 1.3, margin: 0
    }
  );

  // Skill badges
  const badges = ["Python", "SQL", "Power BI", "JavaScript", "Machine Learning"];
  badges.forEach((b, i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.5 + i * 1.26, y: 3.75, w: 1.18, h: 0.38,
      fill: { color: C.dark }, line: { color: C.accent, width: 1 },
      rectRadius: 0.05
    });
    s.addText(b, {
      x: 0.5 + i * 1.26, y: 3.75, w: 1.18, h: 0.38,
      fontSize: 9.5, color: C.accent, align: "center", valign: "middle",
      fontFace: "Calibri", margin: 0
    });
  });

  // Contact row
  s.addText("Vanshranaglobal@gmail.com   |   +91 91057 70369   |   linkedin.com/in/vanshrana369   |   github.com/vanshrana369", {
    x: 0.5, y: 4.35, w: 9, h: 0.4,
    fontSize: 10, color: C.muted, fontFace: "Calibri", margin: 0
  });

  // Right panel — stats
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.4, y: 0.5, w: 2.3, h: 4.7,
    fill: { color: C.bgCard }, line: { color: C.dark, width: 1 },
    shadow: makeShadow()
  });

  const stats = [
    { val: "3", lbl: "Featured\nProjects" },
    { val: "4", lbl: "Certifications" },
    { val: "3rd", lbl: "Hackathon\nFinish" },
  ];
  stats.forEach((st, i) => {
    s.addText(st.val, {
      x: 7.45, y: 0.85 + i * 1.45, w: 2.2, h: 0.7,
      fontSize: 40, bold: true, color: C.accent,
      align: "center", fontFace: "Calibri", margin: 0
    });
    s.addText(st.lbl, {
      x: 7.45, y: 1.5 + i * 1.45, w: 2.2, h: 0.55,
      fontSize: 10, color: C.muted, align: "center",
      fontFace: "Calibri", margin: 0
    });
  });

  // Slide label
  s.addText("PORTFOLIO PRESENTATION  ·  2025", {
    x: 0.5, y: 5.25, w: 9, h: 0.3,
    fontSize: 8, color: "2A4A6B", charSpacing: 2,
    fontFace: "Calibri", margin: 0
  });
}

// ════════════════════════════════════════════
// SLIDE 2 — About Me
// ════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  // Section kicker
  s.addText("ABOUT ME", {
    x: 0.5, y: 0.35, w: 9, h: 0.3,
    fontSize: 9, color: C.accent, bold: true, charSpacing: 3,
    fontFace: "Calibri", margin: 0
  });
  s.addText("About Me", {
    x: 0.5, y: 0.65, w: 9, h: 0.55,
    fontSize: 28, bold: true, color: C.white, fontFace: "Calibri", margin: 0
  });

  // Left card — bio
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.35, w: 5.5, h: 3.6,
    fill: { color: C.bgCard }, line: { color: C.dark, width: 1 },
    shadow: makeShadow()
  });
  s.addText([
    { text: "B.Tech Computer Science", options: { bold: true, color: C.accent } },
    { text: " student at Lovely Professional University.", options: { color: C.white } },
    { text: "\n\nHands-on work includes ", options: { color: C.muted } },
    { text: "Power BI dashboards", options: { color: C.blue } },
    { text: ", data cleaning, exploratory analysis, ", options: { color: C.muted } },
    { text: "machine learning workflows", options: { color: C.blue } },
    { text: ", and ongoing DSA practice.", options: { color: C.muted } },
    { text: "\n\nA fast learner who values clear communication, takes ownership of work, and consistently delivers on commitments.", options: { color: C.muted } },
  ], {
    x: 0.65, y: 1.55, w: 5.0, h: 3.2,
    fontSize: 13, fontFace: "Calibri", lineSpacingMultiple: 1.4
  });

  // Right column — detail cards
  const details = [
    { lbl: "University", val: "Lovely Professional University" },
    { lbl: "Degree", val: "B.Tech CSE · 2023 – Present" },
    { lbl: "CGPA", val: "6.83 / 10" },
    { lbl: "Location", val: "Phagwara, Punjab, India" },
  ];
  details.forEach((d, i) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 6.2, y: 1.35 + i * 0.88, w: 3.5, h: 0.75,
      fill: { color: C.bgCard }, line: { color: C.dark, width: 1 },
      shadow: makeShadow()
    });
    s.addText(d.lbl.toUpperCase(), {
      x: 6.4, y: 1.4 + i * 0.88, w: 3.1, h: 0.22,
      fontSize: 8, color: C.accent, bold: true, charSpacing: 1,
      fontFace: "Calibri", margin: 0
    });
    s.addText(d.val, {
      x: 6.4, y: 1.62 + i * 0.88, w: 3.1, h: 0.35,
      fontSize: 12, color: C.white, fontFace: "Calibri", margin: 0
    });
  });

  s.addText("2  /  10", { x: 9.3, y: 5.3, w: 0.6, h: 0.2, fontSize: 8, color: "2A4A6B", fontFace: "Calibri", margin: 0 });
}

// ════════════════════════════════════════════
// SLIDE 3 — Skills & Technologies
// ════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgMid };

  s.addText("SKILLS & TECHNOLOGIES", {
    x: 0.5, y: 0.35, w: 9, h: 0.3,
    fontSize: 9, color: C.accent, bold: true, charSpacing: 3,
    fontFace: "Calibri", margin: 0
  });
  s.addText("Skills & Technologies", {
    x: 0.5, y: 0.65, w: 9, h: 0.55,
    fontSize: 28, bold: true, color: C.white, fontFace: "Calibri", margin: 0
  });

  const cats = [
    {
      title: "Programming Languages",
      color: C.blue,
      items: ["Python", "HTML", "JavaScript", "C++", "Java", "SQL"]
    },
    {
      title: "Data Analysis & Machine Learning",
      color: C.accent,
      items: ["Pandas", "NumPy", "Matplotlib", "Seaborn", "Scikit-Learn", "Power BI"]
    },
    {
      title: "CS Foundations",
      color: C.purple,
      items: ["DSA", "OOP", "DBMS", "Operating Systems"]
    },
    {
      title: "Developer Tools",
      color: "FCD34D",
      items: ["Git", "GitHub", "VS Code", "Jupyter", "Excel"]
    },
    {
      title: "Soft Skills",
      color: "F472B6",
      items: ["Communication", "Team Leadership", "Problem Solving", "Time Management", "Adaptability", "Ownership"]
    },
  ];

  // 3 cols × 2 rows grid for 5 cards (last row has 2 centred)
  const cardW = 3.0, cardH = 1.65, gapX = 0.18, gapY = 0.18;
  const startX = 0.4;
  const startY = 1.35;
  cats.forEach((cat, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    // Centre last row (only 2 cards in row 1)
    const rowCount = row === 0 ? 3 : cats.length - 3;
    const rowStartX = row === 1 ? startX + (3 - rowCount) * (cardW + gapX) / 2 : startX;
    const x = rowStartX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);

    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: C.bgCard }, line: { color: C.dark, width: 1 },
      shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.07, h: cardH,
      fill: { color: cat.color }, line: { color: cat.color }
    });
    s.addText(cat.title, {
      x: x + 0.16, y: y + 0.1, w: cardW - 0.22, h: 0.32,
      fontSize: 10, bold: true, color: C.white, fontFace: "Calibri", margin: 0
    });
    s.addText(cat.items.join("  ·  "), {
      x: x + 0.16, y: y + 0.45, w: cardW - 0.22, h: 1.05,
      fontSize: 10, color: cat.color, fontFace: "Calibri",
      lineSpacingMultiple: 1.35, margin: 0
    });
  });

  s.addText("3  /  10", { x: 9.3, y: 5.3, w: 0.6, h: 0.2, fontSize: 8, color: "2A4A6B", fontFace: "Calibri", margin: 0 });
}

// ════════════════════════════════════════════
// SLIDE 4 — Projects Overview
// ════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  s.addText("PROJECTS", {
    x: 0.5, y: 0.35, w: 9, h: 0.3,
    fontSize: 9, color: C.accent, bold: true, charSpacing: 3,
    fontFace: "Calibri", margin: 0
  });
  s.addText("Projects", {
    x: 0.5, y: 0.65, w: 9, h: 0.55,
    fontSize: 28, bold: true, color: C.white, fontFace: "Calibri", margin: 0
  });

  const projects = [
    {
      date: "Oct 2025",
      kind: "Analytics",
      title: "Real Estate Analytics Dashboard",
      desc: "Power BI dashboard analyzing 15K+ listings — pricing trends, market KPIs, location performance.",
      color: C.blue,
    },
    {
      date: "May 2025",
      kind: "Machine Learning",
      title: "Parking Fee Prediction",
      desc: "End-to-end ML workflow: data cleaning, EDA, feature engineering, multi-model comparison with Scikit-Learn.",
      color: C.accent,
    },
    {
      date: "Current",
      kind: "Frontend",
      title: "Personal Portfolio Website",
      desc: "Responsive site in HTML/CSS/JS with animated UI, project showcase, and downloadable resume PDF.",
      color: C.purple,
    },
  ];

  projects.forEach((p, i) => {
    const x = 0.38 + i * 3.15;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.35, w: 3.0, h: 3.9,
      fill: { color: C.bgCard }, line: { color: C.dark, width: 1 },
      shadow: makeShadow()
    });
    // Top accent
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.35, w: 3.0, h: 0.07,
      fill: { color: p.color }, line: { color: p.color }
    });
    // Kind badge
    s.addShape(pres.shapes.RECTANGLE, {
      x: x + 0.15, y: 1.55, w: 1.2, h: 0.28,
      fill: { color: C.dark }, line: { color: p.color, width: 1 }
    });
    s.addText(p.kind, {
      x: x + 0.15, y: 1.55, w: 1.2, h: 0.28,
      fontSize: 8, color: p.color, align: "center", valign: "middle",
      fontFace: "Calibri", margin: 0
    });
    s.addText(p.date, {
      x: x + 0.15, y: 1.92, w: 2.7, h: 0.25,
      fontSize: 9, color: C.muted, fontFace: "Calibri", margin: 0
    });
    s.addText(p.title, {
      x: x + 0.15, y: 2.18, w: 2.7, h: 0.7,
      fontSize: 13, bold: true, color: C.white, fontFace: "Calibri",
      lineSpacingMultiple: 1.2, margin: 0
    });
    s.addText(p.desc, {
      x: x + 0.15, y: 2.95, w: 2.7, h: 1.8,
      fontSize: 11, color: C.muted, fontFace: "Calibri",
      lineSpacingMultiple: 1.35, margin: 0
    });
  });

  s.addText("4  /  10", { x: 9.3, y: 5.3, w: 0.6, h: 0.2, fontSize: 8, color: "2A4A6B", fontFace: "Calibri", margin: 0 });
}

// ════════════════════════════════════════════
// SLIDE 5 — Project Deep Dive: Real Estate Dashboard
// ════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgMid };

  s.addText("PROJECT SPOTLIGHT", {
    x: 0.5, y: 0.35, w: 9, h: 0.3,
    fontSize: 9, color: C.blue, bold: true, charSpacing: 3,
    fontFace: "Calibri", margin: 0
  });
  s.addText("Real Estate Analytics Dashboard", {
    x: 0.5, y: 0.65, w: 9, h: 0.6,
    fontSize: 26, bold: true, color: C.white, fontFace: "Calibri", margin: 0
  });

  // Left — what was built
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.4, w: 5.5, h: 3.8,
    fill: { color: C.bgCard }, line: { color: C.dark, width: 1 },
    shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.4, w: 0.07, h: 3.8,
    fill: { color: C.blue }, line: { color: C.blue }
  });
  s.addText("What was built", {
    x: 0.65, y: 1.5, w: 5.1, h: 0.35,
    fontSize: 12, bold: true, color: C.blue, fontFace: "Calibri", margin: 0
  });
  s.addText([
    { text: "Interactive Power BI dashboard to analyze a real estate dataset of 15,000+ listings.", options: { bullet: true, breakLine: true } },
    { text: "Used DAX and Power Query for data transformation and reporting logic.", options: { bullet: true, breakLine: true } },
    { text: "Built KPI panels tracking market value, average price, and total listings.", options: { bullet: true, breakLine: true } },
    { text: "Designed recruiter-friendly business insight visuals for clear storytelling.", options: { bullet: true } },
  ], {
    x: 0.65, y: 1.9, w: 5.0, h: 3.1,
    fontSize: 12, color: C.muted, fontFace: "Calibri",
    lineSpacingMultiple: 1.4
  });

  // Right — tools + outcome
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.2, y: 1.4, w: 3.5, h: 1.7,
    fill: { color: C.bgCard }, line: { color: C.dark, width: 1 },
    shadow: makeShadow()
  });
  s.addText("Tools Used", {
    x: 6.4, y: 1.5, w: 3.1, h: 0.3,
    fontSize: 10, bold: true, color: C.blue, fontFace: "Calibri", margin: 0
  });
  s.addText("Power BI  ·  DAX  ·  Power Query  ·  Excel", {
    x: 6.4, y: 1.85, w: 3.1, h: 0.9,
    fontSize: 12, color: C.white, fontFace: "Calibri",
    lineSpacingMultiple: 1.4, margin: 0
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.2, y: 3.3, w: 3.5, h: 1.9,
    fill: { color: C.bgCard }, line: { color: C.dark, width: 1 },
    shadow: makeShadow()
  });
  s.addText("Key Outcome", {
    x: 6.4, y: 3.4, w: 3.1, h: 0.3,
    fontSize: 10, bold: true, color: C.blue, fontFace: "Calibri", margin: 0
  });
  s.addText("A fully functional analytics dashboard that demonstrates end-to-end data analysis skills for business use cases.", {
    x: 6.4, y: 3.75, w: 3.1, h: 1.2,
    fontSize: 12, color: C.muted, fontFace: "Calibri",
    lineSpacingMultiple: 1.35, margin: 0
  });

  s.addText("5  /  10", { x: 9.3, y: 5.3, w: 0.6, h: 0.2, fontSize: 8, color: "2A4A6B", fontFace: "Calibri", margin: 0 });
}

// ════════════════════════════════════════════
// SLIDE 6 — Project Deep Dive: ML + Portfolio
// ════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  s.addText("PROJECT SPOTLIGHT", {
    x: 0.5, y: 0.35, w: 9, h: 0.3,
    fontSize: 9, color: C.accent, bold: true, charSpacing: 3,
    fontFace: "Calibri", margin: 0
  });
  s.addText("Machine Learning  +  Portfolio Website", {
    x: 0.5, y: 0.65, w: 9, h: 0.6,
    fontSize: 24, bold: true, color: C.white, fontFace: "Calibri", margin: 0
  });

  // ML card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.38, w: 4.6, h: 3.9,
    fill: { color: C.bgCard }, line: { color: C.dark, width: 1 },
    shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.38, w: 0.07, h: 3.9,
    fill: { color: C.accent }, line: { color: C.accent }
  });
  s.addText("Parking Fee Prediction  ·  May 2025", {
    x: 0.62, y: 1.48, w: 4.1, h: 0.35,
    fontSize: 12, bold: true, color: C.accent, fontFace: "Calibri", margin: 0
  });
  s.addText([
    { text: "End-to-end ML pipeline from raw data to trained model.", options: { bullet: true, breakLine: true } },
    { text: "Handled missing values and improved training-readiness.", options: { bullet: true, breakLine: true } },
    { text: "Compared multiple Scikit-Learn regression models.", options: { bullet: true, breakLine: true } },
    { text: "Used Matplotlib / Seaborn for visual analysis to support decisions.", options: { bullet: true } },
  ], {
    x: 0.62, y: 1.9, w: 4.1, h: 2.2,
    fontSize: 11.5, color: C.muted, fontFace: "Calibri", lineSpacingMultiple: 1.4
  });
  s.addText("Python  ·  Pandas  ·  Scikit-Learn  ·  Seaborn", {
    x: 0.62, y: 4.5, w: 4.1, h: 0.4,
    fontSize: 10, color: C.accent, fontFace: "Calibri", margin: 0
  });

  // Portfolio card
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.38, w: 4.6, h: 3.9,
    fill: { color: C.bgCard }, line: { color: C.dark, width: 1 },
    shadow: makeShadow()
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.2, y: 1.38, w: 0.07, h: 3.9,
    fill: { color: C.purple }, line: { color: C.purple }
  });
  s.addText("Portfolio Website  ·  Current Focus", {
    x: 5.42, y: 1.48, w: 4.1, h: 0.35,
    fontSize: 12, bold: true, color: C.purple, fontFace: "Calibri", margin: 0
  });
  s.addText([
    { text: "Responsive, animated portfolio built entirely from scratch.", options: { bullet: true, breakLine: true } },
    { text: "Canvas-based animated background (synthwave perspective grid).", options: { bullet: true, breakLine: true } },
    { text: "Structured for fast recruiter and hiring-manager review.", options: { bullet: true, breakLine: true } },
    { text: "Downloadable resume PDF and live contact form integration.", options: { bullet: true } },
  ], {
    x: 5.42, y: 1.9, w: 4.1, h: 2.2,
    fontSize: 11.5, color: C.muted, fontFace: "Calibri", lineSpacingMultiple: 1.4
  });
  s.addText("HTML  ·  CSS  ·  JavaScript  ·  Canvas API", {
    x: 5.42, y: 4.5, w: 4.1, h: 0.4,
    fontSize: 10, color: C.purple, fontFace: "Calibri", margin: 0
  });

  s.addText("6  /  10", { x: 9.3, y: 5.3, w: 0.6, h: 0.2, fontSize: 8, color: "2A4A6B", fontFace: "Calibri", margin: 0 });
}

// ════════════════════════════════════════════
// SLIDE 7 — Experience & Training
// ════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgMid };

  s.addText("SUMMER TRAINING", {
    x: 0.5, y: 0.35, w: 9, h: 0.3,
    fontSize: 9, color: C.accent, bold: true, charSpacing: 3,
    fontFace: "Calibri", margin: 0
  });
  s.addText("Summer Training", {
    x: 0.5, y: 0.65, w: 9, h: 0.55,
    fontSize: 28, bold: true, color: C.white, fontFace: "Calibri", margin: 0
  });

  const items = [
    {
      title: "Summer Training — DSA in C++",
      date: "Jun 2025",
      desc: "Worked on arrays, linked lists, stacks, queues, trees, graphs, heaps, searching, sorting, and dynamic programming.",
      tags: ["C++", "DSA", "Algorithms"],
      color: C.blue,
    },
  ];

  items.forEach((item, i) => {
    const y = 1.38 + i * 2.0;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 9.2, h: 1.75,
      fill: { color: C.bgCard }, line: { color: C.dark, width: 1 },
      shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 0.07, h: 1.75,
      fill: { color: item.color }, line: { color: item.color }
    });
    s.addText(item.title, {
      x: 0.65, y: y + 0.12, w: 7.0, h: 0.35,
      fontSize: 13, bold: true, color: C.white, fontFace: "Calibri", margin: 0
    });
    s.addText(item.date, {
      x: 8.0, y: y + 0.12, w: 1.4, h: 0.35,
      fontSize: 10, color: item.color, align: "right", fontFace: "Calibri", margin: 0
    });
    s.addText(item.desc, {
      x: 0.65, y: y + 0.52, w: 8.7, h: 0.75,
      fontSize: 12, color: C.muted, fontFace: "Calibri",
      lineSpacingMultiple: 1.35, margin: 0
    });
    s.addText(item.tags.join("   ·   "), {
      x: 0.65, y: y + 1.35, w: 8.7, h: 0.28,
      fontSize: 9.5, color: item.color, fontFace: "Calibri", margin: 0
    });
  });

  // Status callout
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 5.1, w: 9.2, h: 0.38,
    fill: { color: C.dark }, line: { color: C.accent, width: 1 }
  });
  s.addText("Currently seeking internships and entry-level roles in analytics and software development.", {
    x: 0.6, y: 5.1, w: 8.8, h: 0.38,
    fontSize: 11, color: C.accent, valign: "middle", fontFace: "Calibri", margin: 0
  });

  s.addText("7  /  10", { x: 9.3, y: 5.3, w: 0.6, h: 0.2, fontSize: 8, color: "2A4A6B", fontFace: "Calibri", margin: 0 });
}

// ════════════════════════════════════════════
// SLIDE 8 — Certifications
// ════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  s.addText("CERTIFICATIONS & COURSES", {
    x: 0.5, y: 0.35, w: 9, h: 0.3,
    fontSize: 9, color: C.accent, bold: true, charSpacing: 3,
    fontFace: "Calibri", margin: 0
  });
  s.addText("Formal learning, online coursework, and verified upskilling.", {
    x: 0.5, y: 0.65, w: 9, h: 0.55,
    fontSize: 22, bold: true, color: C.white, fontFace: "Calibri", margin: 0
  });

  const certs = [
    { title: "Introduction to Data Analytics", issuer: "IBM", date: "Nov 2025", color: C.blue },
    { title: "Privacy and Security in Social Media", issuer: "NPTEL", date: "May 2025", color: C.accent },
    { title: "Responsive Web Design", issuer: "freeCodeCamp", date: "Oct 2023", color: C.purple },
    { title: "Bits and Bytes of Computer Networking", issuer: "Google / Coursera", date: "Sep 2024", color: "FCD34D" },
  ];

  certs.forEach((cert, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.4 + col * 4.8;
    const y = 1.38 + row * 1.88;

    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 4.6, h: 1.65,
      fill: { color: C.bgCard }, line: { color: C.dark, width: 1 },
      shadow: makeShadow()
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.07, h: 1.65,
      fill: { color: cert.color }, line: { color: cert.color }
    });
    s.addText(cert.title, {
      x: x + 0.2, y: y + 0.12, w: 4.2, h: 0.6,
      fontSize: 13, bold: true, color: C.white, fontFace: "Calibri",
      lineSpacingMultiple: 1.2, margin: 0
    });
    s.addText(cert.issuer, {
      x: x + 0.2, y: y + 0.78, w: 3.0, h: 0.3,
      fontSize: 11, color: cert.color, fontFace: "Calibri", margin: 0
    });
    s.addText(cert.date, {
      x: x + 0.2, y: y + 1.2, w: 4.2, h: 0.28,
      fontSize: 9.5, color: C.muted, fontFace: "Calibri", margin: 0
    });
  });

  s.addText("8  /  10", { x: 9.3, y: 5.3, w: 0.6, h: 0.2, fontSize: 8, color: "2A4A6B", fontFace: "Calibri", margin: 0 });
}

// ════════════════════════════════════════════
// SLIDE 9 — Hackathons & Education
// ════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgMid };

  s.addText("HACKATHONS & EDUCATION", {
    x: 0.5, y: 0.35, w: 9, h: 0.3,
    fontSize: 9, color: C.accent, bold: true, charSpacing: 3,
    fontFace: "Calibri", margin: 0
  });
  s.addText("Competitive performance and academic background.", {
    x: 0.5, y: 0.65, w: 9, h: 0.55,
    fontSize: 22, bold: true, color: C.white, fontFace: "Calibri", margin: 0
  });

  // Hackathon highlight
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.4, y: 1.35, w: 4.5, h: 1.55,
    fill: { color: C.bgCard }, line: { color: "FCD34D", width: 1 },
    shadow: makeShadow()
  });
  s.addText("🏆", {
    x: 0.5, y: 1.4, w: 0.7, h: 0.7,
    fontSize: 28, align: "center", fontFace: "Calibri"
  });
  s.addText("Hackathon — 3rd Place", {
    x: 1.25, y: 1.42, w: 3.5, h: 0.38,
    fontSize: 14, bold: true, color: "FCD34D", fontFace: "Calibri", margin: 0
  });
  s.addText("Placed 3rd among 30 competing groups — Mar 2024", {
    x: 1.25, y: 1.82, w: 3.5, h: 0.35,
    fontSize: 11, color: C.muted, fontFace: "Calibri", margin: 0
  });
  s.addText("Solved problems under time pressure and team constraints.", {
    x: 0.65, y: 2.23, w: 4.1, h: 0.45,
    fontSize: 11, color: C.muted, fontFace: "Calibri",
    lineSpacingMultiple: 1.3, margin: 0
  });

  // DSA Practice
  s.addShape(pres.shapes.RECTANGLE, {
    x: 5.1, y: 1.35, w: 4.5, h: 1.55,
    fill: { color: C.bgCard }, line: { color: C.blue, width: 1 },
    shadow: makeShadow()
  });
  s.addText("DSA Problem Solving", {
    x: 5.3, y: 1.42, w: 4.1, h: 0.38,
    fontSize: 14, bold: true, color: C.blue, fontFace: "Calibri", margin: 0
  });
  s.addText("Ongoing  ·  Regular practice", {
    x: 5.3, y: 1.82, w: 4.1, h: 0.3,
    fontSize: 10, color: C.muted, fontFace: "Calibri", margin: 0
  });
  s.addText("Arrays · Trees · Graphs · Recursion · Sorting · Dynamic Programming", {
    x: 5.3, y: 2.2, w: 4.1, h: 0.5,
    fontSize: 11, color: C.blue, fontFace: "Calibri",
    lineSpacingMultiple: 1.3, margin: 0
  });

  // Education table
  const edu = [
    { year: "2023 – Present", degree: "B.Tech CSE", school: "Lovely Professional University", score: "CGPA 6.83" },
    { year: "2022 – 2023", degree: "Class XII", school: "DR KN Modi Global School", score: "65%" },
    { year: "2020 – 2021", degree: "Class X", school: "DR KN Modi Global School", score: "81%" },
  ];

  s.addText("Education", {
    x: 0.4, y: 3.15, w: 9, h: 0.3,
    fontSize: 11, bold: true, color: C.accent, fontFace: "Calibri", margin: 0
  });

  edu.forEach((e, i) => {
    const y = 3.55 + i * 0.58;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 9.2, h: 0.5,
      fill: { color: i % 2 === 0 ? C.bgCard : C.bg }, line: { color: C.dark, width: 1 }
    });
    s.addText(e.year, { x: 0.55, y: y + 0.07, w: 1.6, h: 0.34, fontSize: 10, color: C.muted, fontFace: "Calibri", margin: 0 });
    s.addText(e.degree, { x: 2.25, y: y + 0.07, w: 2.5, h: 0.34, fontSize: 11, bold: true, color: C.white, fontFace: "Calibri", margin: 0 });
    s.addText(e.school, { x: 4.85, y: y + 0.07, w: 3.3, h: 0.34, fontSize: 10, color: C.muted, fontFace: "Calibri", margin: 0 });
    s.addText(e.score, { x: 8.25, y: y + 0.07, w: 1.2, h: 0.34, fontSize: 11, bold: true, color: C.accent, align: "right", fontFace: "Calibri", margin: 0 });
  });

  s.addText("9  /  10", { x: 9.3, y: 5.3, w: 0.6, h: 0.2, fontSize: 8, color: "2A4A6B", fontFace: "Calibri", margin: 0 });
}

// ════════════════════════════════════════════
// SLIDE 10 — Contact / Closing
// ════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };

  // Left accent bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.06, h: 5.625,
    fill: { color: C.accent }, line: { color: C.accent }
  });

  s.addText("LET'S CONNECT", {
    x: 0.5, y: 0.5, w: 9, h: 0.35,
    fontSize: 9, color: C.accent, bold: true, charSpacing: 3,
    fontFace: "Calibri", margin: 0
  });
  s.addText("Open to internships, fresher roles, and\nopportunities in analytics & development.", {
    x: 0.5, y: 0.88, w: 6.5, h: 1.2,
    fontSize: 26, bold: true, color: C.white, fontFace: "Calibri",
    lineSpacingMultiple: 1.2, margin: 0
  });

  const contacts = [
    { icon: "✉", label: "Vanshranaglobal@gmail.com" },
    { icon: "☎", label: "+91 91057 70369" },
    { icon: "in", label: "linkedin.com/in/vanshrana369" },
    { icon: "GH", label: "github.com/vanshrana369" },
  ];

  contacts.forEach((c, i) => {
    const y = 2.3 + i * 0.65;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 6.0, h: 0.52,
      fill: { color: C.bgCard }, line: { color: C.dark, width: 1 }
    });
    s.addText(c.icon, {
      x: 0.5, y: y + 0.06, w: 0.55, h: 0.38,
      fontSize: 13, color: C.accent, align: "center",
      fontFace: "Calibri", margin: 0
    });
    s.addText(c.label, {
      x: 1.1, y: y + 0.06, w: 5.1, h: 0.38,
      fontSize: 13, color: C.white, fontFace: "Calibri",
      valign: "middle", margin: 0
    });
  });

  // Right — closing statement
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.8, y: 1.2, w: 2.9, h: 4.0,
    fill: { color: C.dark }, line: { color: C.accent, width: 1 },
    shadow: makeShadow()
  });
  s.addText("Thank you", {
    x: 6.9, y: 1.45, w: 2.7, h: 0.6,
    fontSize: 22, bold: true, color: C.white, align: "center",
    fontFace: "Calibri", margin: 0
  });
  s.addText("Vansh Rana\nAspiring Data Analyst\n& Software Developer\n\nB.Tech CSE\nLovely Professional\nUniversity", {
    x: 6.9, y: 2.1, w: 2.7, h: 2.6,
    fontSize: 11, color: C.muted, align: "center",
    fontFace: "Calibri", lineSpacingMultiple: 1.5, margin: 0
  });
  s.addText("VR369", {
    x: 6.9, y: 4.7, w: 2.7, h: 0.35,
    fontSize: 16, bold: true, color: C.accent, align: "center",
    fontFace: "Calibri", margin: 0
  });

  s.addText("10  /  10", { x: 9.3, y: 5.3, w: 0.6, h: 0.2, fontSize: 8, color: "2A4A6B", fontFace: "Calibri", margin: 0 });
}

// ── Write file ──
pres.writeFile({ fileName: "D:/claude/portfolio/VanshRana-Portfolio-Presentation.pptx" })
  .then(() => console.log("✅  Saved: VanshRana-Portfolio-Presentation.pptx"))
  .catch(err => console.error("❌  Error:", err));
