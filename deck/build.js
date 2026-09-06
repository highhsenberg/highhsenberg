const pptxgen = require("pptxgenjs");

const CHERRY = "990011";
const NAVY = "2F3C7E";
const OFFWHITE = "FCF6F5";
const INK = "231F20";
const MUTED = "6B6168";
const BLUSH = "F6E4E6";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5
pres.author = "MBBS IV";
pres.title = "Adult Presentation of Congenital Heart Disease";

const HEAD = "Bookman Old Style";
const BODY = "Calibri";

function titleBar(slide, text, kicker) {
  if (kicker) {
    slide.addText(kicker.toUpperCase(), {
      x: 0.6, y: 0.34, w: 10, h: 0.28, isTextBox: true, margin: 0,
      fontFace: BODY, fontSize: 12, bold: true, charSpacing: 2, color: CHERRY,
    });
  }
  slide.addText(text, {
    x: 0.6, y: kicker ? 0.62 : 0.45, w: 12.1, h: 0.75, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 32, bold: true, color: NAVY,
  });
}

function dot(slide, x, y, label, fill) {
  slide.addShape(pres.ShapeType.ellipse, {
    x: x, y: y, w: 0.34, h: 0.34, fill: { color: fill || CHERRY },
  });
  slide.addText(label, {
    x: x, y: y, w: 0.34, h: 0.34, isTextBox: true, margin: 0, align: "center",
    valign: "middle", fontFace: BODY, fontSize: 13, bold: true, color: "FFFFFF",
  });
}

function card(slide, o) {
  slide.addShape(pres.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.09,
    fill: { color: o.fill || OFFWHITE },
    shadow: { type: "outer", blur: 8, offset: 1, angle: 90, color: "9A8A8A", opacity: 0.2 },
  });
}

/* ---------------------------------------------------------- 1. TITLE */
let s = pres.addSlide();
s.background = { color: NAVY };
s.addShape(pres.ShapeType.ellipse, { x: 9.3, y: -1.5, w: 6.2, h: 6.2, fill: { color: CHERRY }, transparency: 55 });
s.addShape(pres.ShapeType.ellipse, { x: 11.0, y: 3.6, w: 4.2, h: 4.2, fill: { color: "FFFFFF" }, transparency: 88 });

s.addText("SEMINAR PRESENTATION  |  DISEASES OF THE CARDIOVASCULAR SYSTEM", {
  x: 0.8, y: 1.25, w: 9, h: 0.3, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 11.5, bold: true, charSpacing: 2, color: "E7B9BE",
});
s.addText("Adult Presentation of\nCongenital Heart Disease", {
  x: 0.8, y: 1.75, w: 9.2, h: 1.9, isTextBox: true, margin: 0, lineSpacing: 44,
  fontFace: HEAD, fontSize: 38, bold: true, color: "FFFFFF",
});
s.addText("Causes of adult presentation, and the distinguishing features between cyanotic and acyanotic congenital heart disease", {
  x: 0.8, y: 3.85, w: 8.2, h: 0.9, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 16, color: "D9D3E0",
});
s.addShape(pres.ShapeType.roundRect, { x: 0.8, y: 5.15, w: 4.5, h: 0.72, rectRadius: 0.1, fill: { color: CHERRY } });
s.addText("Congenital heart disease affects about 1% of live births", {
  x: 0.8, y: 5.15, w: 4.5, h: 0.72, isTextBox: true, margin: 0, align: "center", valign: "middle",
  fontFace: BODY, fontSize: 12.5, bold: true, color: "FFFFFF",
});
s.addText("MBBS Fourth Year  •  General Medicine", {
  x: 0.8, y: 6.35, w: 6, h: 0.3, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 12, color: "B9B2C6",
});
s.addNotes("Two things to cover today: why congenital heart disease turns up for the first time in an adult, and how to tell cyanotic from acyanotic disease at the bedside.");

/* ------------------------------------------------- 2. WHY ADULTS AT ALL */
s = pres.addSlide();
s.background = { color: "FFFFFF" };
titleBar(s, "Why do we see congenital heart disease in adults?", "Setting the scene");

card(s, { x: 0.6, y: 1.7, w: 6.1, h: 2.05, fill: BLUSH });
s.addText("1%", { x: 0.95, y: 1.9, w: 1.6, h: 0.95, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 46, bold: true, color: CHERRY });
s.addText("of live births are affected by congenital heart disease.", {
  x: 2.55, y: 2.0, w: 3.85, h: 1.2, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 15, color: INK,
});
s.addText("Males are affected more commonly, except ASD and PDA, which are more common in females.", {
  x: 0.95, y: 3.05, w: 5.5, h: 0.55, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, italic: true, color: MUTED,
});

card(s, { x: 6.95, y: 1.7, w: 5.75, h: 2.05, fill: OFFWHITE });
s.addText("The main reason", { x: 7.3, y: 1.95, w: 5, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, bold: true, charSpacing: 1.5, color: CHERRY });
s.addText("Because of improved medical and surgical management, more children with congenital heart disease are surviving into adolescence and adulthood.", {
  x: 7.3, y: 2.3, w: 5.05, h: 1.3, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 15, color: INK,
});

card(s, { x: 0.6, y: 4.05, w: 12.1, h: 2.55, fill: "FFFFFF" });
s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 4.05, w: 12.1, h: 2.55, rectRadius: 0.09, fill: { color: "FFFFFF" }, line: { color: "E4D8DA", width: 1 } });
s.addText("Recognise it early", { x: 0.95, y: 4.28, w: 6, h: 0.35, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 19, bold: true, color: NAVY });
s.addText([
  { text: "Congenital heart disease should be recognized as early as possible, since early treatment is associated with better outcome.", options: { bullet: true, breakLine: true } },
  { text: "When it is missed in childhood, the disease declares itself later through its complications, and that is the adult presentation.", options: { bullet: true } },
], { x: 0.95, y: 4.75, w: 11.4, h: 1.6, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 15, color: INK, paraSpaceAfter: 8 });
s.addNotes("Frame it simply: survival has improved, so the population of adults with congenital heart disease keeps growing, and some of them were never diagnosed as children.");

/* --------------------------------------------- 3. CAUSES OF ADULT PRESENTATION */
s = pres.addSlide();
s.background = { color: "FFFFFF" };
titleBar(s, "Causes of adult presentation", "Enumeration");
s.addText("These complications may be the first time that the presence of congenital heart disease is noted.", {
  x: 0.6, y: 1.45, w: 11.5, h: 0.35, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14, italic: true, color: MUTED,
});

const causes = [
  ["Infective endocarditis", "Occurs at the sites of shunts and damaged valves."],
  ["Atrial and ventricular arrhythmias", "Present as palpitations in a previously undiagnosed adult."],
  ["Right heart failure due to pulmonary hypertension", "Follows a long standing left to right shunt."],
  ["End stage heart failure", "The late outcome of an uncorrected lesion."],
  ["Sudden cardiac death", "May be the very first event that reveals the diagnosis."],
];
let y = 1.95;
causes.forEach((c, i) => {
  card(s, { x: 0.6, y: y, w: 12.1, h: 0.86, fill: i % 2 === 0 ? OFFWHITE : "FFFFFF" });
  dot(s, 0.9, y + 0.26, String(i + 1), i === 4 ? CHERRY : NAVY);
  s.addText(c[0], { x: 1.45, y: y + 0.12, w: 5.4, h: 0.34, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 15.5, bold: true, color: NAVY });
  s.addText(c[1], { x: 1.45, y: y + 0.46, w: 10.9, h: 0.32, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13, color: MUTED });
  y += 0.99;
});
s.addNotes("Say this list in order. Endocarditis, arrhythmias, right heart failure from pulmonary hypertension, end stage heart failure, sudden cardiac death.");

/* ------------------------------------- 4. OTHER ROUTES OF ADULT PRESENTATION */
s = pres.addSlide();
s.background = { color: "FFFFFF" };
titleBar(s, "Other ways it surfaces in adult life", "Enumeration continued");

const others = [
  ["Eisenmenger's syndrome", "Pressure in the pulmonary arterial system exceeds that on the left side, and the shunt reverses from right to left. The patient becomes cyanosed."],
  ["Paradoxical embolism", "A thrombus passes from the systemic veins to the systemic arterial circulation through a communication such as a VSD, raising the risk of cerebrovascular accidents and abscesses."],
  ["Polycythemia and thromboembolism", "Secondary to chronic hypoxemia, leading to hyperviscosity and an increased risk of thromboembolism and strokes."],
  ["Pregnancy", "A woman with congenital heart disease needs close follow up. Pregnancy is usually safe except when pulmonary hypertension is present, where the prognosis for both mother and fetus is poor."],
];
y = 1.65;
others.forEach((o, i) => {
  card(s, { x: 0.6, y: y, w: 12.1, h: 1.2, fill: i % 2 === 0 ? BLUSH : OFFWHITE });
  s.addText(o[0], { x: 1.0, y: y + 0.17, w: 3.6, h: 0.85, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 16, bold: true, color: CHERRY });
  s.addText(o[1], { x: 4.75, y: y + 0.17, w: 7.6, h: 0.9, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 13.5, color: INK });
  y += 1.32;
});
s.addNotes("Eisenmenger, paradoxical embolism, polycythemia and pregnancy. Pregnancy matters because that is often the first proper cardiac assessment a young woman gets.");

/* --------------------------------------------- 5. EISENMENGER DIAGRAM */
s = pres.addSlide();
s.background = { color: NAVY };
s.addText("EISENMENGER'S SYNDROME", { x: 0.6, y: 0.45, w: 8, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, bold: true, charSpacing: 2, color: "E7B9BE" });
s.addText("How a left to right shunt ends up cyanotic", { x: 0.6, y: 0.78, w: 11, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 30, bold: true, color: "FFFFFF" });

const steps = [
  ["Left to right shunt", "Blood from the left side of the heart, under high pressure, enters the right side."],
  ["Into the pulmonary artery", "That extra volume is carried on into the pulmonary artery."],
  ["Pulmonary hypertension", "Pressure in the pulmonary arterial system rises."],
  ["Shunt reversal", "Pulmonary pressure exceeds left sided pressure, so flow reverses from right to left."],
];
let x = 0.6;
steps.forEach((st, i) => {
  s.addShape(pres.ShapeType.roundRect, { x: x, y: 1.95, w: 2.75, h: 2.6, rectRadius: 0.1, fill: { color: "FFFFFF" } });
  dot(s, x + 0.25, 2.2, String(i + 1), i === 3 ? CHERRY : NAVY);
  s.addText(st[0], { x: x + 0.25, y: 2.68, w: 2.3, h: 0.6, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 15, bold: true, color: NAVY });
  s.addText(st[1], { x: x + 0.25, y: 3.3, w: 2.3, h: 1.1, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, color: MUTED });
  if (i < 3) {
    s.addShape(pres.ShapeType.rightArrow, { x: x + 2.85, y: 3.05, w: 0.42, h: 0.36, fill: { color: "E7B9BE" } });
  }
  x += 3.27;
});
s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 5.0, w: 12.1, h: 1.15, rectRadius: 0.1, fill: { color: CHERRY } });
s.addText("This reversal of blood flow is referred to as Eisenmenger's syndrome. The patient who was acyanotic now becomes cyanotic, which is why an adult with a childhood murmur may present for the first time with cyanosis.", {
  x: 1.0, y: 5.15, w: 11.3, h: 0.85, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14.5, color: "FFFFFF",
});
s.addNotes("This is the single most important mechanism in the whole topic, because it links the acyanotic group to the cyanotic group.");

/* --------------------------------------------- 6. ETIOLOGY */
s = pres.addSlide();
s.background = { color: "FFFFFF" };
titleBar(s, "Etiology: the recognized risk factors", "Background");
s.addText("Congenital heart diseases are due to abnormal development of a normal structure, or failure of a normal structure to develop fully. Such maldevelopments are due to multifactorial genetic and environmental causes.", {
  x: 0.6, y: 1.5, w: 12.1, h: 0.75, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14.5, color: INK,
});

const eti = [
  ["Maternal infections", "Rubella infection, causing persistent ductus arteriosus and pulmonary stenosis."],
  ["Drugs", "Alcohol abuse causing septal defects, phenytoin associated with pulmonary stenosis, and radiation."],
  ["Genetic abnormalities", "The familial form of atrial septal defect and congenital heart block."],
  ["Chromosomal abnormalities", "Septal defects and tetralogy of Fallot in Down's syndrome (trisomy 21), coarctation of the aorta in Turner's syndrome (45, XO)."],
];
x = 0.6;
eti.forEach((e, i) => {
  s.addShape(pres.ShapeType.roundRect, { x: x, y: 2.5, w: 2.95, h: 3.5, rectRadius: 0.1, fill: { color: i % 2 === 0 ? OFFWHITE : BLUSH } });
  s.addShape(pres.ShapeType.ellipse, { x: x + 0.35, y: 2.85, w: 0.5, h: 0.5, fill: { color: CHERRY } });
  s.addText(String(i + 1), { x: x + 0.35, y: 2.85, w: 0.5, h: 0.5, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: BODY, fontSize: 16, bold: true, color: "FFFFFF" });
  s.addText(e[0], { x: x + 0.35, y: 3.5, w: 2.3, h: 0.7, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 15.5, bold: true, color: NAVY });
  s.addText(e[1], { x: x + 0.35, y: 4.25, w: 2.3, h: 1.6, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, color: MUTED });
  x += 3.11;
});
s.addNotes("Four headings only: maternal infections, drugs, genetic abnormalities, chromosomal abnormalities.");

/* --------------------------------------------- 7. CLASSIFICATION */
s = pres.addSlide();
s.background = { color: "FFFFFF" };
titleBar(s, "Classification of congenital heart diseases", "Table 3.20");

s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 1.65, w: 5.9, h: 4.55, rectRadius: 0.1, fill: { color: NAVY } });
s.addText("CYANOTIC", { x: 1.0, y: 1.95, w: 5, h: 0.4, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 22, bold: true, color: "FFFFFF" });
s.addText("Right to left shunting, so the patient is blue", { x: 1.0, y: 2.38, w: 5, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, italic: true, color: "E7B9BE" });
s.addText([
  { text: "Tetralogy of Fallot", options: { bullet: true, breakLine: true } },
  { text: "Transposition of the great vessels", options: { bullet: true, breakLine: true } },
  { text: "Tricuspid atresia", options: { bullet: true, breakLine: true } },
  { text: "Truncus arteriosus", options: { bullet: true, breakLine: true } },
  { text: "Total anomalous pulmonary venous connection (TAPVC)", options: { bullet: true } },
], { x: 1.0, y: 2.85, w: 5.1, h: 2.3, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14.5, color: "FFFFFF", paraSpaceAfter: 8 });
s.addShape(pres.ShapeType.roundRect, { x: 1.0, y: 5.35, w: 4.6, h: 0.6, rectRadius: 0.1, fill: { color: CHERRY } });
s.addText("Mnemonic: the five Ts", { x: 1.0, y: 5.35, w: 4.6, h: 0.6, isTextBox: true, margin: 0, align: "center", valign: "middle", fontFace: BODY, fontSize: 13.5, bold: true, color: "FFFFFF" });

s.addShape(pres.ShapeType.roundRect, { x: 6.8, y: 1.65, w: 5.9, h: 4.55, rectRadius: 0.1, fill: { color: BLUSH } });
s.addText("ACYANOTIC", { x: 7.2, y: 1.95, w: 5, h: 0.4, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 22, bold: true, color: CHERRY });
s.addText("Left to right shunt or an obstructive lesion, no cyanosis", { x: 7.2, y: 2.38, w: 5.2, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, italic: true, color: MUTED });
s.addText([
  { text: "Atrial septal defect (ASD)", options: { bullet: true, breakLine: true } },
  { text: "Ventricular septal defect (VSD)", options: { bullet: true, breakLine: true } },
  { text: "Patent ductus arteriosus (PDA)", options: { bullet: true, breakLine: true } },
  { text: "Coarctation of the aorta", options: { bullet: true, breakLine: true } },
  { text: "Aortic stenosis", options: { bullet: true, breakLine: true } },
  { text: "Congenital pulmonary valve regurgitation", options: { bullet: true, breakLine: true } },
  { text: "Pulmonary stenosis", options: { bullet: true } },
], { x: 7.2, y: 2.85, w: 5.1, h: 3.2, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14, color: INK, paraSpaceAfter: 5 });
s.addNotes("Five Ts for the cyanotic group. The acyanotic group is the septal defects, the duct, and the obstructive lesions.");

/* --------------------------------------------- 8. DISTINGUISHING FEATURES */
s = pres.addSlide();
s.background = { color: "FFFFFF" };
titleBar(s, "Distinguishing features at the bedside", "Cyanotic vs acyanotic");

const rows = [
  ["Direction of flow", "Right to left shunting, or mixing of systemic and pulmonary blood flow", "Left to right shunt, or an obstruction to outflow"],
  ["Cyanosis", "Central cyanosis is present", "Absent, until the shunt reverses"],
  ["Clubbing", "Occurs due to prolonged cyanosis", "Not a feature"],
  ["Polycythemia", "Develops secondary to chronic hypoxemia, causing hyperviscosity", "Not expected"],
  ["Growth", "Growth retardation is common in children", "Growth usually preserved"],
  ["Pulmonary circulation", "Pulmonary blood flow is typically reduced", "Pulmonary hypertension develops from the high pressure left to right flow"],
  ["Classic examples", "Tetralogy of Fallot, TGA, tricuspid atresia, truncus arteriosus, TAPVC", "ASD, VSD, PDA, coarctation, aortic and pulmonary stenosis"],
];
s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 1.6, w: 12.1, h: 0.5, rectRadius: 0.07, fill: { color: NAVY } });
s.addText("Feature", { x: 0.85, y: 1.6, w: 2.6, h: 0.5, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 13.5, bold: true, color: "FFFFFF" });
s.addText("Cyanotic congenital heart disease", { x: 3.6, y: 1.6, w: 4.4, h: 0.5, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 13.5, bold: true, color: "FFFFFF" });
s.addText("Acyanotic congenital heart disease", { x: 8.2, y: 1.6, w: 4.3, h: 0.5, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 13.5, bold: true, color: "FFFFFF" });

y = 2.2;
rows.forEach((r, i) => {
  const h = 0.62;
  if (i % 2 === 0) {
    s.addShape(pres.ShapeType.rect, { x: 0.6, y: y, w: 12.1, h: h, fill: { color: OFFWHITE } });
  }
  s.addText(r[0], { x: 0.85, y: y, w: 2.6, h: h, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 12.5, bold: true, color: CHERRY });
  s.addText(r[1], { x: 3.6, y: y, w: 4.4, h: h, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 12, color: INK });
  s.addText(r[2], { x: 8.2, y: y, w: 4.3, h: h, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 12, color: INK });
  y += h;
});
s.addNotes("This is the core comparison table. Cyanosis, clubbing, polycythemia and growth retardation belong to the cyanotic group. Pulmonary hypertension belongs to the acyanotic left to right shunts.");

/* --------------------------------------------- 9. GENERAL CLINICAL FEATURES */
s = pres.addSlide();
s.background = { color: "FFFFFF" };
titleBar(s, "General clinical features to look for", "Examination");

card(s, { x: 0.6, y: 1.65, w: 5.9, h: 2.35, fill: BLUSH });
s.addText("Congestive heart failure", { x: 0.95, y: 1.85, w: 5.2, h: 0.35, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 17, bold: true, color: CHERRY });
s.addText("In as many as 80% of infants with critical disease, congestive heart failure is the presenting symptom. Difficulty in feeding is common, often with tachypnea, sweating and subcostal retraction. Suspect CHD if feeding takes more than 30 minutes. On examination look for an S3 gallop and crepitations in the lungs.", {
  x: 0.95, y: 2.25, w: 5.25, h: 1.65, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, color: INK,
});

card(s, { x: 6.8, y: 1.65, w: 5.9, h: 2.35, fill: OFFWHITE });
s.addText("Syncope", { x: 7.15, y: 1.85, w: 5.2, h: 0.35, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 17, bold: true, color: CHERRY });
s.addText("Common when severe right or left ventricular outflow tract obstruction is present. Exertional syncope with deepening central cyanosis may occur in Fallot's tetralogy, because exercise increases pulmonary vascular resistance and decreases systemic vascular resistance, so the right to left shunt increases and cerebral oxygenation falls.", {
  x: 7.15, y: 2.25, w: 5.25, h: 1.65, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, color: INK,
});

card(s, { x: 0.6, y: 4.2, w: 12.1, h: 2.1, fill: NAVY });
s.addText("Squatting in Fallot's tetralogy", { x: 0.95, y: 4.4, w: 6, h: 0.35, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 17, bold: true, color: "FFFFFF" });
const sq = ["Squatting posture is adopted by the child", "Venous return decreases, peripheral vascular resistance increases", "Right sided pressure falls, left sided pressure rises", "Right to left shunt is reduced and cerebral oxygenation improves"];
x = 0.95;
sq.forEach((t, i) => {
  s.addShape(pres.ShapeType.roundRect, { x: x, y: 4.85, w: 2.55, h: 1.15, rectRadius: 0.08, fill: { color: "FFFFFF" } });
  s.addText(t, { x: x + 0.18, y: 4.95, w: 2.2, h: 0.95, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 11.5, color: INK });
  if (i < 3) s.addShape(pres.ShapeType.rightArrow, { x: x + 2.63, y: 5.28, w: 0.32, h: 0.3, fill: { color: "E7B9BE" } });
  x += 2.95;
});
s.addNotes("If asked why children squat, give the four step chain: less venous return, more peripheral resistance, less right to left shunt, better cerebral oxygenation.");

/* --------------------------------------------- 10. SUMMARY */
s = pres.addSlide();
s.background = { color: NAVY };
s.addShape(pres.ShapeType.ellipse, { x: -1.6, y: 4.2, w: 5.2, h: 5.2, fill: { color: CHERRY }, transparency: 60 });
s.addText("TAKE HOME", { x: 0.7, y: 0.6, w: 6, h: 0.3, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12, bold: true, charSpacing: 2, color: "E7B9BE" });
s.addText("What to carry away", { x: 0.7, y: 0.95, w: 8, h: 0.6, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 32, bold: true, color: "FFFFFF" });

const take = [
  ["Adults present through complications", "Endocarditis, arrhythmias, right heart failure from pulmonary hypertension, end stage heart failure and sudden cardiac death may all be the first sign of congenital heart disease."],
  ["Cyanotic means right to left", "Central cyanosis, clubbing, polycythemia and growth retardation follow from chronic hypoxemia."],
  ["Acyanotic means left to right or obstruction", "No cyanosis at first, but a long standing shunt raises pulmonary pressure and can reverse into Eisenmenger's syndrome."],
  ["Counsel the woman planning pregnancy", "Fetal ultrasound screening is necessary to rule out heart malformations, since patients with congenital heart disease are more likely to have a baby with congenital heart disease."],
];
y = 1.85;
take.forEach((t, i) => {
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: y, w: 11.9, h: 1.1, rectRadius: 0.08, fill: { color: "FFFFFF" }, transparency: i % 2 === 0 ? 0 : 8 });
  dot(s, 1.0, y + 0.38, String(i + 1), CHERRY);
  s.addText(t[0], { x: 1.55, y: y + 0.16, w: 3.6, h: 0.8, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 14, bold: true, color: NAVY });
  s.addText(t[1], { x: 5.3, y: y + 0.16, w: 7.0, h: 0.8, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 12, color: MUTED });
  y += 1.22;
});
s.addText("Thank you", { x: 0.7, y: 6.85, w: 4, h: 0.4, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 18, bold: true, color: "FFFFFF" });
s.addNotes("Close on the four points, then invite questions.");

pres.writeFile({ fileName: "/home/user/highhsenberg/deck/Congenital-Heart-Disease-in-Adults.pptx" }).then(f => console.log("wrote", f));
