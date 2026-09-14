const pptxgen = require("pptxgenjs");

const CHERRY = "990011";
const NAVY = "2F3C7E";
const OFFWHITE = "FCF6F5";
const INK = "231F20";
const MUTED = "6B6168";
const BLUSH = "F6E4E6";
const ROSE = "E7B9BE";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "Monika R";
pres.title = "Adult Presentation of Congenital Heart Disease";

const HEAD = "Bookman Old Style";
const BODY = "Calibri";

function titleBar(slide, kicker, text) {
  slide.addText(kicker.toUpperCase(), {
    x: 0.6, y: 0.36, w: 11, h: 0.28, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 12, bold: true, charSpacing: 2, color: CHERRY,
  });
  slide.addText(text, {
    x: 0.6, y: 0.66, w: 12.1, h: 0.72, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 30, bold: true, color: NAVY,
  });
}

function dot(slide, x, y, label, fill, d) {
  const s = d || 0.36;
  slide.addShape(pres.ShapeType.ellipse, { x: x, y: y, w: s, h: s, fill: { color: fill || CHERRY } });
  slide.addText(label, {
    x: x, y: y, w: s, h: s, isTextBox: true, margin: 0, align: "center",
    valign: "middle", fontFace: BODY, fontSize: s > 0.4 ? 16 : 13.5, bold: true, color: "FFFFFF",
  });
}

function card(slide, o) {
  slide.addShape(pres.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.09,
    fill: { color: o.fill || OFFWHITE },
    shadow: { type: "outer", blur: 9, offset: 1, angle: 90, color: "9A8A8A", opacity: 0.18 },
  });
}

function source(slide, dark) {
  slide.addText("Manipal Prep Manual of Medicine, Manthappa M, 3rd edition", {
    x: 0.6, y: 6.92, w: 8, h: 0.28, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 10, italic: true, color: dark ? "9A93A8" : "A2989C",
  });
}

/* ===================================================== 1. TITLE */
let s = pres.addSlide();
s.background = { color: NAVY };
s.addShape(pres.ShapeType.ellipse, { x: 9.6, y: -1.9, w: 6.4, h: 6.4, fill: { color: CHERRY }, transparency: 50 });
s.addShape(pres.ShapeType.ellipse, { x: 11.2, y: 3.9, w: 4.4, h: 4.4, fill: { color: "FFFFFF" }, transparency: 88 });

s.addShape(pres.ShapeType.roundRect, { x: 0.8, y: 0.85, w: 4.15, h: 0.5, rectRadius: 0.1, fill: { color: CHERRY } });
s.addText("COMPETENCY 1.28  |  IM 1.28", {
  x: 0.8, y: 0.85, w: 4.15, h: 0.5, isTextBox: true, margin: 0, align: "center", valign: "middle",
  fontFace: BODY, fontSize: 12.5, bold: true, charSpacing: 1.5, color: "FFFFFF",
});

s.addText("Enumerate the causes of adult presentation of congenital heart disease and describe the distinguishing features between cyanotic and acyanotic heart disease", {
  x: 0.8, y: 1.65, w: 9.4, h: 2.7, isTextBox: true, margin: 0, lineSpacing: 40,
  fontFace: HEAD, fontSize: 29, bold: true, color: "FFFFFF",
});

s.addShape(pres.ShapeType.roundRect, { x: 0.8, y: 4.85, w: 4.3, h: 1.25, rectRadius: 0.1, fill: { color: "FFFFFF" } });
s.addText("Monika R", {
  x: 1.1, y: 5.02, w: 3.8, h: 0.42, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 21, bold: true, color: NAVY,
});
s.addText("11231126", {
  x: 1.1, y: 5.45, w: 3.8, h: 0.3, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 15, bold: true, charSpacing: 1.5, color: CHERRY,
});
s.addText("MBBS Fourth Year  •  General Medicine  •  Diseases of the Cardiovascular System", {
  x: 0.8, y: 6.35, w: 9, h: 0.3, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 12, color: "C8C1D4",
});
s.addNotes("Competency 1.28. Two parts to answer: the causes of adult presentation, then the cyanotic versus acyanotic comparison.");

/* ===================================================== 2. WHAT IS CHD */
s = pres.addSlide();
s.background = { color: "FFFFFF" };
titleBar(s, "Definition", "What is congenital heart disease?");

card(s, { x: 0.6, y: 1.6, w: 12.1, h: 1.5, fill: BLUSH });
s.addText("Congenital heart disease is due to abnormal development of a normal structure, or failure of a normal structure to develop fully.", {
  x: 1.0, y: 1.85, w: 11.3, h: 1.0, isTextBox: true, margin: 0, valign: "middle",
  fontFace: HEAD, fontSize: 20, bold: true, color: NAVY,
});

const facts = [
  ["1%", "of live births are affected"],
  ["Males", "affected more commonly, except ASD and PDA which are more common in females"],
  ["Early", "recognition matters, since early treatment gives a better outcome"],
];
let x = 0.6;
facts.forEach((f, i) => {
  card(s, { x: x, y: 3.35, w: 3.9, h: 2.6, fill: i === 0 ? NAVY : OFFWHITE });
  s.addText(f[0], {
    x: x + 0.4, y: 3.7, w: 3.1, h: 0.85, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 36, bold: true, color: i === 0 ? "FFFFFF" : CHERRY,
  });
  s.addText(f[1], {
    x: x + 0.4, y: 4.6, w: 3.1, h: 1.2, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 14, color: i === 0 ? ROSE : INK,
  });
  x += 4.1;
});
source(s);
s.addNotes("Define it in one line, then the three facts: 1% of live births, male predominance except ASD and PDA, and the value of early recognition.");

/* ===================================================== 3. ETIOLOGY */
s = pres.addSlide();
s.background = { color: "FFFFFF" };
titleBar(s, "Etiology", "Why does congenital heart disease occur?");
s.addText("Maldevelopment is due to multifactorial genetic and environmental causes. The recognized risk factors are:", {
  x: 0.6, y: 1.48, w: 12.1, h: 0.32, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 14.5, italic: true, color: MUTED,
});

const eti = [
  ["Maternal infections", "Rubella infection, causing persistent ductus arteriosus and pulmonary stenosis"],
  ["Drugs and radiation", "Alcohol abuse causing septal defects, phenytoin associated with pulmonary stenosis, and radiation"],
  ["Genetic abnormalities", "Familial form of atrial septal defect, and congenital heart block"],
  ["Chromosomal abnormalities", "Septal defects and tetralogy of Fallot in Down's syndrome (trisomy 21); coarctation of the aorta in Turner's syndrome (45, XO)"],
];
x = 0.6;
eti.forEach((e, i) => {
  card(s, { x: x, y: 2.05, w: 2.95, h: 3.9, fill: i % 2 === 0 ? OFFWHITE : BLUSH });
  dot(s, x + 0.4, 2.4, String(i + 1), CHERRY, 0.5);
  s.addText(e[0], {
    x: x + 0.4, y: 3.1, w: 2.25, h: 0.85, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 15.5, bold: true, color: NAVY,
  });
  s.addText(e[1], {
    x: x + 0.4, y: 4.0, w: 2.25, h: 1.75, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 12.5, color: INK,
  });
  x += 3.11;
});
source(s);
s.addNotes("Four headings: maternal infections, drugs and radiation, genetic abnormalities, chromosomal abnormalities.");

/* ===================================================== 4. WHY IN ADULTS */
s = pres.addSlide();
s.background = { color: NAVY };
s.addText("SETTING THE SCENE", {
  x: 0.6, y: 0.36, w: 11, h: 0.28, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 12, bold: true, charSpacing: 2, color: ROSE,
});
s.addText("Why do we see congenital heart disease in adults?", {
  x: 0.6, y: 0.66, w: 12.1, h: 0.72, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 30, bold: true, color: "FFFFFF",
});

s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 1.65, w: 12.1, h: 1.35, rectRadius: 0.1, fill: { color: CHERRY } });
s.addText("Because of improved medical and surgical management, more children with congenital heart disease are surviving into adolescence and adulthood.", {
  x: 1.0, y: 1.85, w: 11.3, h: 0.95, isTextBox: true, margin: 0, valign: "middle",
  fontFace: HEAD, fontSize: 19, bold: true, color: "FFFFFF",
});

const why = [
  ["Survivors of childhood disease", "Operated or palliated patients now reach adult clinics and need lifelong follow up."],
  ["Never diagnosed in childhood", "A mild or silent lesion is missed, and the disease declares itself later through its complications."],
  ["Eisenmenger's syndrome", "A long standing left to right shunt raises pulmonary pressure until flow reverses right to left, and the previously pink patient turns cyanosed."],
];
x = 0.6;
why.forEach((w, i) => {
  s.addShape(pres.ShapeType.roundRect, { x: x, y: 3.3, w: 3.9, h: 3.05, rectRadius: 0.1, fill: { color: "FFFFFF" } });
  dot(s, x + 0.4, 3.62, String(i + 1), NAVY, 0.5);
  s.addText(w[0], {
    x: x + 0.4, y: 4.32, w: 3.1, h: 0.75, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 15.5, bold: true, color: CHERRY,
  });
  s.addText(w[1], {
    x: x + 0.4, y: 5.1, w: 3.1, h: 1.15, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 12.5, color: INK,
  });
  x += 4.1;
});
source(s, true);
s.addNotes("One sentence is the whole answer: improved management means more survivors. Then add the undiagnosed patient and Eisenmenger's syndrome.");

/* ===================================================== 5. CAUSES OF ADULT PRESENTATION */
s = pres.addSlide();
s.background = { color: "FFFFFF" };
titleBar(s, "Enumeration", "Causes of adult presentation of congenital heart disease");
s.addText("Any of these may be the first time that the presence of congenital heart disease is noted.", {
  x: 0.6, y: 1.48, w: 12.1, h: 0.32, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 14.5, italic: true, color: MUTED,
});

const causes = [
  ["Infective endocarditis", "At the sites of shunts and damaged valves"],
  ["Atrial and ventricular arrhythmias", "Palpitations in a previously undiagnosed adult"],
  ["Right heart failure due to pulmonary hypertension", "Follows a long standing left to right shunt"],
  ["End stage heart failure", "The late outcome of an uncorrected lesion"],
  ["Sudden cardiac death", "May be the very first event that reveals the diagnosis"],
  ["Paradoxical embolism", "Thrombus crosses from systemic veins to arteries through a defect such as VSD, causing stroke and abscesses"],
  ["Polycythemia and thromboembolism", "Chronic hypoxemia leads to hyperviscosity and stroke"],
  ["Pregnancy", "Detected on assessment; unsafe when pulmonary hypertension is present"],
];
let y = 2.0;
causes.forEach((c, i) => {
  const col = i < 4 ? 0.6 : 6.85;
  const row = i % 4;
  const ty = 2.0 + row * 1.2;
  card(s, { x: col, y: ty, w: 5.85, h: 1.05, fill: i < 5 ? BLUSH : OFFWHITE });
  dot(s, col + 0.25, ty + 0.34, String(i + 1), i < 5 ? CHERRY : NAVY);
  s.addText(c[0], {
    x: col + 0.75, y: ty + 0.14, w: 4.9, h: 0.36, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 14, bold: true, color: NAVY,
  });
  s.addText(c[1], {
    x: col + 0.75, y: ty + 0.5, w: 4.9, h: 0.45, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 11.5, color: MUTED,
  });
});
source(s);
s.addNotes("First five are the core list from the book: endocarditis, arrhythmias, right heart failure from pulmonary hypertension, end stage heart failure, sudden cardiac death. Then paradoxical embolism, polycythemia and pregnancy.");

/* ===================================================== 6. CLASSIFICATION */
s = pres.addSlide();
s.background = { color: "FFFFFF" };
titleBar(s, "Table 3.20", "Classification of congenital heart diseases");

s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 1.6, w: 5.9, h: 4.75, rectRadius: 0.1, fill: { color: NAVY } });
s.addText("CYANOTIC", { x: 1.0, y: 1.9, w: 5, h: 0.42, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 23, bold: true, color: "FFFFFF" });
s.addText("Right to left shunting, or mixing of systemic and pulmonary blood flow", {
  x: 1.0, y: 2.35, w: 5.1, h: 0.5, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, italic: true, color: ROSE,
});
s.addText([
  { text: "Tetralogy of Fallot", options: { bullet: true, breakLine: true } },
  { text: "Transposition of the great vessels", options: { bullet: true, breakLine: true } },
  { text: "Tricuspid atresia", options: { bullet: true, breakLine: true } },
  { text: "Truncus arteriosus", options: { bullet: true, breakLine: true } },
  { text: "Total anomalous pulmonary venous connection (TAPVC)", options: { bullet: true } },
], { x: 1.0, y: 2.95, w: 5.1, h: 2.2, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14.5, color: "FFFFFF", paraSpaceAfter: 7 });
s.addShape(pres.ShapeType.roundRect, { x: 1.0, y: 5.45, w: 4.6, h: 0.6, rectRadius: 0.1, fill: { color: CHERRY } });
s.addText("Mnemonic: the five Ts", {
  x: 1.0, y: 5.45, w: 4.6, h: 0.6, isTextBox: true, margin: 0, align: "center", valign: "middle",
  fontFace: BODY, fontSize: 13.5, bold: true, color: "FFFFFF",
});

s.addShape(pres.ShapeType.roundRect, { x: 6.8, y: 1.6, w: 5.9, h: 4.75, rectRadius: 0.1, fill: { color: BLUSH } });
s.addText("ACYANOTIC", { x: 7.2, y: 1.9, w: 5, h: 0.42, isTextBox: true, margin: 0, fontFace: HEAD, fontSize: 23, bold: true, color: CHERRY });
s.addText("Left to right shunt, or an obstructive lesion", {
  x: 7.2, y: 2.35, w: 5.1, h: 0.5, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 12.5, italic: true, color: MUTED,
});
s.addText([
  { text: "Atrial septal defect (ASD)", options: { bullet: true, breakLine: true } },
  { text: "Ventricular septal defect (VSD)", options: { bullet: true, breakLine: true } },
  { text: "Patent ductus arteriosus (PDA)", options: { bullet: true, breakLine: true } },
  { text: "Coarctation of the aorta", options: { bullet: true, breakLine: true } },
  { text: "Aortic stenosis", options: { bullet: true, breakLine: true } },
  { text: "Congenital pulmonary valve regurgitation", options: { bullet: true, breakLine: true } },
  { text: "Pulmonary stenosis", options: { bullet: true } },
], { x: 7.2, y: 2.95, w: 5.1, h: 3.1, isTextBox: true, margin: 0, fontFace: BODY, fontSize: 14, color: INK, paraSpaceAfter: 5 });
source(s);
s.addNotes("Five Ts on the cyanotic side. Acyanotic side is the septal defects, the duct, and the obstructive lesions.");

/* ===================================================== 7. DISTINGUISHING FEATURES */
s = pres.addSlide();
s.background = { color: "FFFFFF" };
titleBar(s, "Comparison", "Distinguishing features: cyanotic vs acyanotic");

const rows = [
  ["Shunt direction", "Right to left, or mixing of circulations", "Left to right, or an obstruction to outflow"],
  ["Central cyanosis", "Present", "Absent, until the shunt reverses"],
  ["Clubbing", "Present, from prolonged cyanosis", "Absent"],
  ["Polycythemia", "Present, from chronic hypoxemia", "Absent"],
  ["Growth", "Growth retardation is common", "Usually preserved"],
  ["Pulmonary blood flow", "Typically reduced", "Increased, leading to pulmonary hypertension"],
  ["Squatting and syncope", "Seen in Fallot's tetralogy", "Not a feature"],
  ["Examples", "TOF, TGA, tricuspid atresia, truncus arteriosus, TAPVC", "ASD, VSD, PDA, coarctation, aortic and pulmonary stenosis"],
];
s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 1.55, w: 12.1, h: 0.52, rectRadius: 0.07, fill: { color: NAVY } });
s.addText("Feature", { x: 0.85, y: 1.55, w: 2.7, h: 0.52, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 13.5, bold: true, color: "FFFFFF" });
s.addText("Cyanotic", { x: 3.7, y: 1.55, w: 4.3, h: 0.52, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 13.5, bold: true, color: "FFFFFF" });
s.addText("Acyanotic", { x: 8.2, y: 1.55, w: 4.3, h: 0.52, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 13.5, bold: true, color: "FFFFFF" });

y = 2.15;
rows.forEach((r, i) => {
  const h = 0.58;
  if (i % 2 === 0) s.addShape(pres.ShapeType.rect, { x: 0.6, y: y, w: 12.1, h: h, fill: { color: OFFWHITE } });
  s.addText(r[0], { x: 0.85, y: y, w: 2.7, h: h, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 12.5, bold: true, color: CHERRY });
  s.addText(r[1], { x: 3.7, y: y, w: 4.3, h: h, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 12, color: INK });
  s.addText(r[2], { x: 8.2, y: y, w: 4.3, h: h, isTextBox: true, margin: 0, valign: "middle", fontFace: BODY, fontSize: 12, color: INK });
  y += h;
});
source(s);
s.addNotes("Cyanosis, clubbing, polycythemia and growth retardation all belong to the cyanotic group. Pulmonary hypertension belongs to the acyanotic left to right shunts.");

/* ===================================================== 8. THANK YOU */
s = pres.addSlide();
s.background = { color: NAVY };
s.addShape(pres.ShapeType.ellipse, { x: -2.0, y: 3.4, w: 6.4, h: 6.4, fill: { color: CHERRY }, transparency: 55 });
s.addShape(pres.ShapeType.ellipse, { x: 10.4, y: -1.4, w: 4.6, h: 4.6, fill: { color: "FFFFFF" }, transparency: 90 });

s.addText("Thank you", {
  x: 3.4, y: 2.55, w: 7, h: 1.1, isTextBox: true, margin: 0, align: "center",
  fontFace: HEAD, fontSize: 54, bold: true, color: "FFFFFF",
});
s.addText("Questions are welcome", {
  x: 3.4, y: 3.7, w: 7, h: 0.4, isTextBox: true, margin: 0, align: "center",
  fontFace: BODY, fontSize: 17, color: ROSE,
});
s.addShape(pres.ShapeType.roundRect, { x: 4.9, y: 4.5, w: 3.95, h: 1.1, rectRadius: 0.1, fill: { color: "FFFFFF" } });
s.addText("Monika R", { x: 4.9, y: 4.62, w: 3.95, h: 0.4, isTextBox: true, margin: 0, align: "center", fontFace: HEAD, fontSize: 19, bold: true, color: NAVY });
s.addText("11231126", { x: 4.9, y: 5.03, w: 3.95, h: 0.3, isTextBox: true, margin: 0, align: "center", fontFace: BODY, fontSize: 14, bold: true, charSpacing: 1.5, color: CHERRY });
source(s, true);
s.addNotes("Close here and take questions.");

pres.writeFile({ fileName: "/home/user/highhsenberg/deck/Congenital-Heart-Disease-in-Adults.pptx" }).then(f => console.log("wrote", f));
