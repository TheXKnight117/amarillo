const pptxgen = require("pptxgenjs");
const PDFDocument = require("pdfkit");
const fs = require("fs");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Estudiante";
pptx.subject = "Actividad formativa - Semana 1";
pptx.title = "Análisis de requerimientos de WhatsApp";
pptx.company = "Instituto Profesional Santo Tomás";
pptx.lang = "es-CL";
pptx.theme = {
  headFontFace: "Aptos Display",
  bodyFontFace: "Aptos",
  lang: "es-CL",
};

const C = {
  green: "128C7E",
  dark: "075E54",
  light: "EAF7F3",
  pale: "F7FAF9",
  ink: "263A3A",
  muted: "58706D",
  white: "FFFFFF",
  line: "CFE4DE",
  accent: "25D366",
};

const slides = [
  {
    title: "Análisis de requerimientos de WhatsApp",
    subtitle: "Actividad formativa · Semana 1",
    cover: true,
  },
  {
    title: "1. Dominio de aplicación",
    intro:
      "WhatsApp pertenece al área de la comunicación digital y permite que personas y organizaciones se comuniquen a distancia.",
    sections: [
      {
        heading: "Necesidad que resuelve",
        bullets: [
          "Facilita una comunicación rápida mediante mensajes, llamadas y archivos.",
          "Permite mantener el contacto individual o grupal desde un dispositivo conectado a internet.",
        ],
      },
      {
        heading: "Usuarios principales",
        bullets: [
          "Personas: conversan y comparten contenido con sus contactos.",
          "Administradores: organizan grupos y comunidades.",
          "Negocios: atienden consultas y se comunican con clientes.",
        ],
      },
    ],
  },
  {
    title: "2. Delimitación del sistema",
    columns: [
      {
        heading: "Dentro del alcance",
        bullets: [
          "Enviar y recibir mensajes de texto y voz.",
          "Realizar llamadas y videollamadas.",
          "Compartir imágenes, videos, documentos y ubicación.",
          "Crear grupos, comunidades y administrar participantes.",
          "Configurar privacidad, bloquear y reportar contactos.",
        ],
      },
      {
        heading: "Fuera del alcance",
        bullets: [
          "Proporcionar la conexión a internet del usuario.",
          "Fabricar o reparar el teléfono donde funciona la app.",
          "Controlar la calidad de la red móvil o Wi-Fi.",
          "Garantizar que el receptor lea o responda un mensaje.",
          "Reemplazar servicios de emergencia o atención presencial.",
        ],
      },
    ],
  },
  {
    title: "3. Restricciones del sistema",
    cards: [
      {
        label: "Tecnológica",
        text: "La aplicación necesita un dispositivo compatible, un número telefónico válido y conexión a internet para operar.",
      },
      {
        label: "Legal y privacidad",
        text: "Debe proteger los datos personales y cumplir la normativa aplicable en los países donde presta el servicio.",
      },
      {
        label: "Rendimiento",
        text: "Debe mantener una respuesta estable ante una gran cantidad de usuarios y distintas calidades de conexión.",
      },
    ],
  },
  {
    title: "4. Requerimientos de usuario",
    note: "Redactados desde lo que necesita la persona, sin explicar detalles técnicos.",
    cards: [
      {
        label: "RU-01",
        text: "Como usuario, quiero enviar mensajes y archivos a mis contactos para comunicarme con ellos.",
      },
      {
        label: "RU-02",
        text: "Como usuario, quiero crear grupos para conversar y compartir información con varias personas.",
      },
      {
        label: "RU-03",
        text: "Como usuario, quiero controlar quién puede contactarme y ver mi información para cuidar mi privacidad.",
      },
    ],
  },
  {
    title: "5. Requerimientos del sistema",
    requirements: [
      {
        id: "RS-01 · Funcional",
        text: "El sistema deberá permitir enviar mensajes de texto, voz y archivos a un contacto seleccionado, mostrando su estado de envío.",
      },
      {
        id: "RS-02 · Funcional",
        text: "El sistema deberá permitir crear grupos, agregar o eliminar participantes y asignar permisos de administración.",
      },
      {
        id: "RS-03 · No funcional (seguridad)",
        text: "El sistema deberá proteger el contenido de mensajes y llamadas mediante cifrado de extremo a extremo.",
      },
    ],
  },
  {
    title: "6. Conclusión",
    conclusion:
      "Esta actividad me permitió entender que una aplicación no se define solamente por sus funciones. También es necesario conocer su contexto, sus límites y las condiciones que debe cumplir. Clasificar bien los requerimientos ayuda a evitar confusiones y permite construir un sistema que responda mejor a las necesidades de sus usuarios.",
    takeaways: [
      "El dominio explica dónde y para quién funciona la aplicación.",
      "Los límites aclaran qué responsabilidades pertenecen al sistema.",
      "Los requerimientos transforman necesidades en características concretas.",
    ],
  },
];

function addHeader(slide, title, page) {
  slide.background = { color: C.pale };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.16,
    fill: { color: C.accent },
    line: { color: C.accent },
  });
  slide.addText(title, {
    x: 0.72,
    y: 0.45,
    w: 11.5,
    h: 0.55,
    fontFace: "Aptos Display",
    fontSize: 26,
    bold: true,
    color: C.dark,
    margin: 0,
  });
  slide.addText(String(page).padStart(2, "0"), {
    x: 12.1,
    y: 0.53,
    w: 0.5,
    h: 0.3,
    fontSize: 10,
    bold: true,
    align: "right",
    color: C.muted,
    margin: 0,
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 0.72,
    y: 1.15,
    w: 11.9,
    h: 0,
    line: { color: C.line, width: 1 },
  });
}

function addFooter(slide) {
  slide.addText("Ingeniería de requerimientos · Semana 1", {
    x: 0.72,
    y: 7.13,
    w: 4.3,
    h: 0.18,
    fontSize: 8,
    color: C.muted,
    margin: 0,
  });
}

function addBullet(slide, text, x, y, w, size = 16) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x,
    y: y + 0.12,
    w: 0.1,
    h: 0.1,
    fill: { color: C.accent },
    line: { color: C.accent },
  });
  slide.addText(text, {
    x: x + 0.22,
    y,
    w: w - 0.22,
    h: 0.5,
    fontSize: size,
    color: C.ink,
    breakLine: false,
    valign: "mid",
    margin: 0,
    fit: "shrink",
  });
}

function addSection(slide, section, x, y, w, h) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    fill: { color: C.white },
    line: { color: C.line, width: 1 },
  });
  slide.addText(section.heading, {
    x: x + 0.35,
    y: y + 0.25,
    w: w - 0.7,
    h: 0.36,
    fontSize: 18,
    bold: true,
    color: C.green,
    margin: 0,
  });
  section.bullets.forEach((text, index) =>
    addBullet(slide, text, x + 0.38, y + 0.82 + index * 0.68, w - 0.76, 14.5),
  );
}

function createCover() {
  const slide = pptx.addSlide();
  slide.background = { color: C.dark };
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 9.15,
    y: -0.9,
    w: 5.3,
    h: 5.3,
    fill: { color: C.green, transparency: 5 },
    line: { color: C.green, transparency: 100 },
  });
  slide.addShape(pptx.ShapeType.ellipse, {
    x: 9.8,
    y: 1.0,
    w: 2.0,
    h: 2.0,
    fill: { color: C.accent },
    line: { color: C.white, width: 3 },
  });
  slide.addText("☎", {
    x: 10.13,
    y: 1.38,
    w: 1.35,
    h: 0.95,
    fontSize: 40,
    bold: true,
    align: "center",
    color: C.white,
    margin: 0,
  });
  slide.addText("ANÁLISIS DE REQUERIMIENTOS", {
    x: 0.85,
    y: 1.2,
    w: 6.9,
    h: 0.35,
    fontSize: 13,
    bold: true,
    charSpacing: 2.1,
    color: C.accent,
    margin: 0,
  });
  slide.addText("WhatsApp", {
    x: 0.85,
    y: 1.75,
    w: 7.2,
    h: 0.9,
    fontSize: 42,
    bold: true,
    color: C.white,
    margin: 0,
  });
  slide.addText("Actividad formativa · Semana 1", {
    x: 0.88,
    y: 2.8,
    w: 6,
    h: 0.4,
    fontSize: 18,
    color: "D7EFEB",
    margin: 0,
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 0.88,
    y: 4.15,
    w: 5.8,
    h: 0,
    line: { color: "55BCAA", width: 1 },
  });
  slide.addText("Nombre: [ESCRIBE TU NOMBRE]\nRUT: [ESCRIBE TU RUT]\nAsignatura: Ingeniería de requerimientos", {
    x: 0.88,
    y: 4.45,
    w: 6.8,
    h: 1.25,
    fontSize: 16,
    color: C.white,
    breakLine: false,
    margin: 0,
    breakLineOnOverflow: false,
  });
  slide.addText("INSTITUTO PROFESIONAL SANTO TOMÁS", {
    x: 0.88,
    y: 6.72,
    w: 5.6,
    h: 0.25,
    fontSize: 9,
    bold: true,
    charSpacing: 1.2,
    color: "B8D8D2",
    margin: 0,
  });
}

function createSlides() {
  createCover();

  const domain = pptx.addSlide();
  addHeader(domain, slides[1].title, 2);
  domain.addText(slides[1].intro, {
    x: 0.82,
    y: 1.45,
    w: 11.7,
    h: 0.72,
    fontSize: 19,
    color: C.ink,
    bold: false,
    margin: 0,
    fit: "shrink",
  });
  addSection(domain, slides[1].sections[0], 0.82, 2.4, 5.72, 3.75);
  addSection(domain, slides[1].sections[1], 6.79, 2.4, 5.72, 3.75);
  addFooter(domain);

  const boundary = pptx.addSlide();
  addHeader(boundary, slides[2].title, 3);
  addSection(boundary, slides[2].columns[0], 0.82, 1.48, 5.72, 5.25);
  addSection(boundary, slides[2].columns[1], 6.79, 1.48, 5.72, 5.25);
  addFooter(boundary);

  const restrictions = pptx.addSlide();
  addHeader(restrictions, slides[3].title, 4);
  slides[3].cards.forEach((card, i) => {
    const x = 0.82 + i * 4.02;
    restrictions.addShape(pptx.ShapeType.roundRect, {
      x,
      y: 1.6,
      w: 3.72,
      h: 4.85,
      fill: { color: i === 1 ? C.light : C.white },
      line: { color: C.line, width: 1 },
    });
    restrictions.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.32,
      y: 1.95,
      w: 0.52,
      h: 0.52,
      fill: { color: C.green },
      line: { color: C.green },
    });
    restrictions.addText(String(i + 1), {
      x: x + 0.32,
      y: 2.05,
      w: 0.52,
      h: 0.22,
      fontSize: 12,
      bold: true,
      color: C.white,
      align: "center",
      margin: 0,
    });
    restrictions.addText(card.label, {
      x: x + 0.35,
      y: 2.75,
      w: 3.0,
      h: 0.48,
      fontSize: 19,
      bold: true,
      color: C.dark,
      margin: 0,
      align: "center",
    });
    restrictions.addText(card.text, {
      x: x + 0.45,
      y: 3.55,
      w: 2.82,
      h: 1.85,
      fontSize: 15.5,
      color: C.ink,
      margin: 0,
      valign: "mid",
      align: "center",
      fit: "shrink",
    });
  });
  addFooter(restrictions);

  const userReq = pptx.addSlide();
  addHeader(userReq, slides[4].title, 5);
  userReq.addText(slides[4].note, {
    x: 0.82,
    y: 1.35,
    w: 11.6,
    h: 0.42,
    fontSize: 14,
    italic: true,
    color: C.muted,
    margin: 0,
  });
  slides[4].cards.forEach((card, i) => {
    const y = 2.02 + i * 1.5;
    userReq.addShape(pptx.ShapeType.roundRect, {
      x: 0.9,
      y,
      w: 11.5,
      h: 1.15,
      fill: { color: i % 2 === 0 ? C.white : C.light },
      line: { color: C.line, width: 1 },
    });
    userReq.addText(card.label, {
      x: 1.15,
      y: y + 0.35,
      w: 0.85,
      h: 0.3,
      fontSize: 14,
      bold: true,
      color: C.green,
      margin: 0,
    });
    userReq.addText(card.text, {
      x: 2.1,
      y: y + 0.23,
      w: 9.85,
      h: 0.66,
      fontSize: 17,
      color: C.ink,
      valign: "mid",
      margin: 0,
      fit: "shrink",
    });
  });
  addFooter(userReq);

  const systemReq = pptx.addSlide();
  addHeader(systemReq, slides[5].title, 6);
  slides[5].requirements.forEach((req, i) => {
    const y = 1.55 + i * 1.65;
    systemReq.addShape(pptx.ShapeType.roundRect, {
      x: 0.88,
      y,
      w: 11.58,
      h: 1.28,
      fill: { color: C.white },
      line: { color: C.line, width: 1 },
    });
    systemReq.addShape(pptx.ShapeType.rect, {
      x: 0.88,
      y,
      w: 0.12,
      h: 1.28,
      fill: { color: i === 2 ? C.accent : C.green },
      line: { color: i === 2 ? C.accent : C.green },
    });
    systemReq.addText(req.id, {
      x: 1.25,
      y: y + 0.2,
      w: 2.55,
      h: 0.32,
      fontSize: 13.5,
      bold: true,
      color: C.green,
      margin: 0,
    });
    systemReq.addText(req.text, {
      x: 3.65,
      y: y + 0.17,
      w: 8.25,
      h: 0.88,
      fontSize: 15.5,
      color: C.ink,
      valign: "mid",
      margin: 0,
      fit: "shrink",
    });
  });
  addFooter(systemReq);

  const conclusion = pptx.addSlide();
  addHeader(conclusion, slides[6].title, 7);
  conclusion.addShape(pptx.ShapeType.roundRect, {
    x: 0.85,
    y: 1.5,
    w: 7.25,
    h: 4.9,
    fill: { color: C.white },
    line: { color: C.line, width: 1 },
  });
  conclusion.addText(slides[6].conclusion, {
    x: 1.27,
    y: 2.0,
    w: 6.4,
    h: 3.8,
    fontSize: 19,
    color: C.ink,
    valign: "mid",
    margin: 0,
    breakLineOnOverflow: false,
    fit: "shrink",
  });
  conclusion.addText("Ideas principales", {
    x: 8.65,
    y: 1.65,
    w: 3.4,
    h: 0.4,
    fontSize: 19,
    bold: true,
    color: C.dark,
    margin: 0,
  });
  slides[6].takeaways.forEach((text, i) =>
    addBullet(conclusion, text, 8.65, 2.4 + i * 1.2, 3.65, 15),
  );
  addFooter(conclusion);
}

function makePdf() {
  const doc = new PDFDocument({
    size: [960, 540],
    margin: 0,
    info: {
      Title: "Análisis de requerimientos de WhatsApp",
      Author: "Estudiante",
      Subject: "Actividad formativa - Semana 1",
    },
  });
  doc.pipe(fs.createWriteStream("Analisis_requerimientos_WhatsApp.pdf"));

  function pageBase(title, number) {
    if (number > 1) doc.addPage();
    doc.rect(0, 0, 960, 540).fill(`#${C.pale}`);
    doc.rect(0, 0, 960, 12).fill(`#${C.accent}`);
    doc.fillColor(`#${C.dark}`).font("Helvetica-Bold").fontSize(25).text(title, 52, 35);
    doc.fillColor(`#${C.muted}`).fontSize(9).text(String(number).padStart(2, "0"), 875, 42);
    doc.moveTo(52, 82).lineTo(908, 82).strokeColor(`#${C.line}`).stroke();
    doc.fillColor(`#${C.muted}`).font("Helvetica").fontSize(8).text("Ingeniería de requerimientos · Semana 1", 52, 512);
  }

  function box(x, y, w, h, fill = C.white) {
    doc.roundedRect(x, y, w, h, 8).fillAndStroke(`#${fill}`, `#${C.line}`);
  }

  function pdfBullets(items, x, y, w, gap = 45, size = 13) {
    items.forEach((item, i) => {
      const yy = y + i * gap;
      doc.circle(x + 4, yy + 7, 3.5).fill(`#${C.accent}`);
      doc.fillColor(`#${C.ink}`).font("Helvetica").fontSize(size).text(item, x + 16, yy, {
        width: w - 16,
        lineGap: 2,
      });
    });
  }

  doc.rect(0, 0, 960, 540).fill(`#${C.dark}`);
  doc.circle(790, 80, 175).fill(`#${C.green}`);
  doc.circle(790, 135, 62).fillAndStroke(`#${C.accent}`, "#FFFFFF");
  doc.fillColor("#FFFFFF").font("Helvetica-Bold").fontSize(39).text("☎", 755, 110, { width: 70, align: "center" });
  doc.fillColor(`#${C.accent}`).font("Helvetica-Bold").fontSize(12).text("ANÁLISIS DE REQUERIMIENTOS", 62, 92, {
    characterSpacing: 1.4,
  });
  doc.fillColor("#FFFFFF").fontSize(42).text("WhatsApp", 62, 130);
  doc.fillColor("#D7EFEB").font("Helvetica").fontSize(17).text("Actividad formativa · Semana 1", 64, 195);
  doc.moveTo(64, 292).lineTo(480, 292).strokeColor("#55BCAA").stroke();
  doc.fillColor("#FFFFFF").fontSize(14).text(
    "Nombre: [ESCRIBE TU NOMBRE]\nRUT: [ESCRIBE TU RUT]\nAsignatura: Ingeniería de requerimientos",
    64,
    320,
    { lineGap: 8 },
  );
  doc.fillColor("#B8D8D2").font("Helvetica-Bold").fontSize(8).text("INSTITUTO PROFESIONAL SANTO TOMÁS", 64, 490);

  pageBase(slides[1].title, 2);
  doc.fillColor(`#${C.ink}`).font("Helvetica").fontSize(17).text(slides[1].intro, 58, 105, { width: 840 });
  slides[1].sections.forEach((section, i) => {
    const x = 58 + i * 438;
    box(x, 166, 410, 300);
    doc.fillColor(`#${C.green}`).font("Helvetica-Bold").fontSize(17).text(section.heading, x + 24, 190);
    pdfBullets(section.bullets, x + 25, 238, 360, 66, 13);
  });

  pageBase(slides[2].title, 3);
  slides[2].columns.forEach((section, i) => {
    const x = 58 + i * 438;
    box(x, 108, 410, 382);
    doc.fillColor(`#${C.green}`).font("Helvetica-Bold").fontSize(17).text(section.heading, x + 24, 132);
    pdfBullets(section.bullets, x + 25, 178, 360, 55, 12.5);
  });

  pageBase(slides[3].title, 4);
  slides[3].cards.forEach((card, i) => {
    const x = 58 + i * 286;
    box(x, 120, 262, 340, i === 1 ? C.light : C.white);
    doc.circle(x + 30, 152, 16).fill(`#${C.green}`);
    doc.fillColor("#FFFFFF").font("Helvetica-Bold").fontSize(11).text(String(i + 1), x + 24, 146, { width: 12, align: "center" });
    doc.fillColor(`#${C.dark}`).fontSize(16).text(card.label, x + 25, 200, { width: 212, align: "center" });
    doc.fillColor(`#${C.ink}`).font("Helvetica").fontSize(13.5).text(card.text, x + 30, 260, {
      width: 202,
      align: "center",
      lineGap: 4,
    });
  });

  pageBase(slides[4].title, 5);
  doc.fillColor(`#${C.muted}`).font("Helvetica-Oblique").fontSize(12).text(slides[4].note, 58, 100);
  slides[4].cards.forEach((card, i) => {
    const y = 140 + i * 105;
    box(64, y, 832, 76, i % 2 === 0 ? C.white : C.light);
    doc.fillColor(`#${C.green}`).font("Helvetica-Bold").fontSize(12).text(card.label, 86, y + 29);
    doc.fillColor(`#${C.ink}`).font("Helvetica").fontSize(14.5).text(card.text, 160, y + 20, {
      width: 700,
      lineGap: 2,
    });
  });

  pageBase(slides[5].title, 6);
  slides[5].requirements.forEach((req, i) => {
    const y = 112 + i * 112;
    box(64, y, 832, 88);
    doc.rect(64, y, 8, 88).fill(`#${i === 2 ? C.accent : C.green}`);
    doc.fillColor(`#${C.green}`).font("Helvetica-Bold").fontSize(12).text(req.id, 90, y + 20, { width: 175 });
    doc.fillColor(`#${C.ink}`).font("Helvetica").fontSize(13.5).text(req.text, 270, y + 16, {
      width: 590,
      lineGap: 2,
    });
  });

  pageBase(slides[6].title, 7);
  box(58, 112, 520, 345);
  doc.fillColor(`#${C.ink}`).font("Helvetica").fontSize(16.5).text(slides[6].conclusion, 88, 165, {
    width: 460,
    lineGap: 7,
    align: "left",
  });
  doc.fillColor(`#${C.dark}`).font("Helvetica-Bold").fontSize(17).text("Ideas principales", 625, 130);
  pdfBullets(slides[6].takeaways, 625, 185, 275, 78, 13);

  doc.end();
}

createSlides();
Promise.all([
  pptx.writeFile({ fileName: "Analisis_requerimientos_WhatsApp.pptx" }),
  Promise.resolve(makePdf()),
]).catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
