const express = require('express');
const cors = require('cors');
const { 
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, BorderStyle, WidthType, ShadingType, AlignmentType,
  VerticalAlign, PageBreak
} = require('docx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

const PORT = process.env.PORT || 3000;

// ── COULEURS BALKIRA ──────────────────────────────────────────
const C = {
  navy: '0D1B2A', navyMid: '1A2B3C', orange: 'F4821F',
  orangeLight: 'FEF0E6', gray: 'F5F5F5', grayBorder: 'CCCCCC',
  grayText: '888888', white: 'FFFFFF', amber: 'FFF8E1',
  amberText: '856404', redText: 'C0392B', black: '1A1A1A',
  tealLight: 'E6FAF7'
};

const border = (color = C.grayBorder) => ({ style: BorderStyle.SINGLE, size: 1, color });
const borders = { top: border(), bottom: border(), left: border(), right: border() };
const noBorder = { style: BorderStyle.NONE, size: 0, color: C.white };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function txt(text, opts = {}) {
  return new TextRun({ text: String(text || ''), font: 'Arial', size: opts.size || 20,
    bold: opts.bold || false, italics: opts.italic || false, color: opts.color || C.black });
}
function spacer() {
  return new Paragraph({ children: [txt('')], spacing: { before: 40, after: 40 } });
}
function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 100 },
    children: [txt(text, { bold: true, size: 24, color: C.navy })] });
}
function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 180, after: 80 },
    children: [txt(text, { bold: true, size: 21, color: C.navyMid })] });
}
function divider() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.orange } },
    spacing: { before: 100, after: 100 }, children: [txt('')]
  });
}

function labelCell(text, w = 2800) {
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: { fill: C.gray, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [txt(text, { bold: true, size: 18, color: C.navyMid })] })]
  });
}

function valueCell(text, w = 6560, opts = {}) {
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [txt(text, { size: 18, italic: opts.italic, color: opts.color })] })]
  });
}

function hCell(text, w) {
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: { fill: C.navy, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [txt(text, { bold: true, size: 18, color: C.white })] })]
  });
}

function noteBox(text, type = 'info') {
  const colors = {
    info: { bg: C.tealLight, text: '0A7A65', border: '00C9A7' },
    warn: { bg: C.amber, text: C.amberText, border: 'F4A020' },
    v2: { bg: C.orangeLight, text: 'B85D0A', border: C.orange }
  };
  const c = colors[type] || colors.info;
  return new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
    rows: [new TableRow({ children: [new TableCell({
      borders: { top: { style: BorderStyle.SINGLE, size: 4, color: c.border },
        bottom: noBorder, left: { style: BorderStyle.SINGLE, size: 12, color: c.border }, right: noBorder },
      shading: { fill: c.bg, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 200, right: 200 },
      children: [new Paragraph({ children: [txt(text, { size: 18, color: c.text })] })]
    })]})],
  });
}

// ── HELPERS POUR CASES À COCHER ───────────────────────────────
function checkboxLine(options, selectedValue) {
  // options = [{label, value}], selectedValue = valeur à cocher
  const parts = [];
  options.forEach((opt, i) => {
    const isChecked = selectedValue && 
      (selectedValue.toLowerCase().includes(opt.value.toLowerCase()) ||
       opt.value.toLowerCase().includes(selectedValue.toLowerCase().split(' ')[0]));
    const box = isChecked ? '☑' : '☐';
    if (i > 0) parts.push(txt('    '));
    parts.push(txt(box + ' ', { bold: isChecked, color: isChecked ? C.orange : C.black }));
    parts.push(txt(opt.label, { bold: isChecked, color: isChecked ? C.orange : C.black }));
  });
  return parts;
}

// ── PARSER TEXTE MULTI-LIGNES EN PARAGRAPHES ─────────────────
function textToParagraphs(text, indent = false) {
  if (!text) return [new Paragraph({ children: [txt('<À compléter>', { italic: true, color: C.grayText })] })];
  const lines = text.split(/\\n|\n/).filter(l => l.trim());
  if (lines.length === 0) return [new Paragraph({ children: [txt('<À compléter>', { italic: true, color: C.grayText })] })];
  return lines.map(line => new Paragraph({
    indent: indent ? { left: 360 } : undefined,
    spacing: { before: 40, after: 40 },
    children: [txt(line.trim(), { size: 18 })]
  }));
}

// ── GÉNÉRATION ANOMALY FORM ───────────────────────────────────
async function generateAnomalyFormV1(formData, requestId, language) {
  const lang = language === 'fr';
  const date = new Date().toLocaleDateString('fr-FR');

  // Mapping des priorités
  const priorityMap = {
    'critical': 'Critical', 'high': 'High', 'medium': 'Medium', 'low': 'Low',
    'critique': 'Critical', 'élevée': 'High', 'moyenne': 'Medium', 'faible': 'Low',
    'bloquant': 'Critical', 'impact majeur': 'High', 'impact modéré': 'Medium', 'impact mineur': 'Low'
  };
  const priorityRaw = (formData.priority || '').toLowerCase();
  let priorityNorm = 'Low';
  for (const [key, val] of Object.entries(priorityMap)) {
    if (priorityRaw.includes(key)) { priorityNorm = val; break; }
  }

  // Parser les étapes de reproduction
  const steps = (formData.steps_to_reproduce || '')
    .split(/\\n|\n/)
    .filter(s => s.trim())
    .map(s => s.replace(/^\d+\.\s*/, '').trim())
    .filter(s => s);
  while (steps.length < 5) steps.push('');

  // Parser le plan d'action
  const actions = (formData.action_plan || '')
    .split(/\\n|\n/)
    .filter(s => s.trim())
    .map(s => s.replace(/^\d+\.\s*/, '').trim())
    .filter(s => s);
  while (actions.length < 3) actions.push('');

  const children = [
    // ── EN-TÊTE BALKIRA ──────────────────────────────────────
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2000, 5360, 2000],
      rows: [
        new TableRow({ children: [
          new TableCell({
            borders, rowSpan: 3,
            shading: { fill: C.navy, type: ShadingType.CLEAR },
            margins: { top: 160, bottom: 160, left: 200, right: 200 },
            verticalAlign: VerticalAlign.CENTER,
            width: { size: 2000, type: WidthType.DXA },
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER,
                children: [txt('BALKIRA', { bold: true, size: 32, color: C.orange })] }),
              new Paragraph({ alignment: AlignmentType.CENTER,
                children: [txt('Engineering Tomorrow', { size: 16, color: C.white, italic: true })] }),
            ]
          }),
          new TableCell({
            borders, shading: { fill: C.navy, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 60, left: 200, right: 200 },
            verticalAlign: VerticalAlign.CENTER, width: { size: 5360, type: WidthType.DXA },
            children: [new Paragraph({ alignment: AlignmentType.CENTER,
              children: [txt('ANOMALY FORM', { bold: true, size: 28, color: C.white })] })]
          }),
          new TableCell({
            borders, shading: { fill: C.navyMid, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            verticalAlign: VerticalAlign.CENTER, width: { size: 2000, type: WidthType.DXA },
            children: [
              new Paragraph({ children: [txt('Référence : ', { size: 16, color: C.grayText }),
                txt(formData.veeva_ref || 'BLK-ANO-XXXX', { size: 16, color: C.white, bold: true })] }),
              new Paragraph({ children: [txt('Version : ', { size: 16, color: C.grayText }),
                txt('V1', { size: 16, color: C.orange, bold: true })] }),
              new Paragraph({ children: [txt('Date : ', { size: 16, color: C.grayText }),
                txt(date, { size: 16, color: C.white })] }),
            ]
          }),
        ]}),
        new TableRow({ children: [
          new TableCell({
            borders, shading: { fill: C.orangeLight, type: ShadingType.CLEAR },
            margins: { top: 60, bottom: 60, left: 140, right: 140 },
            width: { size: 5360, type: WidthType.DXA },
            children: [new Paragraph({ alignment: AlignmentType.CENTER,
              children: [txt('⚠  BROUILLON — À réviser avant approbation Veeva Vault',
                { size: 16, color: C.amberText, bold: true })] })]
          }),
          new TableCell({
            borders, shading: { fill: C.navyMid, type: ShadingType.CLEAR },
            margins: { top: 60, bottom: 60, left: 120, right: 120 },
            width: { size: 2000, type: WidthType.DXA },
            children: [new Paragraph({ children: [txt('Statut : ', { size: 16, color: C.grayText }),
              txt('DRAFT', { size: 16, color: C.amber, bold: true })] })]
          }),
        ]}),
        new TableRow({ children: [
          new TableCell({
            borders, columnSpan: 2,
            shading: { fill: C.gray, type: ShadingType.CLEAR },
            margins: { top: 60, bottom: 60, left: 140, right: 140 },
            children: [new Paragraph({ children: [
              txt('Auteur : ', { size: 16, bold: true, color: C.navyMid }),
              txt('<À compléter avant approbation>   ', { size: 16, color: C.grayText, italic: true }),
              txt('Approbateur : ', { size: 16, bold: true, color: C.navyMid }),
              txt('<À compléter>   ', { size: 16, color: C.grayText, italic: true }),
              txt('Signatures : ', { size: 16, bold: true, color: C.navyMid }),
              txt('Via Veeva Vault', { size: 16, color: C.grayText, italic: true }),
            ]})]
          }),
        ]}),
      ]
    }),
    spacer(),

    // ── HISTORIQUE DES RÉVISIONS ─────────────────────────────
    h2('Historique des révisions'),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [1200, 1800, 4560, 1800],
      rows: [
        new TableRow({ children: [hCell('Version', 1200), hCell('Date', 1800),
          hCell('Description des modifications', 4560), hCell('Auteur', 1800)] }),
        new TableRow({ children: [
          new TableCell({ borders, width: { size: 1200, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 140, right: 140 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [txt('1.0', { size: 18, bold: true })] })] }),
          new TableCell({ borders, width: { size: 1800, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 140, right: 140 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [txt(date, { size: 18 })] })] }),
          new TableCell({ borders, width: { size: 4560, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 140, right: 140 },
            children: [new Paragraph({ children: [txt('Création initiale du document (V1)', { size: 18 })] })] }),
          new TableCell({ borders, width: { size: 1800, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 140, right: 140 },
            children: [new Paragraph({ children: [txt('<À compléter>', { size: 18, italic: true, color: C.grayText })] })] }),
        ]}),
        new TableRow({ children: [
          new TableCell({ borders, width: { size: 1200, type: WidthType.DXA }, children: [new Paragraph({ children: [txt('')] })] }),
          new TableCell({ borders, width: { size: 1800, type: WidthType.DXA }, children: [new Paragraph({ children: [txt('')] })] }),
          new TableCell({ borders, width: { size: 4560, type: WidthType.DXA }, children: [new Paragraph({ children: [txt('')] })] }),
          new TableCell({ borders, width: { size: 1800, type: WidthType.DXA }, children: [new Paragraph({ children: [txt('')] })] }),
        ]}),
      ]
    }),
    spacer(),
    divider(),

    // ── SECTION 1 — IDENTIFICATION ───────────────────────────
    h1('1. Identification de l\'anomalie'),
    noteBox('Ce tableau doit être complété dès la détection de l\'anomalie. La référence Veeva est générée lors de la création dans Veeva Vault.', 'info'),
    spacer(),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [3120, 6240],
      rows: [
        new TableRow({ children: [labelCell('Anomaly ID (Veeva)', 3120), valueCell(formData.veeva_ref || '<À compléter>', 6240)] }),
        new TableRow({ children: [labelCell('Application / Composant', 3120), valueCell(formData.application_component || '<À compléter>', 6240)] }),
        new TableRow({ children: [labelCell('Version du composant', 3120), valueCell(formData.component_version || '<À compléter>', 6240)] }),
        new TableRow({ children: [labelCell('Référence Test Case', 3120), valueCell(formData.test_case_ref || '<À compléter>', 6240)] }),
        new TableRow({ children: [
          labelCell('Priorité', 3120),
          new TableCell({
            borders, width: { size: 6240, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 140, right: 140 },
            children: [new Paragraph({ children: checkboxLine([
              { label: 'Critical', value: 'Critical' },
              { label: 'High', value: 'High' },
              { label: 'Medium', value: 'Medium' },
              { label: 'Low', value: 'Low' }
            ], priorityNorm) })]
          })
        ]}),
        new TableRow({ children: [
          labelCell('Statut', 3120),
          new TableCell({
            borders, width: { size: 6240, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 140, right: 140 },
            children: [new Paragraph({ children: checkboxLine([
              { label: 'Open', value: 'Open' },
              { label: 'In Progress', value: 'In Progress' },
              { label: 'Closed', value: 'Closed' }
            ], 'Open') })]  // V1 = toujours Open
          })
        ]}),
        new TableRow({ children: [labelCell('Date de détection', 3120), valueCell(date, 6240)] }),
        new TableRow({ children: [labelCell('Détecté par', 3120), valueCell(formData.assigned_to || '<À compléter>', 6240)] }),
      ]
    }),
    spacer(),

    // ── SECTION 2 — TITRE ────────────────────────────────────
    h1('2. Titre de l\'anomalie'),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
      rows: [new TableRow({ children: [new TableCell({
        borders, shading: { fill: C.gray, type: ShadingType.CLEAR },
        margins: { top: 140, bottom: 140, left: 200, right: 200 },
        children: [new Paragraph({ children: [txt(formData.title || '<À compléter>', { size: 20, bold: true })] })]
      })]})]
    }),
    spacer(),

    // ── SECTION 3 — DESCRIPTION ──────────────────────────────
    h1('3. Description de l\'anomalie'),
    noteBox('Décrivez précisément l\'écart entre le résultat obtenu et le résultat attendu. Un auditeur FDA/EMA doit pouvoir comprendre sans contexte.', 'info'),
    spacer(),

    h2('3.1  Résultat attendu'),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
      rows: [new TableRow({ children: [new TableCell({
        borders, margins: { top: 140, bottom: 140, left: 200, right: 200 },
        children: textToParagraphs(formData.expected_result)
      })]})]
    }),
    spacer(),

    h2('3.2  Résultat obtenu'),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
      rows: [new TableRow({ children: [new TableCell({
        borders, margins: { top: 140, bottom: 140, left: 200, right: 200 },
        children: textToParagraphs(formData.obtained_result)
      })]})]
    }),
    spacer(),

    h2('3.3  Description détaillée'),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
      rows: [new TableRow({ children: [new TableCell({
        borders, margins: { top: 140, bottom: 140, left: 200, right: 200 },
        children: textToParagraphs(formData.description)
      })]})]
    }),
    spacer(),

    // ── SECTION 4 — ÉTAPES ───────────────────────────────────
    h1('4. Étapes de reproduction'),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [800, 8560],
      rows: [
        new TableRow({ children: [hCell('Étape', 800), hCell('Action / Observation', 8560)] }),
        ...steps.slice(0, 5).map((step, i) => new TableRow({ children: [
          new TableCell({ borders, width: { size: 800, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 140, right: 140 },
            verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({ alignment: AlignmentType.CENTER,
              children: [txt(String(i + 1), { bold: true, size: 18 })] })] }),
          new TableCell({ borders, width: { size: 8560, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 140, right: 140 },
            children: [new Paragraph({ children: [txt(step || '<À compléter>', {
              size: 18, italic: !step, color: step ? C.black : C.grayText })] })] }),
        ]})),
      ]
    }),
    spacer(),

    // ── SECTION 5 — ANALYSE CAUSES ───────────────────────────
    h1('5. Analyse des causes'),
    h2('5.1  Hypothèse cause racine'),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
      rows: [new TableRow({ children: [new TableCell({
        borders, margins: { top: 140, bottom: 140, left: 200, right: 200 },
        children: textToParagraphs(formData.root_cause_hypothesis ||
          '<Première hypothèse sur la cause racine — à valider lors de l\'investigation>')
      })]})]
    }),
    spacer(),

    // ── SECTION 6 — PLAN D'ACTION V1 ─────────────────────────
    h1('6. Plan d\'action — Version 1'),
    noteBox('La V1 décrit le plan d\'action initial. La V2 (après correctif) documentera les actions réellement effectuées et la preuve de résolution.', 'warn'),
    spacer(),

    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [5360, 2200, 1800],
      rows: [
        new TableRow({ children: [hCell('Action corrective', 5360), hCell('Responsable', 2200), hCell('Date cible', 1800)] }),
        ...actions.slice(0, 3).map((action) => new TableRow({ children: [
          new TableCell({ borders, width: { size: 5360, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 140, right: 140 },
            children: [new Paragraph({ children: [txt(action || '<À compléter>',
              { size: 18, italic: !action, color: action ? C.black : C.grayText })] })] }),
          new TableCell({ borders, width: { size: 2200, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 140, right: 140 },
            children: [new Paragraph({ children: [txt(formData.assigned_to || '<À compléter>',
              { size: 18 })] })] }),
          new TableCell({ borders, width: { size: 1800, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 140, right: 140 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER,
              children: [txt(formData.target_date || '<JJ/MM/AAAA>', { size: 18 })] })] }),
        ]})),
      ]
    }),
    spacer(),
    spacer(),
    divider(),

    // ── SECTION 7 — CLÔTURE V2 (grisée en V1) ───────────────
    h1('7. Clôture — Version 2'),
    noteBox('⚠ Cette section sera complétée lors du passage en Version 2, après déploiement du correctif et re-test du cas de test.', 'v2'),
    spacer(),

    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [3120, 6240],
      rows: [
        new TableRow({ children: [
          labelCell('Résultat test case après correction', 3120),
          new TableCell({
            borders, width: { size: 6240, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 140, right: 140 },
            shading: { fill: 'F0F0F0', type: ShadingType.CLEAR },
            children: [new Paragraph({ children: [
              txt('☐ PASS', { size: 18, color: C.grayText }),
              txt('    ☐ PASS with observation', { size: 18, color: C.grayText }),
              txt('    ☐ FAIL', { size: 18, color: C.grayText }),
            ] })]
          })
        ]}),
        new TableRow({ children: [
          labelCell('Nouvelle anomalie créée (si FAIL)', 3120),
          valueCell('<À compléter en V2>', 6240, { italic: true, color: C.grayText })
        ]}),
        new TableRow({ children: [
          labelCell('Spécification(s) mise(s) à jour', 3120),
          valueCell('<À compléter en V2>', 6240, { italic: true, color: C.grayText })
        ]}),
        new TableRow({ children: [
          labelCell('Correctif déployé en', 3120),
          new TableCell({
            borders, width: { size: 6240, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 140, right: 140 },
            shading: { fill: 'F0F0F0', type: ShadingType.CLEAR },
            children: [new Paragraph({ children: [
              txt('☐ Production', { size: 18, color: C.grayText }),
              txt('    ☐ Validation', { size: 18, color: C.grayText }),
              txt('    ☐ Dev', { size: 18, color: C.grayText }),
              txt('    ☐ Non déployé', { size: 18, color: C.grayText }),
            ] })]
          })
        ]}),
        new TableRow({ children: [
          labelCell('Justification de clôture', 3120),
          valueCell('<À compléter en V2>', 6240, { italic: true, color: C.grayText })
        ]}),
        new TableRow({ children: [
          labelCell('Date de clôture', 3120),
          valueCell('<À compléter en V2>', 6240, { italic: true, color: C.grayText })
        ]}),
      ]
    }),
    spacer(),
    spacer(),

    // ── PIED DE PAGE ─────────────────────────────────────────
    new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 2, color: C.grayBorder } },
      alignment: AlignmentType.CENTER, spacing: { before: 400 },
      children: [txt(
        `Généré par Balkira GxP Doc Companion — ${requestId} — ${new Date().toISOString()} — Données de démonstration`,
        { size: 16, color: C.grayText, italic: true }
      )]
    }),
  ];

  const doc = new Document({
    styles: {
      default: { document: { run: { font: 'Arial', size: 20 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 24, bold: true, font: 'Arial', color: C.navy },
          paragraph: { spacing: { before: 280, after: 100 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 21, bold: true, font: 'Arial', color: C.navyMid },
          paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 1 } },
      ]
    },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 },
        margin: { top: 900, right: 1080, bottom: 900, left: 1080 } } },
      children
    }]
  });

  return Packer.toBuffer(doc);
}

// ── GÉNÉRATION ANOMALY FORM V2 ────────────────────────────────
async function generateAnomalyFormV2(formData, requestId, language) {
  const date = new Date().toLocaleDateString('fr-FR');

  const children = [
    // En-tête identique mais Version V2
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2000, 5360, 2000],
      rows: [
        new TableRow({ children: [
          new TableCell({
            borders, rowSpan: 3,
            shading: { fill: C.navy, type: ShadingType.CLEAR },
            margins: { top: 160, bottom: 160, left: 200, right: 200 },
            verticalAlign: VerticalAlign.CENTER,
            width: { size: 2000, type: WidthType.DXA },
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER,
                children: [txt('BALKIRA', { bold: true, size: 32, color: C.orange })] }),
              new Paragraph({ alignment: AlignmentType.CENTER,
                children: [txt('Engineering Tomorrow', { size: 16, color: C.white, italic: true })] }),
            ]
          }),
          new TableCell({
            borders, shading: { fill: C.navy, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 60, left: 200, right: 200 },
            verticalAlign: VerticalAlign.CENTER, width: { size: 5360, type: WidthType.DXA },
            children: [new Paragraph({ alignment: AlignmentType.CENTER,
              children: [txt('ANOMALY FORM — CLÔTURE', { bold: true, size: 26, color: C.white })] })]
          }),
          new TableCell({
            borders, shading: { fill: C.navyMid, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            verticalAlign: VerticalAlign.CENTER, width: { size: 2000, type: WidthType.DXA },
            children: [
              new Paragraph({ children: [txt('Référence : ', { size: 16, color: C.grayText }),
                txt(formData.veeva_ref || 'BLK-ANO-XXXX', { size: 16, color: C.white, bold: true })] }),
              new Paragraph({ children: [txt('Version : ', { size: 16, color: C.grayText }),
                txt('V2', { size: 16, color: C.orange, bold: true })] }),
              new Paragraph({ children: [txt('Date : ', { size: 16, color: C.grayText }),
                txt(date, { size: 16, color: C.white })] }),
            ]
          }),
        ]}),
        new TableRow({ children: [
          new TableCell({
            borders, shading: { fill: C.orangeLight, type: ShadingType.CLEAR },
            margins: { top: 60, bottom: 60, left: 140, right: 140 },
            width: { size: 5360, type: WidthType.DXA },
            children: [new Paragraph({ alignment: AlignmentType.CENTER,
              children: [txt('⚠  BROUILLON — À réviser avant approbation Veeva Vault',
                { size: 16, color: C.amberText, bold: true })] })]
          }),
          new TableCell({
            borders, shading: { fill: C.navyMid, type: ShadingType.CLEAR },
            margins: { top: 60, bottom: 60, left: 120, right: 120 },
            width: { size: 2000, type: WidthType.DXA },
            children: [new Paragraph({ children: [txt('Statut : ', { size: 16, color: C.grayText }),
              txt('DRAFT', { size: 16, color: C.amber, bold: true })] })]
          }),
        ]}),
        new TableRow({ children: [
          new TableCell({
            borders, columnSpan: 2,
            shading: { fill: C.gray, type: ShadingType.CLEAR },
            margins: { top: 60, bottom: 60, left: 140, right: 140 },
            children: [new Paragraph({ children: [
              txt('Auteur : ', { size: 16, bold: true, color: C.navyMid }),
              txt('<À compléter>   ', { size: 16, color: C.grayText, italic: true }),
              txt('Approbateur : ', { size: 16, bold: true, color: C.navyMid }),
              txt('<À compléter>   ', { size: 16, color: C.grayText, italic: true }),
              txt('Signatures : ', { size: 16, bold: true, color: C.navyMid }),
              txt('Via Veeva Vault', { size: 16, color: C.grayText, italic: true }),
            ]})]
          }),
        ]}),
      ]
    }),
    spacer(),

    noteBox(`Ce document est la Version 2 de l'anomalie ${formData.veeva_ref || 'BLK-ANO-XXXX'}. Il documente la clôture après déploiement du correctif.`, 'warn'),
    spacer(),

    // Résumé V1
    h1('1. Rappel de l\'anomalie (V1)'),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
      rows: [new TableRow({ children: [new TableCell({
        borders, margins: { top: 140, bottom: 140, left: 200, right: 200 },
        children: textToParagraphs(formData.v1_summary)
      })]})]
    }),
    spacer(),

    // Actions correctives réalisées
    h1('2. Actions correctives réalisées'),
    new Table({
      width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
      rows: [new TableRow({ children: [new TableCell({
        borders, margins: { top: 140, bottom: 140, left: 200, right: 200 },
        children: textToParagraphs(formData.corrective_actions_done)
      })]})]
    }),
    spacer(),

    // Résultat test case
    h1('3. Résultat du test case après correction'),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [3120, 6240],
      rows: [
        new TableRow({ children: [
          labelCell('Résultat test case', 3120),
          new TableCell({
            borders, width: { size: 6240, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 140, right: 140 },
            children: [new Paragraph({ children: checkboxLine([
              { label: 'PASS', value: 'PASS' },
              { label: 'PASS with observation', value: 'observation' },
              { label: 'FAIL', value: 'FAIL' }
            ], formData.test_case_result || '') })]
          })
        ]}),
        new TableRow({ children: [
          labelCell('Nouvelle anomalie créée (si FAIL)', 3120),
          valueCell(formData.new_anomaly_ref || 'N/A', 6240)
        ]}),
        new TableRow({ children: [
          labelCell('Spécification(s) mise(s) à jour', 3120),
          valueCell(formData.spec_ref_updated || (formData.spec_updated === 'Non applicable' ? 'Non applicable' : '<À compléter>'), 6240)
        ]}),
        new TableRow({ children: [
          labelCell('Correctif déployé en', 3120),
          new TableCell({
            borders, width: { size: 6240, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 140, right: 140 },
            children: [new Paragraph({ children: checkboxLine([
              { label: 'Production', value: 'Production' },
              { label: 'Validation', value: 'Validation' },
              { label: 'Dev', value: 'Dev' },
              { label: 'Non déployé', value: 'Non déployé' }
            ], formData.deployment_env || '') })]
          })
        ]}),
        new TableRow({ children: [
          labelCell('Justification de clôture', 3120),
          new TableCell({
            borders, width: { size: 6240, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 140, right: 140 },
            children: textToParagraphs(formData.closure_justification)
          })
        ]}),
        new TableRow({ children: [
          labelCell('Date de clôture', 3120),
          valueCell(date, 6240)
        ]}),
      ]
    }),
    spacer(),
    spacer(),

    new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 2, color: C.grayBorder } },
      alignment: AlignmentType.CENTER, spacing: { before: 400 },
      children: [txt(
        `Généré par Balkira GxP Doc Companion — ${requestId} — ${new Date().toISOString()} — Données de démonstration`,
        { size: 16, color: C.grayText, italic: true }
      )]
    }),
  ];

  const doc = new Document({
    styles: { default: { document: { run: { font: 'Arial', size: 20 } } } },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 },
        margin: { top: 900, right: 1080, bottom: 900, left: 1080 } } },
      children
    }]
  });
  return Packer.toBuffer(doc);
}

// ── ROUTE PRINCIPALE ──────────────────────────────────────────
app.post('/generate-docx', async (req, res) => {
  try {
    const { content, filename, template_id, version, request_id, language, form_data } = req.body;
    
    // Récupérer form_data soit directement soit parsé depuis content
    let formData = form_data || {};
    
    // Si pas de form_data mais un content markdown, extraire les données
    if (!form_data && content) {
      // Parser basique du markdown pour extraire les valeurs
      const lines = content.split('\n');
      for (const line of lines) {
        const match = line.match(/^\*\*(.+?)\*\*\s*:\s*(.+)$/);
        if (match) {
          const key = match[1].toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');
          formData[key] = match[2].trim();
        }
      }
      // Fallback: utiliser le contenu brut
      formData._raw_content = content;
    }

    let buffer;
    const ver = (version || 'V1').toUpperCase();
    
    // Router vers le bon générateur
    if (template_id === 'anomaly_form' && ver === 'V1') {
      buffer = await generateAnomalyFormV1(formData, request_id || 'REQ-UNKNOWN', language || 'fr');
    } else if (template_id === 'anomaly_form' && ver === 'V2') {
      buffer = await generateAnomalyFormV2(formData, request_id || 'REQ-UNKNOWN', language || 'fr');
    } else {
      // Fallback générique pour les autres templates
      buffer = await generateGenericDoc(content || '', filename || 'document.docx', 
        template_id, ver, request_id, language);
    }

    res.json({
      base64: buffer.toString('base64'),
      filename: filename || `BALKIRA_${(template_id||'DOC').toUpperCase()}_${ver}_${request_id}.docx`,
      size_kb: Math.round(buffer.length / 1024)
    });
  } catch (err) {
    console.error('Error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ── FALLBACK GÉNÉRIQUE ────────────────────────────────────────
async function generateGenericDoc(content, filename, templateId, version, requestId, language) {
  const lines = (content || '').split('\n');
  const children = [];
  
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 240 },
    children: [txt(`${templateId?.toUpperCase() || 'DOCUMENT'} — ${version}`, 
      { bold: true, size: 30, color: '0D1B2A' })] }));
  children.push(new Paragraph({ children: [txt(
    `Référence: ${requestId}  |  Date: ${new Date().toLocaleDateString('fr-FR')}  |  BROUILLON`,
    { size: 18, color: '666666' })] }));
  children.push(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'F4821F' } },
    spacing: { before: 120, after: 240 }, children: [txt('')] }));

  for (const line of lines) {
    const t = line.trim();
    if (!t) { children.push(new Paragraph({ children: [txt('')] })); continue; }
    if (t.startsWith('# ')) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 100 },
        children: [txt(t.replace(/^#+\s*/, ''), { bold: true, size: 24, color: '0D1B2A' })] }));
    } else if (t.startsWith('## ')) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 180, after: 80 },
        children: [txt(t.replace(/^##\s*/, ''), { bold: true, size: 21, color: '1A2B3C' })] }));
    } else if (t.startsWith('### ')) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 140, after: 60 },
        children: [txt(t.replace(/^###\s*/, ''), { bold: true, size: 20 })] }));
    } else if (t.startsWith('- ') || t.startsWith('* ')) {
      children.push(new Paragraph({ indent: { left: 360 },
        children: [txt('• ' + t.replace(/^[-*]\s*/, ''), { size: 20 })] }));
    } else {
      children.push(new Paragraph({ children: [txt(t, { size: 20 })] }));
    }
  }

  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 480 },
    border: { top: { style: BorderStyle.SINGLE, size: 2, color: 'CCCCCC' } },
    children: [txt(`Balkira GxP Doc Companion — ${requestId} — Démonstration`, 
      { size: 16, color: '888888', italic: true })] }));

  const doc = new Document({ sections: [{ properties: { page: { size: { width: 11906, height: 16838 },
    margin: { top: 900, right: 1080, bottom: 900, left: 1080 } } }, children }] });
  return Packer.toBuffer(doc);
}

app.get('/health', (req, res) => res.json({ status: 'ok', service: 'balkira-docx-service' }));

app.listen(PORT, () => console.log(`Balkira DOCX Service running on port ${PORT}`));