'use strict';

const express = require('express');
const cors = require('cors');
const {
  Document, Packer, Paragraph, TextRun,
  HeadingLevel, AlignmentType, BorderStyle,
  Table, TableRow, TableCell, WidthType,
  ShadingType, VerticalAlign, Footer,
  convertInchesToTwip
} = require('docx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

// ─── Palette Balkira ──────────────────────────────────────────────────────────
const NAVY       = '0D1B2A';
const NAVY_LIGHT = '1A2E45';
const ORANGE     = 'F4821F';
const WHITE      = 'FFFFFF';
const AMBER_BG   = 'FFF8E1';
const AMBER_TEXT = '856404';
const GRAY       = '888888';

// 2.5 cm en twips
const MARGIN = convertInchesToTwip(0.984);

// Labels des templates GxP
const TEMPLATE_LABELS = {
  anomaly_form:    "Formulaire d'Anomalie",
  interface_spec:  'Spécification d\'Interface',
  urs:             'User Requirements Specification',
  dira_pdfm:       'DIRA / PDFM',
  sop:             'Standard Operating Procedure',
  iq_oq_pq:        'IQ / OQ / PQ',
  uat_design:      'UAT Design',
  capa:            'CAPA',
  change_control:  'Change Control'
};

// ─── Helpers texte inline ─────────────────────────────────────────────────────

/**
 * Parse le markdown inline (**bold**) et retourne un tableau de TextRun.
 */
function parseInline(text, { color, size = 20, italics = false } = {}) {
  const runs = [];
  const parts = text.split(/(\*\*[^*]+\*\*)/);
  for (const part of parts) {
    if (!part) continue;
    if (part.startsWith('**') && part.endsWith('**')) {
      runs.push(new TextRun({
        text: part.slice(2, -2),
        bold: true, font: 'Arial', size,
        ...(color ? { color } : {}),
        ...(italics ? { italics } : {})
      }));
    } else {
      runs.push(new TextRun({
        text: part, font: 'Arial', size,
        ...(color ? { color } : {}),
        ...(italics ? { italics } : {})
      }));
    }
  }
  return runs;
}

// ─── En-tête 3 colonnes ───────────────────────────────────────────────────────

function buildHeader(templateLabel, version, requestId, language) {
  const date = new Date().toLocaleDateString(language === 'fr' ? 'fr-FR' : 'en-GB');

  const noBorder = { style: BorderStyle.NONE, size: 0, color: 'auto' };
  const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideH: noBorder, insideV: noBorder };

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorders,
    rows: [new TableRow({
      height: { value: 1200 },
      children: [
        // ── Colonne gauche : Balkira branding ─────────────────────────────
        new TableCell({
          width: { size: 28, type: WidthType.PERCENTAGE },
          shading: { type: ShadingType.SOLID, color: NAVY, fill: NAVY },
          verticalAlign: VerticalAlign.CENTER,
          borders: noBorders,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 100, after: 40 },
              children: [new TextRun({ text: 'BALKIRA', bold: true, size: 32, color: ORANGE, font: 'Arial' })]
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 100 },
              children: [new TextRun({ text: 'Engineering Tomorrow', size: 16, color: WHITE, font: 'Arial', italics: true })]
            })
          ]
        }),

        // ── Colonne centre : nom du template ──────────────────────────────
        new TableCell({
          width: { size: 44, type: WidthType.PERCENTAGE },
          shading: { type: ShadingType.SOLID, color: NAVY, fill: NAVY },
          verticalAlign: VerticalAlign.CENTER,
          borders: noBorders,
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({
              text: templateLabel,
              bold: true, size: 26, color: WHITE, font: 'Arial'
            })]
          })]
        }),

        // ── Colonne droite : références ───────────────────────────────────
        new TableCell({
          width: { size: 28, type: WidthType.PERCENTAGE },
          shading: { type: ShadingType.SOLID, color: NAVY_LIGHT, fill: NAVY_LIGHT },
          verticalAlign: VerticalAlign.CENTER,
          borders: noBorders,
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              spacing: { before: 80, after: 20 },
              children: [new TextRun({ text: `Réf : ${requestId}`, size: 16, color: ORANGE, font: 'Arial', bold: true })]
            }),
            new Paragraph({
              alignment: AlignmentType.RIGHT, spacing: { before: 0, after: 20 },
              children: [new TextRun({ text: `Version : ${version}`, size: 16, color: WHITE, font: 'Arial' })]
            }),
            new Paragraph({
              alignment: AlignmentType.RIGHT, spacing: { before: 0, after: 20 },
              children: [new TextRun({ text: `Date : ${date}`, size: 16, color: WHITE, font: 'Arial' })]
            }),
            new Paragraph({
              alignment: AlignmentType.RIGHT, spacing: { before: 0, after: 80 },
              children: [new TextRun({ text: 'Statut : DRAFT', size: 16, color: ORANGE, font: 'Arial', bold: true })]
            })
          ]
        })
      ]
    })]
  });
}

// ─── Parser de tableau Markdown ───────────────────────────────────────────────

function buildMarkdownTable(lines) {
  // Filtre les lignes séparateurs (|---|---|)
  const dataRows = lines.filter(l => !l.match(/^\|[\s|:-]+\|$/));
  if (dataRows.length === 0) return null;

  const cellBorder = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
  const cellBorders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder, insideH: cellBorder, insideV: cellBorder };

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    margins: { top: 80, bottom: 80, left: 80, right: 80 },
    rows: dataRows.map((line, rowIdx) => {
      const cells = line.split('|').slice(1, -1).map(c => c.trim());
      const isHeader = rowIdx === 0;
      return new TableRow({
        tableHeader: isHeader,
        children: cells.map(cell => new TableCell({
          shading: isHeader
            ? { type: ShadingType.SOLID, color: NAVY, fill: NAVY }
            : undefined,
          borders: cellBorders,
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({
              text: cell,
              bold: isHeader,
              color: isHeader ? WHITE : undefined,
              size: 18,
              font: 'Arial'
            })]
          })]
        }))
      });
    })
  });
}

// ─── Parser Markdown principal ────────────────────────────────────────────────

function parseMarkdown(content, requestId, templateLabel, version, language) {
  const elements = [];

  // En-tête
  elements.push(buildHeader(templateLabel, version, requestId, language));
  elements.push(new Paragraph({ children: [new TextRun('')], spacing: { before: 240, after: 0 } }));

  const lines = content.split('\n');
  let i = 0;

  while (i < lines.length) {
    const raw  = lines[i];
    const line = raw.trim();

    // ── Ligne vide ──────────────────────────────────────────────────────────
    if (!line) {
      elements.push(new Paragraph({ children: [new TextRun('')] }));
      i++; continue;
    }

    // ── Tableau Markdown ────────────────────────────────────────────────────
    if (line.startsWith('|')) {
      const tableLines = [];
      while (i < lines.length && lines[i].trim().startsWith('|')) {
        tableLines.push(lines[i].trim());
        i++;
      }
      const table = buildMarkdownTable(tableLines);
      if (table) {
        elements.push(table);
        elements.push(new Paragraph({ children: [new TextRun('')] }));
      }
      continue;
    }

    // ── Ligne horizontale ───────────────────────────────────────────────────
    if (/^[-*_]{3,}$/.test(line)) {
      elements.push(new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ORANGE } },
        children: [new TextRun('')],
        spacing: { before: 120, after: 120 }
      }));
      i++; continue;
    }

    // ── Heading 1 ───────────────────────────────────────────────────────────
    if (/^# /.test(line)) {
      elements.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 360, after: 160 },
        children: [new TextRun({
          text: line.replace(/^#\s+/, ''),
          bold: true, size: 28, color: NAVY, font: 'Arial'
        })]
      }));
      i++; continue;
    }

    // ── Heading 2 ───────────────────────────────────────────────────────────
    if (/^## /.test(line)) {
      elements.push(new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 240, after: 120 },
        children: [new TextRun({
          text: line.replace(/^##\s+/, ''),
          bold: true, size: 24, color: NAVY, font: 'Arial'
        })]
      }));
      i++; continue;
    }

    // ── Heading 3 ───────────────────────────────────────────────────────────
    if (/^### /.test(line)) {
      elements.push(new Paragraph({
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 180, after: 80 },
        children: [new TextRun({
          text: line.replace(/^###\s+/, ''),
          bold: true, size: 22, font: 'Arial'
        })]
      }));
      i++; continue;
    }

    // ── Bullet list ─────────────────────────────────────────────────────────
    if (/^[-*]\s/.test(line)) {
      elements.push(new Paragraph({
        indent: { left: 360 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: '• ', size: 20, font: 'Arial', color: ORANGE }),
          ...parseInline(line.replace(/^[-*]\s+/, ''))
        ]
      }));
      i++; continue;
    }

    // ── Numbered list ───────────────────────────────────────────────────────
    const numMatch = line.match(/^(\d+)\.\s+(.*)$/);
    if (numMatch) {
      elements.push(new Paragraph({
        indent: { left: 360 },
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: `${numMatch[1]}. `, bold: true, size: 20, font: 'Arial', color: NAVY }),
          ...parseInline(numMatch[2])
        ]
      }));
      i++; continue;
    }

    // ── Avertissement ⚠ / À VÉRIFIER ───────────────────────────────────────
    if (line.includes('⚠') || line.toUpperCase().includes('À VÉRIFIER') || line.toUpperCase().includes('TO VERIFY')) {
      elements.push(new Paragraph({
        indent: { left: 180, right: 180 },
        spacing: { before: 100, after: 100 },
        border: {
          left: { style: BorderStyle.THICK, size: 8, color: ORANGE }
        },
        shading: { type: ShadingType.SOLID, color: AMBER_BG, fill: AMBER_BG },
        children: [new TextRun({
          text: line,
          size: 20, font: 'Arial', color: AMBER_TEXT, bold: true
        })]
      }));
      i++; continue;
    }

    // ── Paragraphe normal (supporte inline bold) ────────────────────────────
    elements.push(new Paragraph({
      spacing: { before: 60, after: 60 },
      children: parseInline(line)
    }));
    i++;
  }

  return elements;
}

// ─── Endpoint principal ───────────────────────────────────────────────────────

app.post('/generate-docx', async (req, res) => {
  try {
    const {
      content    = '',
      filename,
      template_id = 'doc',
      version    = 'V1',
      request_id,
      language   = 'fr'
    } = req.body;

    if (!content) {
      return res.status(400).json({ error: 'Le champ "content" est requis.' });
    }

    const requestId      = request_id || `REQ-${Date.now()}`;
    const templateLabel  = TEMPLATE_LABELS[template_id] || template_id.replace(/_/g, ' ').toUpperCase();
    const safeFilename   = filename || `BALKIRA_${template_id.toUpperCase()}_${version}_${requestId}.docx`;
    const footerText     = `Généré par Balkira GxP Doc Companion — Mistral AI · EU · RGPD — ${requestId} — Données de démonstration`;

    const bodyElements = parseMarkdown(content, requestId, templateLabel, version, language);

    const doc = new Document({
      styles: {
        default: {
          document: {
            run: { font: 'Arial', size: 20 }
          }
        }
      },
      sections: [{
        properties: {
          page: {
            size:   { width: 11906, height: 16838 },  // A4 en twips
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
          }
        },
        footers: {
          default: new Footer({
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({
                text: footerText,
                size: 16, color: GRAY, font: 'Arial', italics: true
              })]
            })]
          })
        },
        children: bodyElements
      }]
    });

    const buffer  = await Packer.toBuffer(doc);
    const base64  = buffer.toString('base64');
    const size_kb = Math.round(buffer.length / 1024);

    console.log(`[OK] ${safeFilename} — ${size_kb} KB`);
    res.json({ base64, filename: safeFilename, size_kb });

  } catch (err) {
    console.error('[ERR]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── Healthcheck ──────────────────────────────────────────────────────────────
app.get('/health', (_req, res) => res.json({ status: 'ok', service: 'balkira-docx-service' }));

// ─── Start ────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Balkira DocX Service — port ${PORT}`));
