'use strict';
const express = require('express');
const cors = require('cors');
const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));
const PORT = process.env.PORT || 3000;

// ── DATE HELPERS ─────────────────────────────────────────────
function todayFR() {
  return new Date().toLocaleDateString('fr-FR', { day: '2-digit', month: '2-digit', year: 'numeric' });
}
function todayFRTime() {
  const n = new Date();
  return n.toLocaleDateString('fr-FR', { day: '2-digit', month: '2-digit', year: 'numeric' })
    + '  ' + n.toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });
}

// ── XML HELPERS ──────────────────────────────────────────────
// Encode plain text value for safe insertion into XML
function xe(str) {
  return String(str == null ? '' : str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// Replace first occurrence of a raw XML pattern with an encoded plain-text value
function rf(xml, search, value) {
  const idx = xml.indexOf(search);
  if (idx === -1) return xml;
  return xml.substring(0, idx) + xe(value) + xml.substring(idx + search.length);
}

// Replace ALL occurrences of a raw XML pattern with an encoded plain-text value
function ra(xml, search, value) {
  if (!xml.includes(search)) return xml;
  return xml.split(search).join(xe(value));
}

// Replace first occurrence without encoding (for already-encoded replacements)
function rfRaw(xml, search, replacement) {
  const idx = xml.indexOf(search);
  if (idx === -1) return xml;
  return xml.substring(0, idx) + replacement + xml.substring(idx + search.length);
}

// Replace ALL without encoding
function raRaw(xml, search, replacement) {
  if (!xml.includes(search)) return xml;
  return xml.split(search).join(replacement);
}

// Checkbox: mark selected option(s) with ☑, keep others as ☐
// Uses global replace — only suitable when each option appears exactly once in the doc
function checkboxes(xml, options, selected) {
  if (!selected) return xml;
  const sel = String(selected).toLowerCase();
  for (const opt of options) {
    const mark = sel.includes(opt.toLowerCase()) ? '☑ ' : '☐ ';
    xml = xml.split('☐ ' + opt).join(mark + opt);
  }
  return xml;
}

// OUI/NON binary pair: replaces the FIRST occurrence of "☐ OUI  ☐ NON" only.
// Use this for table rows where the same ☐ OUI  ☐ NON pattern repeats per row.
// value: true/"OUI"/"oui" → ☑ OUI  ☐ NON
//        false/"NON"/"non" or anything else → ☐ OUI  ☑ NON
function checkboxOuiNon(xml, value) {
  const isOui = value === true || String(value).toUpperCase().trim() === 'OUI';
  const pattern = '☐ OUI  ☐ NON';
  const replacement = isOui ? '☑ OUI  ☐ NON' : '☐ OUI  ☑ NON';
  const idx = xml.indexOf(pattern);
  if (idx === -1) return xml;
  return xml.substring(0, idx) + replacement + xml.substring(idx + pattern.length);
}

// ── MISTRAL CONTENT PARSER ───────────────────────────────────
function parseMistral(content) {
  if (!content) return {};
  if (typeof content === 'object') return content;
  try {
    const cleaned = String(content)
      .replace(/```json\n?/g, '')
      .replace(/\n?```/g, '')
      .trim();
    return JSON.parse(cleaned);
  } catch (_) {
    return {};
  }
}

// Safe array/object accessor with fallback
function safeGet(obj, ...keys) {
  let cur = obj;
  for (const k of keys) {
    if (cur == null || typeof cur !== 'object') return '';
    cur = Array.isArray(cur) ? cur[k] : cur[k];
  }
  return cur == null ? '' : cur;
}

// ── ANOMALY FORM MAPPING ──────────────────────────────────────
function applyAnomalyForm(xml, fd, m, requestId) {
  const today = todayFR();

  // ── Header reference & date ─────────────────────────────
  const ref = fd.veeva_ref || ('BLK-ANO-' + requestId);
  xml = rfRaw(xml, 'Référence : BLK-ANO-XXXX', 'Référence : ' + xe(ref));
  xml = rfRaw(xml, 'Date : JJ/MM/AAAA', 'Date : ' + today);

  // ── Identification section ──────────────────────────────
  xml = ra(xml, 'BALKIRA-ANO-&lt;XXXX&gt;  [Généré par Veeva Vault]', ref);
  xml = ra(xml, '&lt;Ex: PAS-X MES, Veeva Vault, iLearn LMS&gt;', fd.application_component || '');
  xml = ra(xml, '&lt;Ex: v3.2.1&gt;', fd.component_version || '');
  xml = ra(xml, '&lt;Ex: TC-UAT-045 / Run 3 / Step 7&gt;', fd.test_case_ref || '');

  // Priority & Status checkboxes
  xml = checkboxes(xml, ['Critical', 'High', 'Medium', 'Low'], fd.priority || '');
  xml = checkboxes(xml, ['Open', 'In Progress', 'Closed'], fd.status || 'Open');

  // Date + detected by
  xml = rf(xml, '&lt;JJ/MM/AAAA  HH:MM&gt;', fd.detection_date || todayFRTime());
  xml = rf(xml, '&lt;Nom Prénom — Fonction&gt;', fd.detected_by || fd.assigned_to || '');

  // ── Revision history ────────────────────────────────────
  xml = rf(xml, '&lt;JJ/MM/AAAA&gt;', today);        // revision date (1st occurrence)
  xml = rf(xml, '&lt;Nom Prénom&gt;', fd.assigned_to || '');  // revision author

  // ── Auteur / Approbateur in header ───────────────────────
  xml = rf(xml, '&lt;À compléter&gt;', fd.assigned_to || '');  // Auteur
  xml = rf(xml, '&lt;À compléter&gt;', fd.approver || fd.assigned_to || '');  // Approbateur

  // ── Titre ───────────────────────────────────────────────
  xml = ra(xml, "&lt;Donner un titre court et descriptif à l'anomalie&gt;", fd.title || '');

  // ── Expected / Obtained results (section 3) ─────────────
  xml = ra(xml, '&lt;Décrivez ce qui aurait dû se passer selon la spécification de référence&gt;',
    m.expected_result_gxp || fd.expected_result || '');
  xml = ra(xml, "&lt;Décrivez précisément ce qui s'est réellement passé&gt;",
    m.obtained_result_gxp || fd.obtained_result || '');

  // ── Steps to reproduce (section 4) — 5 rows ─────────────
  const steps = Array.isArray(m.steps_enriched) ? m.steps_enriched : [];
  const rawSteps = fd.steps_to_reproduce ? String(fd.steps_to_reproduce).split(/\n/).filter(Boolean) : [];
  for (let i = 0; i < 5; i++) {
    const val = safeGet(steps, i, 'action') || rawSteps[i] || fd['step_' + (i + 1)] || '';
    xml = rf(xml, '&lt;À compléter&gt;', val);
  }

  // ── Root cause (section 5) ───────────────────────────────
  xml = ra(xml, "&lt;Première hypothèse sur la cause racine — à valider lors de l'investigation&gt;",
    m.root_cause_analysis || fd.root_cause_hypothesis || '');

  // ── Action plan (section 6) — 3 rows ────────────────────
  const actions = Array.isArray(m.actions_plan) ? m.actions_plan : [];
  for (let i = 0; i < 3; i++) {
    const a = actions[i] || {};
    xml = rf(xml, "&lt;Décrivez l'action à mener&gt;", a.action || (i === 0 ? fd.action_plan || '' : ''));
    xml = rf(xml, '&lt;Nom / Équipe&gt;', a.responsable || fd.assigned_to || '');
    xml = rf(xml, '&lt;JJ/MM/AAAA&gt;', a.delai || fd.target_date || today);
  }

  // ── Closure section V2 (left mostly empty) ──────────────
  xml = ra(xml, '&lt;Référence Veeva nouvelle anomalie si applicable&gt;', '');
  xml = ra(xml, '&lt;Référence(s) doc mis à jour — ex: BLK-IFS-001 v2.0&gt;', '');
  xml = ra(xml, "&lt;Pourquoi cette anomalie peut être considérée comme clôturée&gt;",
    m.closure_criteria || '');

  // ── Remaining date placeholders ──────────────────────────
  xml = ra(xml, '&lt;JJ/MM/AAAA&gt;', today);

  // ── Footer ──────────────────────────────────────────────
  xml = raRaw(xml,
    'Données de démonstration — Ne pas utiliser en production',
    xe('Réf: ' + requestId + ' — ' + today));

  return xml;
}

// ── URS MAPPING ───────────────────────────────────────────────
function applyURS(xml, fd, m, requestId) {
  const today = todayFR();
  const ref = fd.doc_ref || ('BLK-URS-' + requestId);

  xml = rfRaw(xml, 'Référence : BLK-URS-XXXX', 'Référence : ' + xe(ref));
  xml = rfRaw(xml, 'Date : JJ/MM/AAAA', 'Date : ' + today);

  // System info
  xml = ra(xml, '&lt;Nom complet du système&gt;', fd.system_name || '');
  xml = ra(xml, '&lt;Nom du site de déploiement&gt;', fd.site_name || '');
  xml = ra(xml, '&lt;Description du besoin métier — pourquoi ce système est nécessaire&gt;',
    m.system_description || fd.description || '');
  xml = ra(xml, '&lt;Équipes impliquées — Production, Qualité, IT, etc.&gt;', fd.stakeholders || '');

  // Revision history
  xml = rf(xml, '&lt;JJ/MM/AAAA&gt;', today);
  xml = rf(xml, '&lt;Nom Prénom&gt;', fd.assigned_to || '');

  // Auteur / Approbateur
  xml = rf(xml, '&lt;À compléter&gt;', fd.assigned_to || '');
  xml = rf(xml, '&lt;À compléter&gt;', fd.approver || fd.assigned_to || '');

  // Document references (ID01-ID03) — leave as N/A
  xml = rf(xml, '&lt;À compléter&gt;', 'N/A');
  xml = rf(xml, '&lt;À compléter&gt;', 'N/A');
  xml = rf(xml, '&lt;À compléter&gt;', 'N/A');

  // Glossary custom rows
  xml = rf(xml, '&lt;À compléter&gt;', '');
  xml = rf(xml, '&lt;À compléter&gt;', '');

  // Requirements custom rows
  const reqs = Array.isArray(m.requirements) ? m.requirements : [];
  const customReqIds = reqs.slice(4).map(r => r.id);
  xml = rf(xml, '&lt;ID&gt;', customReqIds[0] || 'REQ-EXT-001');
  xml = rf(xml, '&lt;Décrivez l\'exigence de façon précise et vérifiable&gt;',
    safeGet(reqs, 4, 'description') || fd.additional_requirement || '');
  xml = rf(xml, '&lt;M/D/O&gt;', safeGet(reqs, 4, 'priority') || 'D');
  xml = rf(xml, '&lt;GxP/HSE/OP&gt;', safeGet(reqs, 4, 'classification') || 'OP');

  // Data requirements custom row
  xml = rf(xml, '&lt;ID&gt;', customReqIds[1] || 'REQ-EXT-002');
  xml = rf(xml, '&lt;Décrivez l\'exigence&gt;', '');
  xml = rf(xml, '&lt;M/D/O&gt;', 'D');
  xml = rf(xml, '&lt;GxP/HSE/OP&gt;', 'OP');

  // User requirements custom row
  xml = rf(xml, '&lt;ID&gt;', '');
  xml = rf(xml, '&lt;Décrivez l\'exigence&gt;', '');
  xml = rf(xml, '&lt;M/D/O&gt;', '');
  xml = rf(xml, '&lt;GxP/HSE/OP&gt;', '');

  // Data sources rows
  xml = rf(xml, '&lt;Nom équipement ou système&gt;', fd.source_system || '');
  xml = rf(xml, '&lt;OPC-UA / ODBC / Fichier plat / API&gt;', fd.connection_type || 'OPC-UA');
  xml = rf(xml, '&lt;Remarques&gt;', '');
  xml = rf(xml, '&lt;Nom équipement ou système&gt;', '');
  xml = rf(xml, '&lt;OPC-UA / ODBC / Fichier plat / API&gt;', '');
  xml = rf(xml, '&lt;Remarques&gt;', '');
  xml = rf(xml, '&lt;Nom équipement ou système&gt;', '');
  xml = rf(xml, '&lt;OPC-UA / ODBC / Fichier plat / API&gt;', '');
  xml = rf(xml, '&lt;Remarques&gt;', '');

  // Interfaces custom row
  const ifaces = Array.isArray(m.interfaces) ? m.interfaces : [];
  xml = rf(xml, '&lt;ID&gt;', 'INT-003');
  xml = rf(xml, '&lt;Système cible&gt;', safeGet(ifaces, 2, 'system') || '');
  xml = rf(xml, '&lt;Description de l\'interface&gt;', safeGet(ifaces, 2, 'description') || '');

  // Records custom row
  const recs = Array.isArray(m.records) ? m.records : [];
  xml = rf(xml, '&lt;ID&gt;', safeGet(recs, 1, 'id') || 'REC-002');
  xml = rf(xml, '&lt;Description de l\'enregistrement&gt;', safeGet(recs, 1, 'description') || '');
  xml = rf(xml, '&lt;Processus&gt;', safeGet(recs, 1, 'process') || '');

  // Checkbox OUI/NON for GxP records
  xml = checkboxes(xml, ['OUI', 'NON'], 'OUI');

  // Clean up any remaining placeholders
  xml = ra(xml, '&lt;À compléter&gt;', '');
  xml = ra(xml, '&lt;ID&gt;', '');
  xml = ra(xml, '&lt;JJ/MM/AAAA&gt;', today);

  xml = raRaw(xml,
    'Données de démonstration — Ne pas utiliser en production',
    xe('Réf: ' + requestId + ' — ' + today));

  return xml;
}

// ── INTERFACE SPEC MAPPING ────────────────────────────────────
function applyInterfaceSpec(xml, fd, m, requestId) {
  const today = todayFR();
  const ref = fd.doc_ref || ('BLK-IFS-' + requestId);

  xml = rfRaw(xml, 'Référence : BLK-IFS-XXXX', 'Référence : ' + xe(ref));
  xml = rfRaw(xml, 'Date : JJ/MM/AAAA', 'Date : ' + today);

  // Equipment description
  xml = ra(xml, "&lt;Domaine de l'équipement — ex: Zone de production, Utilities&gt;",
    fd.domain || fd.description || '');
  xml = ra(xml, '&lt;Classe générique de l\'équipement&gt;', fd.equipment_class || '');
  xml = ra(xml, '&lt;Description courte de l\'équipement et de son rôle&gt;',
    fd.description || m.objective || '');

  // Revision history
  xml = rf(xml, '&lt;JJ/MM/AAAA&gt;', today);
  xml = rf(xml, '&lt;Nom Prénom&gt;', fd.assigned_to || '');

  // Auteur / Approbateur
  xml = rf(xml, '&lt;À compléter&gt;', fd.assigned_to || '');
  xml = rf(xml, '&lt;À compléter&gt;', fd.approver || fd.assigned_to || '');

  // Document references
  xml = rf(xml, '&lt;À compléter&gt;', 'N/A');
  xml = rf(xml, '&lt;À compléter&gt;', 'N/A');
  xml = rf(xml, '&lt;À compléter&gt;', 'N/A');

  // Systems list
  const systems = Array.isArray(m.systems) ? m.systems : [];
  xml = rf(xml, '&lt;Identifiant système&gt;', safeGet(systems, 0, 'id') || 'SYS-001');
  xml = rf(xml, '&lt;Nom du device ou serveur&gt;', safeGet(systems, 0, 'name') || fd.source_system || '');
  xml = rf(xml, '&lt;x.x.x.x&gt;', safeGet(systems, 0, 'ip') || '');
  xml = rf(xml, '&lt;Description&gt;', safeGet(systems, 0, 'description') || '');
  xml = rf(xml, '&lt;Identifiant système&gt;', safeGet(systems, 1, 'id') || '');
  xml = rf(xml, '&lt;Nom du device ou serveur&gt;', safeGet(systems, 1, 'name') || '');
  xml = rf(xml, '&lt;x.x.x.x&gt;', '');
  xml = rf(xml, '&lt;Description&gt;', '');

  // OPC-UA config
  xml = ra(xml, '&lt;Nom ou adresse du serveur OPC-UA&gt;', m.opc_server || fd.opc_server || '');
  xml = ra(xml, '&lt;À préciser selon politique de sécurité du site&gt;',
    m.auth_method || 'Username-Password');
  xml = ra(xml, '&lt;Politique de validation des certificats&gt;', 'Trust on first use (TOFU)');

  // Tag rows
  const tags = Array.isArray(m.tags) ? m.tags : [];
  for (let i = 0; i < 3; i++) {
    xml = rf(xml, '&lt;Nom tag cible&gt;', safeGet(tags, i, 'name') || ('TAG_00' + (i + 1)));
    xml = rf(xml, '&lt;Chemin OPC source&gt;', safeGet(tags, i, 'opc_path') || '');
    xml = rf(xml, '&lt;Unité&gt;', safeGet(tags, i, 'unit') || '');
    xml = rf(xml, '&lt;x sec&gt;', safeGet(tags, i, 'frequency') || '10s');
  }

  // DeltaV / ODBC section
  xml = ra(xml, '&lt;Nom ou adresse du serveur DeltaV&gt;', '');
  xml = ra(xml, '&lt;Nom de la base de données&gt;', '');
  xml = ra(xml, '&lt;ODBC / API / Fichier plat&gt;', '');
  xml = rf(xml, '&lt;À compléter&gt;', fd.db_type || '');
  xml = rf(xml, '&lt;À compléter&gt;', fd.db_server || '');
  xml = rf(xml, '&lt;À compléter&gt;', fd.db_name || '');
  xml = rf(xml, '&lt;À compléter&gt;', fd.db_driver || '');
  xml = ra(xml, '&lt;UTC / Local — préciser&gt;', 'UTC');
  xml = ra(xml, '&lt;À préciser&gt;', '');

  // Alarm mapping rows
  const events = Array.isArray(m.event_mapping) ? m.event_mapping : [];
  xml = rf(xml, '&lt;Type alarme — paramètre critique&gt;', fd.alarm_critical || safeGet(events, 0, 'source') || '');
  xml = rf(xml, '&lt;Type alarme — paramètre non critique&gt;', fd.alarm_noncritical || safeGet(events, 1, 'source') || '');
  xml = rf(xml, '&lt;Type alarme maintenance&gt;', fd.alarm_maintenance || safeGet(events, 2, 'source') || '');
  xml = rf(xml, '&lt;À compléter&gt;', '');
  xml = rf(xml, '&lt;À compléter&gt;', '');
  xml = rf(xml, '&lt;À compléter&gt;', '');

  // Event structure
  xml = ra(xml, '&lt;Convention de nommage des équipements dans les événements&gt;', fd.naming_convention || '');
  xml = ra(xml, '&lt;Champs obligatoires dans l\'en-tête de chaque événement&gt;', '');
  xml = ra(xml, "&lt;Intervalle d'exécution de la requête — ex: toutes les 30 secondes&gt;",
    fd.sync_frequency || '30 secondes');

  // Audit Trail
  const atSelected = m.audit_trail !== false ? 'OUI — GxP critique' : 'Non applicable';
  xml = checkboxes(xml, ['OUI — GxP critique', 'OUI — Non GxP', 'Non applicable'], atSelected);
  xml = ra(xml, '&lt;Base de données / Fichier plat / API&gt;', m.audit_trail_source || '');
  xml = ra(xml, '&lt;Méthode technique d\'extraction&gt;', '');
  xml = ra(xml, '&lt;Temps réel / Périodique — préciser&gt;', fd.sync_frequency || 'Temps réel');
  xml = ra(xml, '&lt;Décrire comment l\'Audit Trail est stocké&gt;', '');

  // MES interface
  const mesRequired = m.mes_required === true || String(fd.mes_required || '').toLowerCase() === 'oui';
  xml = checkboxes(xml, ['OUI', 'Non applicable'], mesRequired ? 'OUI' : 'Non applicable');
  xml = ra(xml, '&lt;Nom et version du système MES&gt;', safeGet(m, 'mes_config', 'system') || fd.mes_system || '');
  xml = ra(xml, '&lt;Liste des tags / paramètres échangés avec le MES&gt;', '');
  xml = ra(xml, '&lt;Structure de données attendue par le MES — unités, alias, calculs&gt;', '');
  xml = ra(xml, '&lt;Détailler les calculs configurés pour l\'interface&gt;', '');

  // Cleanup
  xml = ra(xml, '&lt;À compléter&gt;', '');
  xml = ra(xml, '&lt;JJ/MM/AAAA&gt;', today);

  xml = raRaw(xml,
    'Données de démonstration — Ne pas utiliser en production',
    xe('Réf: ' + requestId + ' — ' + today));

  return xml;
}

// ── DIRA / PDFM MAPPING ───────────────────────────────────────
function applyDiraPDFM(xml, fd, m, requestId) {
  const today = todayFR();
  const ref = fd.doc_ref || ('BLK-DIRA-' + requestId);

  xml = rfRaw(xml, 'Référence : BLK-DIRA-XXXX', 'Référence : ' + xe(ref));
  xml = rfRaw(xml, 'Date : JJ/MM/AAAA', 'Date : ' + today);

  // Scope & context
  xml = ra(xml, "&lt;Nom du processus métier faisant l'objet de cette analyse&gt;",
    fd.process_name || fd.system_name || '');
  xml = ra(xml, '&lt;Nom du site&gt;', fd.site_name || '');
  xml = ra(xml, '&lt;Référence de la SOP DIRA utilisée&gt;', m.sop_ref || fd.sop_ref || 'SOP-DIRA-001');
  xml = rf(xml, '&lt;JJ/MM/AAAA&gt;', today);  // analysis date
  xml = ra(xml, "&lt;Listez les sous-processus et activités inclus dans l'analyse PDFM/DIRA&gt;",
    m.scope_in || fd.scope_in || '');
  xml = ra(xml, '&lt;Listez explicitement ce qui est exclu de l\'analyse et justifiez chaque exclusion&gt;',
    m.scope_out || fd.scope_out || '');

  // Revision history
  xml = rf(xml, '&lt;JJ/MM/AAAA&gt;', today);
  xml = rf(xml, '&lt;Nom Prénom&gt;', fd.assigned_to || '');

  // Auteur / Approbateur
  xml = rf(xml, '&lt;À compléter&gt;', fd.assigned_to || '');
  xml = rf(xml, '&lt;À compléter&gt;', fd.approver || fd.assigned_to || '');

  // Document references
  xml = rf(xml, '&lt;À compléter&gt;', 'N/A');
  xml = rf(xml, '&lt;À compléter&gt;', 'N/A');
  xml = rf(xml, '&lt;À compléter&gt;', 'N/A');

  // Systems criticality
  const systems = Array.isArray(m.systems) ? m.systems : [];
  xml = rf(xml, '&lt;Système 1&gt;', safeGet(systems, 0, 'name') || fd.system_name || '');
  xml = rf(xml, '&lt;Processus principal supporté&gt;', safeGet(systems, 0, 'process') || fd.process_name || '');
  xml = rf(xml, '&lt;Système 2&gt;', safeGet(systems, 1, 'name') || '');
  xml = rf(xml, '&lt;Processus secondaire supporté&gt;', safeGet(systems, 1, 'process') || '');

  // Extra system row (criticality checkbox)
  xml = rf(xml, '&lt;À compléter&gt;', safeGet(systems, 2, 'name') || '');
  const sys2crit = safeGet(systems, 2, 'criticality') || 'Major';
  xml = checkboxes(xml, ['Critical', 'Major'], sys2crit);
  xml = rf(xml, '&lt;À compléter&gt;', safeGet(systems, 2, 'process') || '');

  // Actors — use checkboxOuiNon (replaces FIRST occurrence only, one cell at a time)
  // Template layout per actor row: [PDFM cell] [DIRA cell] [DIRA mitigé cell] [Nom cell]
  // Defaults from template: System Owner OUI/OUI/OUI, Process Owner OUI/OUI/OUI,
  //                         Quality Manager NON/OUI/OUI, IT Lead OUI/OUI/OUI
  const actors = Array.isArray(m.actors) ? m.actors : [];
  const defaultActors = [
    { pdfm: true,  dira: true,  dira_mitigated: true,  name: fd.system_owner || fd.assigned_to || '' },
    { pdfm: true,  dira: true,  dira_mitigated: true,  name: fd.process_owner || fd.assigned_to || '' },
    { pdfm: false, dira: true,  dira_mitigated: true,  name: fd.quality_manager || '' },
    { pdfm: true,  dira: true,  dira_mitigated: true,  name: fd.it_lead || fd.assigned_to || '' }
  ];
  const actorData = defaultActors.map((def, i) => ({
    pdfm: actors[i] ? actors[i].pdfm : def.pdfm,
    dira: actors[i] ? actors[i].dira : def.dira,
    dira_mitigated: actors[i] ? actors[i].dira_mitigated : def.dira_mitigated,
    name: actors[i] ? (actors[i].name || def.name) : def.name
  }));
  for (const actor of actorData) {
    xml = checkboxOuiNon(xml, actor.pdfm);
    xml = checkboxOuiNon(xml, actor.dira);
    xml = checkboxOuiNon(xml, actor.dira_mitigated);
    xml = rf(xml, '&lt;Nom Prénom&gt;', actor.name);
  }

  // Extra actor row
  xml = rf(xml, '&lt;À compléter&gt;', '');
  xml = rf(xml, '&lt;&gt;', '');
  xml = rf(xml, '&lt;&gt;', '');
  xml = rf(xml, '&lt;&gt;', '');
  xml = rf(xml, '&lt;Nom Prénom&gt;', '');

  // PDFM data flow rows
  const risks = Array.isArray(m.risks) ? m.risks : [];
  for (let i = 0; i < 3; i++) {
    xml = rf(xml, '&lt;Nom sous-processus&gt;', safeGet(risks, i, 'subprocess') || '');
    xml = rf(xml, '&lt;Systèmes source et cible&gt;', '');
    xml = rf(xml, '&lt;Type et nature des données échangées&gt;', '');
    xml = rf(xml, '&lt;Contrôles ALCOA+ en place ou manquants&gt;', '');
  }

  // DIRA risk table — pre-filled risks rows use literal values, custom row is placeholders
  // Replace &lt;Sous-processus&gt; (appears in pre-filled risks 1-3)
  for (let i = 0; i < 4; i++) {
    xml = rf(xml, '&lt;Sous-processus&gt;', safeGet(risks, i, 'subprocess') || '');
  }

  // Custom risk row (risk #4)
  xml = rf(xml, '&lt;Décrivez le risque identifié&gt;', safeGet(risks, 3, 'description') || '');
  xml = rf(xml, '&lt;1-3&gt;', String(safeGet(risks, 3, 'severity') || '1'));
  xml = rf(xml, '&lt;1-3&gt;', String(safeGet(risks, 3, 'frequency') || '1'));
  xml = rf(xml, '&lt;1-3&gt;', String(safeGet(risks, 3, 'detectability') || '1'));
  xml = rf(xml, '&lt;RPN&gt;', String(safeGet(risks, 3, 'rpn') || '1'));
  xml = rf(xml, '&lt;Action de mitigation proposée&gt;', safeGet(risks, 3, 'mitigation') || '');

  // Action plan rows
  const actionsM = Array.isArray(m.actions) ? m.actions : [];
  for (let i = 0; i < 2; i++) {
    xml = rf(xml, '&lt;Détail de l\'action corrective à mener&gt;',
      safeGet(actionsM, i, 'action') || '');
    xml = rf(xml, '&lt;Nom / Équipe&gt;',
      safeGet(actionsM, i, 'owner') || fd.assigned_to || '');
    xml = rf(xml, '&lt;JJ/MM/AAAA&gt;',
      safeGet(actionsM, i, 'deadline') || fd.target_date || today);
    xml = checkboxes(xml, ['Ouvert', 'Clôturé'], 'Ouvert');
  }

  // Risk summary counts
  xml = rf(xml, '&lt;X&gt;', String(m.risks_before_l1 || '1'));
  xml = rf(xml, '&lt;X&gt;', String(m.risks_before_l2 || '3'));
  xml = rf(xml, '&lt;X&gt;', String(m.risks_before_l3 || '1'));
  xml = rf(xml, '&lt;X&gt;', String(m.risks_after_l1 || '4'));
  xml = rf(xml, '&lt;X&gt;', String(m.risks_after_l2 || '1'));
  xml = rf(xml, '&lt;X&gt;', String(m.risks_after_l3 || '0'));

  // Conclusion
  xml = ra(xml, '&lt;Rédigez la conclusion : tous les risques DI ont-ils été traités ? Tous les risques résiduels sont-ils à un niveau acceptable ? Quel est le statut final ?&gt;',
    m.conclusion || '');

  // Cleanup remaining
  xml = ra(xml, '&lt;À compléter&gt;', '');
  xml = ra(xml, '&lt;Nom Prénom&gt;', fd.assigned_to || '');
  xml = ra(xml, '&lt;JJ/MM/AAAA&gt;', today);
  xml = ra(xml, '&lt;&gt;', '');

  xml = raRaw(xml,
    'Données de démonstration — Ne pas utiliser en production',
    xe('Réf: ' + requestId + ' — ' + today));

  return xml;
}

// ── DISPATCH ─────────────────────────────────────────────────
function applyMapping(xml, templateId, form_data, mistral, requestId) {
  const fd = form_data || {};
  const m = mistral || {};
  const tid = String(templateId || '').toLowerCase().replace(/-/g, '_');

  // Normalize &apos; inside Word XML placeholders &lt;...&gt; so our string matching works
  // (some placeholders have apostrophes double-encoded as &apos; within already-encoded brackets)
  xml = xml.replace(/(&lt;[^<]{0,300}?)&apos;([^<]{0,300}&gt;)/g, "$1'$2");

  if (tid.startsWith('anomaly')) return applyAnomalyForm(xml, fd, m, requestId);
  if (tid.startsWith('urs')) return applyURS(xml, fd, m, requestId);
  if (tid.startsWith('interface')) return applyInterfaceSpec(xml, fd, m, requestId);
  if (tid.startsWith('dira')) return applyDiraPDFM(xml, fd, m, requestId);

  // Unknown template — just replace dates and footer
  xml = raRaw(xml, 'Date : JJ/MM/AAAA', 'Date : ' + todayFR());
  xml = ra(xml, '&lt;JJ/MM/AAAA&gt;', todayFR());
  return xml;
}

// ── ROUTES ───────────────────────────────────────────────────
app.get('/health', (req, res) => {
  const tplDir = path.join(__dirname, 'templates');
  const templates = fs.existsSync(tplDir) ? fs.readdirSync(tplDir).filter(f => f.endsWith('.docx')) : [];
  res.json({ status: 'ok', port: PORT, templates });
});

app.post('/generate-docx', async (req, res) => {
  try {
    const {
      content, filename, template_id, version,
      request_id, language, form_data
    } = req.body;

    if (!template_id) return res.status(400).json({ error: 'template_id requis' });

    // Build template filename: anomaly_form + V1 → anomaly_form_v1.docx
    const ver = String(version || 'V1').toLowerCase();
    const tplName = String(template_id).toLowerCase().replace(/-/g, '_') + '_' + ver + '.docx';
    const tplPath = path.join(__dirname, 'templates', tplName);

    if (!fs.existsSync(tplPath)) {
      return res.status(404).json({
        error: 'Template introuvable',
        tried: tplName,
        available: fs.readdirSync(path.join(__dirname, 'templates'))
      });
    }

    // Load template as ZIP
    const tplBuf = fs.readFileSync(tplPath);
    const zip = await JSZip.loadAsync(tplBuf);

    // Extract & patch word/document.xml
    let xml = await zip.file('word/document.xml').async('string');

    const mistral = parseMistral(content);
    const reqId = request_id || ('REQ-' + Date.now());

    xml = applyMapping(xml, template_id, form_data, mistral, reqId);

    // Write patched XML back into the ZIP
    zip.file('word/document.xml', xml);

    // Generate output buffer
    const outBuf = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });

    const outFilename = filename ||
      ('BALKIRA_' + String(template_id).toUpperCase() + '_' + String(version || 'V1').toUpperCase() +
       '_' + reqId + '.docx');

    res.json({
      base64: outBuf.toString('base64'),
      filename: outFilename,
      size_kb: Math.round(outBuf.length / 1024)
    });

  } catch (err) {
    console.error('[generate-docx]', err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Balkira DOCX Service — port ${PORT}`);
  console.log(`Templates: ${path.join(__dirname, 'templates')}`);
});
