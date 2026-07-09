/**
 * MAISON ABEILLE — KPI Dashboard Setup (v2)
 * Adapté à la version actuelle de l'enquête patients.
 *
 * Instructions :
 * 1. Extensions > Apps Script dans Google Sheets
 * 2. Coller ce code entier (remplacer l'ancien)
 * 3. Ctrl+S pour sauvegarder
 * 4. Cliquer sur ▶ Exécuter (fonction : createKPIDashboard)
 * 5. Autoriser les permissions si demandé
 */

// ── Génère une référence de colonne dynamique basée sur le nom du champ ──
// Cherche le nom dans la ligne 1 du sheet source → renvoie toute la colonne
function colRange(D, field) {
  var m = 'MATCH("' + field + '",\'' + D + '\'!1:1,0)';
  var letter = 'SUBSTITUTE(ADDRESS(1,' + m + ',4),"1","")';
  return 'INDIRECT("\'' + D + '\'!"&' + letter + '&"2:"&' + letter + ')';
}

// ── Formule : moyenne d'une colonne numérique (exclut 0 et vide) ──
function fAvg(D, field, decimals) {
  var d = decimals !== undefined ? decimals : 1;
  return '=IFERROR(ROUND(AVERAGEIF(' + colRange(D, field) + ',">0"),' + d + '),"–")';
}

// ── Formule : % de réponses correspondant à une valeur exacte ──
function fPct(D, field, value) {
  var range = colRange(D, field);
  var total = 'COUNTA(' + colRange(D, field) + ')';
  return '=IFERROR(TEXT(COUNTIF(' + range + ',"' + value + '")/' + total + ',"0%"),"–")';
}

// ── Formule : % contenant un texte (pour checkboxes multi-valeurs) ──
function fPctContains(D, field, value) {
  var range = colRange(D, field);
  var total = 'COUNTA(' + colRange(D, field) + ')';
  return '=IFERROR(TEXT(COUNTIF(' + range + ',"*' + value + '*")/' + total + ',"0%"),"–")';
}

// ── Formule : compte total de réponses ──
function fCount(D, field) {
  return '=IFERROR(COUNTA(' + colRange(D, field) + '),"–")';
}

// ── Formule : NPS promoteurs ──
function fNpsGroup(D, field, op, threshold) {
  var range = colRange(D, field);
  var total = 'COUNTA(' + colRange(D, field) + ')';
  return '=IFERROR(TEXT(COUNTIF(' + range + ',"' + op + threshold + '")/' + total + ',"0%"),"–")';
}


function createKPIDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Feuille source = premier onglet ──
  var dataSheet = ss.getSheets()[0];
  var D = dataSheet.getName();

  // ── Création ou réinitialisation de l'onglet KPI ──
  var kpiSheet = ss.getSheetByName('KPI Dashboard');
  if (kpiSheet) {
    kpiSheet.clear();
    kpiSheet.clearFormats();
    kpiSheet.clearConditionalFormatRules();
  } else {
    kpiSheet = ss.insertSheet('KPI Dashboard', 1);
  }

  // ── Palette ──
  var C = {
    dark:        '#192038',
    gold:        '#B8945A',
    goldLight:   '#F5ECD9',
    goldMid:     '#D4A96A',
    cream:       '#FAF6EF',
    bg:          '#F0EBE1',
    gray:        '#F6F3EF',
    grayDark:    '#888888',
    green:       '#1E8449',
    greenLight:  '#E9F7EF',
    orange:      '#9A6400',
    orangeLight: '#FEF5E4',
    red:         '#B03020',
    redLight:    '#FDEDEC',
    border:      '#DDD5C5',
    white:       '#FFFFFF',
    autoTag:     '#7A5C2E',
  };

  // ── Largeurs colonnes ──
  //    Cat   Nom KPI   Cible   Valeur   Tendance   Statut   Source
  [155, 300, 110, 130, 110, 155, 170].forEach(function(w, i) {
    kpiSheet.setColumnWidth(i + 1, w);
  });

  // ── Données KPI ──
  // type: 'sep' | 'auto' | 'manual'
  var kpis = [

    // ─────────────────────────────────────────────────
    { type:'sep', label:'🌟  Score NPS & Recommandation' },

    { type:'auto', cat:'NPS',
      name:'Score NPS moyen (0 – 10)',
      target:'≥ 8,0',
      formula: fAvg(D, 'q12_nps', 1),
      source:'Enquête · q12_nps' },

    { type:'auto', cat:'NPS',
      name:'% Promoteurs  (note 9 – 10)',
      target:'≥ 60 %',
      formula: fNpsGroup(D, 'q12_nps', '>=', 9),
      source:'Enquête · q12_nps' },

    { type:'auto', cat:'NPS',
      name:'% Passifs  (note 7 – 8)',
      target:'< 30 %',
      formula: '=IFERROR(TEXT((COUNTIF(' + colRange(D,"q12_nps") + ',">=7")-COUNTIF(' + colRange(D,"q12_nps") + ',">=9"))/COUNTA(' + colRange(D,"q12_nps") + '),"0%"),"–")',
      source:'Enquête · q12_nps' },

    { type:'auto', cat:'NPS',
      name:'% Détracteurs  (note ≤ 6)',
      target:'≤ 10 %',
      formula: fNpsGroup(D, 'q12_nps', '<=', 6),
      source:'Enquête · q12_nps' },

    { type:'auto', cat:'NPS',
      name:'Total réponses reçues',
      target:'≥ 50 / mois',
      formula: fCount(D, 'submitted_at'),
      source:'Enquête · toutes questions' },

    // ─────────────────────────────────────────────────
    { type:'sep', label:'📅  Prise de Rendez-vous  (note ★ / 5)' },

    { type:'auto', cat:'RDV',
      name:'Trouver le bon créneau  ★',
      target:'≥ 4,0 / 5',
      formula: fAvg(D, 'q3a', 1),
      source:'Enquête · q3a' },

    { type:'auto', cat:'RDV',
      name:'Comprendre quel motif choisir  ★',
      target:'≥ 4,0 / 5',
      formula: fAvg(D, 'q3b', 1),
      source:'Enquête · q3b' },

    { type:'auto', cat:'RDV',
      name:'Joindre le cabinet si besoin  ★',
      target:'≥ 4,0 / 5',
      formula: fAvg(D, 'q3c', 1),
      source:'Enquête · q3c' },

    { type:'auto', cat:'RDV',
      name:'Délai avant obtention du RDV  ★',
      target:'≥ 3,5 / 5',
      formula: fAvg(D, 'q3d', 1),
      source:'Enquête · q3d' },

    // ─────────────────────────────────────────────────
    { type:'sep', label:'⏱  Expérience en Cabinet' },

    { type:'auto', cat:'Cabinet',
      name:'Temps d\'attente < 15 min (%)',
      target:'≥ 70 %',
      formula: fPct(D, 'q8', 'moins15'),
      source:'Enquête · q8' },

    { type:'auto', cat:'Cabinet',
      name:'Patient informé du délai d\'attente (%)',
      target:'≥ 75 %',
      formula: fPct(D, 'q8b', 'Oui'),
      source:'Enquête · q8b' },

    { type:'auto', cat:'Cabinet',
      name:'Inquiétude adressée en consultation (%)',
      target:'≥ 70 %',
      formula: fPct(D, 'q7', 'Prise en charge'),
      source:'Enquête · q7' },

    { type:'auto', cat:'Cabinet',
      name:'Instructions post-consultation claires (%)',
      target:'≥ 80 %',
      formula: fPct(D, 'q9', 'Clair'),
      source:'Enquête · q9' },

    // ─────────────────────────────────────────────────
    { type:'sep', label:'📣  Communication & Transparence' },

    { type:'auto', cat:'Comm.',
      name:'Tarifs communiqués clairement (%)',
      target:'≥ 75 %',
      formula: fPct(D, 'q5', 'Oui clairement'),
      source:'Enquête · q5' },

    { type:'auto', cat:'Comm.',
      name:'Site web consulté et utile (%)',
      target:'≥ 50 %',
      formula: fPct(D, 'q4', 'Oui utile'),
      source:'Enquête · q4' },

    { type:'auto', cat:'Comm.',
      name:'Canal principal : WhatsApp (%)',
      target:'Info',
      formula: fPctContains(D, 'q6', 'WhatsApp'),
      source:'Enquête · q6' },

    { type:'auto', cat:'Comm.',
      name:'Canal principal : Téléphone (%)',
      target:'Info',
      formula: fPctContains(D, 'q6', 'Telephone'),
      source:'Enquête · q6' },

    // ─────────────────────────────────────────────────
    { type:'sep', label:'🔍  Découverte & Acquisition' },

    { type:'auto', cat:'Acq.',
      name:'Découverte via Google (%)',
      target:'Info',
      formula: fPctContains(D, 'q1', 'Google'),
      source:'Enquête · q1' },

    { type:'auto', cat:'Acq.',
      name:'Découverte via recommandation médecin (%)',
      target:'Info',
      formula: fPctContains(D, 'q1', 'Recommandation medecin'),
      source:'Enquête · q1' },

    { type:'auto', cat:'Acq.',
      name:'Découverte via bouche-à-oreille (%)',
      target:'Info',
      formula: fPctContains(D, 'q1', 'Bouche-a-oreille'),
      source:'Enquête · q1' },

    { type:'auto', cat:'Acq.',
      name:'Découverte via Instagram / réseaux (%)',
      target:'Info',
      formula: fPctContains(D, 'q1', 'Instagram'),
      source:'Enquête · q1' },

    // ─────────────────────────────────────────────────
    { type:'sep', label:'📅  Activité Cabinet  (saisie manuelle)' },

    { type:'manual', cat:'Activité', name:'Nouveaux patients / mois',         target:'≥ 40',       source:'Manuel' },
    { type:'manual', cat:'Activité', name:'Taux de remplissage agenda (%)',   target:'≥ 85 %',     source:'Manuel' },
    { type:'manual', cat:'Activité', name:'Taux d\'annulation RDV (%)',       target:'≤ 8 %',      source:'Manuel' },
    { type:'manual', cat:'Activité', name:'Délai moyen avant RDV (jours)',    target:'≤ 7 jours',  source:'Manuel' },
    { type:'manual', cat:'Activité', name:'Taux de fidélisation (%)',         target:'≥ 65 %',     source:'Manuel' },

    // ─────────────────────────────────────────────────
    { type:'sep', label:'⭐  Réputation en Ligne  (saisie manuelle)' },

    { type:'manual', cat:'Réputation', name:'Note Google (/ 5)',        target:'≥ 4,7', source:'Manuel · Google' },
    { type:'manual', cat:'Réputation', name:'Nombre d\'avis Google',   target:'+ 5 / mois', source:'Manuel · Google' },
    { type:'manual', cat:'Réputation', name:'Note Doctolib (/ 5)',      target:'≥ 4,8', source:'Manuel · Doctolib' },

    // ─────────────────────────────────────────────────
    { type:'sep', label:'💶  Performance Commerciale  (saisie manuelle)' },

    { type:'manual', cat:'Commercial', name:'CA mensuel (€)',                    target:'Budget défini', source:'Manuel · Compta' },
    { type:'manual', cat:'Commercial', name:'Panier moyen par patient (€)',      target:'À définir',     source:'Manuel' },
    { type:'manual', cat:'Commercial', name:'Taux de conversion 1ère visite (%)', target:'≥ 70 %',      source:'Manuel' },

  ];

  // ── Écriture des lignes ──
  var ROW = 1;

  // Titre
  var t = kpiSheet.getRange(ROW, 1, 1, 7);
  t.merge(); t.setValue('MAISON ABEILLE — Tableau de bord des KPI');
  t.setBackground(C.dark); t.setFontColor(C.gold);
  t.setFontWeight('bold'); t.setFontSize(15); t.setFontFamily('Arial');
  t.setHorizontalAlignment('center'); t.setVerticalAlignment('middle');
  kpiSheet.setRowHeight(ROW, 52); ROW++;

  // Sous-titre
  var sub = kpiSheet.getRange(ROW, 1, 1, 7);
  sub.merge();
  sub.setValue('Les lignes « Enquête » se mettent à jour automatiquement dès qu\'une réponse arrive · Compléter manuellement les lignes « Manuel »');
  sub.setBackground(C.bg); sub.setFontColor(C.grayDark);
  sub.setFontSize(8); sub.setFontFamily('Arial');
  sub.setHorizontalAlignment('center'); sub.setVerticalAlignment('middle');
  kpiSheet.setRowHeight(ROW, 26); ROW++;

  // En-têtes
  var headers = ['Catégorie', 'Indicateur (KPI)', 'Cible', 'Valeur actuelle', 'Tendance', 'Statut', 'Source'];
  var hR = kpiSheet.getRange(ROW, 1, 1, 7);
  hR.setValues([headers]);
  hR.setBackground(C.dark); hR.setFontColor('#FFFFFF');
  hR.setFontWeight('bold'); hR.setFontSize(9); hR.setFontFamily('Arial');
  hR.setHorizontalAlignment('center'); hR.setVerticalAlignment('middle');
  hR.setBorder(false, false, true, false, false, false, C.gold, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  kpiSheet.setRowHeight(ROW, 38); ROW++;

  var firstDataRow = ROW;
  var altIdx = 0;

  kpis.forEach(function(k) {
    if (k.type === 'sep') {
      // ── Séparateur ──
      kpiSheet.setRowHeight(ROW, 34);
      var sep = kpiSheet.getRange(ROW, 1, 1, 7);
      sep.merge(); sep.setValue(k.label);
      sep.setBackground(C.gold); sep.setFontColor('#FFFFFF');
      sep.setFontWeight('bold'); sep.setFontSize(9); sep.setFontFamily('Arial');
      sep.setVerticalAlignment('middle'); sep.setHorizontalAlignment('left');
      altIdx = 0;

    } else {
      // ── Ligne KPI ──
      kpiSheet.setRowHeight(ROW, 36);
      var isAuto = (k.type === 'auto');
      var bg = (altIdx % 2 === 0) ? C.white : C.gray;

      // Catégorie
      var c1 = kpiSheet.getRange(ROW, 1);
      c1.setValue(k.cat);
      c1.setBackground(bg); c1.setFontColor(C.grayDark);
      c1.setFontSize(8); c1.setFontFamily('Arial');
      c1.setVerticalAlignment('middle');

      // Nom KPI
      var c2 = kpiSheet.getRange(ROW, 2);
      c2.setValue(k.name);
      c2.setBackground(bg); c2.setFontColor(C.dark);
      c2.setFontSize(9); c2.setFontFamily('Arial');
      c2.setVerticalAlignment('middle');

      // Cible
      var c3 = kpiSheet.getRange(ROW, 3);
      c3.setValue(k.target);
      c3.setBackground(bg); c3.setFontColor(C.grayDark);
      c3.setFontSize(9); c3.setFontFamily('Arial');
      c3.setHorizontalAlignment('center'); c3.setVerticalAlignment('middle');

      // Valeur actuelle
      var c4 = kpiSheet.getRange(ROW, 4);
      if (isAuto && k.formula) {
        c4.setFormula(k.formula);
      }
      c4.setBackground(bg); c4.setFontColor(C.dark);
      c4.setFontSize(12); c4.setFontFamily('Arial');
      c4.setFontWeight('bold');
      c4.setHorizontalAlignment('center'); c4.setVerticalAlignment('middle');

      // Tendance
      var c5 = kpiSheet.getRange(ROW, 5);
      c5.setBackground(bg); c5.setFontSize(12);
      c5.setHorizontalAlignment('center'); c5.setVerticalAlignment('middle');

      // Statut
      var c6 = kpiSheet.getRange(ROW, 6);
      c6.setBackground(bg); c6.setFontSize(9); c6.setFontFamily('Arial');
      c6.setHorizontalAlignment('center'); c6.setVerticalAlignment('middle');

      // Source
      var c7 = kpiSheet.getRange(ROW, 7);
      c7.setValue(k.source);
      c7.setBackground(isAuto ? C.goldLight : bg);
      c7.setFontColor(isAuto ? C.autoTag : C.grayDark);
      c7.setFontSize(8); c7.setFontFamily('Arial');
      c7.setFontStyle('italic');
      c7.setHorizontalAlignment('center'); c7.setVerticalAlignment('middle');

      // Bordure inférieure
      kpiSheet.getRange(ROW, 1, 1, 7).setBorder(
        false, false, true, false, false, false,
        C.border, SpreadsheetApp.BorderStyle.SOLID
      );

      altIdx++;
    }
    ROW++;
  });

  var lastDataRow = ROW - 1;

  // ── Validation : Tendance (colonne 5) ──
  var tendRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['↑ Hausse', '→ Stable', '↓ Baisse', ''], true)
    .setAllowInvalid(false).build();
  kpiSheet.getRange(firstDataRow, 5, lastDataRow - firstDataRow + 1, 1)
    .setDataValidation(tendRule);

  // ── Validation : Statut (colonne 6) ──
  var statRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['✅ OK', '⚠️ À surveiller', '🔴 Action requise', ''], true)
    .setAllowInvalid(false).build();
  kpiSheet.getRange(firstDataRow, 6, lastDataRow - firstDataRow + 1, 1)
    .setDataValidation(statRule);

  // ── Mise en forme conditionnelle ──
  var statusRange = kpiSheet.getRange(firstDataRow, 6, lastDataRow - firstDataRow + 1, 1);
  var trendRange  = kpiSheet.getRange(firstDataRow, 5, lastDataRow - firstDataRow + 1, 1);
  var rules = [];

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('✅').setBackground(C.greenLight).setFontColor(C.green).setBold(true)
    .setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('⚠️').setBackground(C.orangeLight).setFontColor(C.orange).setBold(true)
    .setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('🔴').setBackground(C.redLight).setFontColor(C.red).setBold(true)
    .setRanges([statusRange]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('↑').setFontColor(C.green).setBold(true)
    .setRanges([trendRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('↓').setFontColor(C.red).setBold(true)
    .setRanges([trendRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('→').setFontColor(C.grayDark)
    .setRanges([trendRange]).build());

  kpiSheet.setConditionalFormatRules(rules);

  // ── Bordure extérieure ──
  kpiSheet.getRange(3, 1, lastDataRow - 2, 7).setBorder(
    true, true, true, true, false, false,
    C.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // ── Figer les 3 premières lignes ──
  kpiSheet.setFrozenRows(3);

  SpreadsheetApp.getUi().alert(
    '✅  KPI Dashboard mis à jour !\n\n' +
    'Onglets présents : ' + ss.getSheets().map(function(s){ return s.getName(); }).join(', ') + '\n\n' +
    '• Colonnes auto : calculées depuis "' + D + '"\n' +
    '• Colonnes manuelles : saisir Valeur + choisir Tendance & Statut\n' +
    '  via les listes déroulantes (↑ → ↓ et ✅ ⚠️ 🔴)'
  );
}
