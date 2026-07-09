/**
 * MAISON ABEILLE — KPI Dashboard Setup
 *
 * Instructions :
 * 1. Dans Google Sheets, ouvrir Extensions > Apps Script
 * 2. Coller ce code entier, remplacer le code existant
 * 3. Cliquer sur ▶ Exécuter (fonction : createKPIDashboard)
 * 4. Autoriser les permissions demandées
 * 5. L'onglet "KPI Dashboard" apparaît automatiquement
 *
 * Pour actualiser les formules : relancer la fonction à tout moment.
 */

function createKPIDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Détection de la feuille de réponses (premier onglet) ──
  var dataSheet = ss.getSheets()[0];
  var D = dataSheet.getName(); // nom dynamique pour les formules

  // ── Création ou réinitialisation de l'onglet KPI ──
  var kpiSheet = ss.getSheetByName('KPI Dashboard');
  if (kpiSheet) {
    kpiSheet.clear();
    kpiSheet.clearFormats();
    kpiSheet.clearConditionalFormatRules();
  } else {
    kpiSheet = ss.insertSheet('KPI Dashboard', 1);
  }

  // ── Palette couleurs ──
  var C = {
    dark:        '#192038',
    gold:        '#B8945A',
    goldLight:   '#F5ECD9',
    cream:       '#FAF6EF',
    gray:        '#F4F1EC',
    grayDark:    '#888888',
    green:       '#1E8449',
    greenLight:  '#EAFAF1',
    orange:      '#B7770D',
    orangeLight: '#FEF9E7',
    red:         '#C0392B',
    redLight:    '#FDEDEC',
    border:      '#DDD5C5',
    white:       '#FFFFFF',
  };

  // ── Largeurs des colonnes ──
  var colWidths = [155, 290, 110, 130, 110, 150, 175];
  for (var c = 0; c < colWidths.length; c++) {
    kpiSheet.setColumnWidth(c + 1, colWidths[c]);
  }

  // ── Structure des données ──
  // Chaque ligne : [catégorie, nom KPI, cible, formule/valeur, source, type]
  // type: 'sep' = séparateur de catégorie, 'auto' = formule, 'manual' = saisie manuelle
  var kpis = [

    { type: 'sep', label: '🌟  Satisfaction & Fidélisation' },

    { type: 'auto', cat: 'Satisfaction',
      name: 'Score NPS moyen (0 – 10)',
      target: '≥ 8,0',
      formula: '=IFERROR(ROUND(AVERAGEIF(\'' + D + '\'!S:S,"<>"),1),"–")',
      source: 'Enquête – auto' },

    { type: 'auto', cat: 'Satisfaction',
      name: '% Promoteurs  (note 9 – 10)',
      target: '≥ 60 %',
      formula: '=IFERROR(TEXT(COUNTIF(\'' + D + '\'!S:S,">=9")/COUNTA(\'' + D + '\'!S2:S),"0%"),"–")',
      source: 'Enquête – auto' },

    { type: 'auto', cat: 'Satisfaction',
      name: '% Détracteurs  (note ≤ 6)',
      target: '≤ 10 %',
      formula: '=IFERROR(TEXT(COUNTIF(\'' + D + '\'!S:S,"<=6")/COUNTA(\'' + D + '\'!S2:S),"0%"),"–")',
      source: 'Enquête – auto' },

    { type: 'auto', cat: 'Satisfaction',
      name: 'Total réponses enquête',
      target: '≥ 50 / mois',
      formula: '=IFERROR(COUNTA(\'' + D + '\'!A2:A),"–")',
      source: 'Enquête – auto' },

    { type: 'sep', label: '⏱  Parcours Patient' },

    { type: 'auto', cat: 'Parcours',
      name: 'Temps d\'attente < 15 min (%)',
      target: '≥ 70 %',
      formula: '=IFERROR(TEXT(COUNTIF(\'' + D + '\'!N:N,"moins15")/COUNTA(\'' + D + '\'!N2:N),"0%"),"–")',
      source: 'Enquête – auto' },

    { type: 'auto', cat: 'Parcours',
      name: 'Instructions post-consultation claires (%)',
      target: '≥ 80 %',
      formula: '=IFERROR(TEXT(COUNTIF(\'' + D + '\'!P:P,"Clair")/COUNTA(\'' + D + '\'!P2:P),"0%"),"–")',
      source: 'Enquête – auto' },

    { type: 'auto', cat: 'Parcours',
      name: 'Tarifs communiqués clairement (%)',
      target: '≥ 75 %',
      formula: '=IFERROR(TEXT(COUNTIF(\'' + D + '\'!J:J,"Oui clairement")/COUNTA(\'' + D + '\'!J2:J),"0%"),"–")',
      source: 'Enquête – auto' },

    { type: 'auto', cat: 'Parcours',
      name: 'Inquiétude adressée en consultation (%)',
      target: '≥ 70 %',
      formula: '=IFERROR(TEXT(COUNTIF(\'' + D + '\'!M:M,"Prise en charge")/COUNTA(\'' + D + '\'!M2:M),"0%"),"–")',
      source: 'Enquête – auto' },

    { type: 'sep', label: '📅  Activité Cabinet' },

    { type: 'manual', cat: 'Activité', name: 'Nouveaux patients / mois',           target: '≥ 40',      source: 'Manuel' },
    { type: 'manual', cat: 'Activité', name: 'Taux de remplissage agenda (%)',      target: '≥ 85 %',    source: 'Manuel' },
    { type: 'manual', cat: 'Activité', name: 'Taux d\'annulation RDV (%)',          target: '≤ 8 %',     source: 'Manuel' },
    { type: 'manual', cat: 'Activité', name: 'Délai moyen avant RDV (jours)',       target: '≤ 7 jours', source: 'Manuel' },
    { type: 'manual', cat: 'Activité', name: 'Taux de fidélisation patients (%)',   target: '≥ 65 %',    source: 'Manuel' },

    { type: 'sep', label: '⭐  Réputation en Ligne' },

    { type: 'manual', cat: 'Réputation', name: 'Note Google (/ 5)',         target: '≥ 4,7', source: 'Manuel – Google' },
    { type: 'manual', cat: 'Réputation', name: 'Nombre d\'avis Google',     target: '+ 5 / mois', source: 'Manuel – Google' },
    { type: 'manual', cat: 'Réputation', name: 'Note Doctolib (/ 5)',       target: '≥ 4,8', source: 'Manuel – Doctolib' },

    { type: 'sep', label: '💶  Performance Commerciale' },

    { type: 'manual', cat: 'Commercial', name: 'CA mensuel (€)',                   target: 'Budget défini', source: 'Manuel – Compta' },
    { type: 'manual', cat: 'Commercial', name: 'Panier moyen par patient (€)',     target: 'À définir',     source: 'Manuel' },
    { type: 'manual', cat: 'Commercial', name: 'Taux de conversion 1ère visite (%)', target: '≥ 70 %',    source: 'Manuel' },
  ];

  // ── Écriture ligne par ligne ──
  var ROW = 1; // compteur de ligne courant

  // Titre principal
  var titleRange = kpiSheet.getRange(ROW, 1, 1, 7);
  titleRange.merge();
  titleRange.setValue('MAISON ABEILLE — Tableau de bord des KPI');
  titleRange.setBackground(C.dark);
  titleRange.setFontColor(C.gold);
  titleRange.setFontWeight('bold');
  titleRange.setFontSize(15);
  titleRange.setFontFamily('Arial');
  titleRange.setHorizontalAlignment('center');
  titleRange.setVerticalAlignment('middle');
  kpiSheet.setRowHeight(ROW, 52);
  ROW++;

  // Sous-titre
  var subRange = kpiSheet.getRange(ROW, 1, 1, 7);
  subRange.merge();
  subRange.setValue('Les indicateurs « Enquête – auto » se mettent à jour dès qu\'une nouvelle réponse arrive · Compléter manuellement les autres lignes');
  subRange.setBackground('#F0EBE1');
  subRange.setFontColor(C.grayDark);
  subRange.setFontSize(8);
  subRange.setFontFamily('Arial');
  subRange.setHorizontalAlignment('center');
  subRange.setVerticalAlignment('middle');
  kpiSheet.setRowHeight(ROW, 26);
  ROW++;

  // En-têtes colonnes
  var headers = ['Catégorie', 'Indicateur (KPI)', 'Cible', 'Valeur actuelle', 'Tendance', 'Statut', 'Source'];
  var hRange = kpiSheet.getRange(ROW, 1, 1, 7);
  hRange.setValues([headers]);
  hRange.setBackground(C.dark);
  hRange.setFontColor('#FFFFFF');
  hRange.setFontWeight('bold');
  hRange.setFontSize(9);
  hRange.setFontFamily('Arial');
  hRange.setHorizontalAlignment('center');
  hRange.setVerticalAlignment('middle');
  hRange.setBorder(false, false, true, false, false, false, C.gold, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  kpiSheet.setRowHeight(ROW, 38);
  ROW++;

  var firstDataRow = ROW;
  var alternateIndex = 0;

  for (var i = 0; i < kpis.length; i++) {
    var k = kpis[i];

    if (k.type === 'sep') {
      // ── Séparateur de catégorie ──
      kpiSheet.setRowHeight(ROW, 34);
      var sepRange = kpiSheet.getRange(ROW, 1, 1, 7);
      sepRange.merge();
      sepRange.setValue(k.label);
      sepRange.setBackground(C.gold);
      sepRange.setFontColor('#FFFFFF');
      sepRange.setFontWeight('bold');
      sepRange.setFontSize(9);
      sepRange.setFontFamily('Arial');
      sepRange.setVerticalAlignment('middle');
      sepRange.setHorizontalAlignment('left');
      sepRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      alternateIndex = 0;

    } else {
      // ── Ligne KPI ──
      kpiSheet.setRowHeight(ROW, 36);
      var bg = (alternateIndex % 2 === 0) ? C.white : C.gray;

      // Catégorie
      var c1 = kpiSheet.getRange(ROW, 1);
      c1.setValue(k.cat);
      c1.setBackground(bg);
      c1.setFontColor(C.grayDark);
      c1.setFontSize(8);
      c1.setFontFamily('Arial');
      c1.setVerticalAlignment('middle');

      // Nom KPI
      var c2 = kpiSheet.getRange(ROW, 2);
      c2.setValue(k.name);
      c2.setBackground(bg);
      c2.setFontSize(9);
      c2.setFontFamily('Arial');
      c2.setFontWeight('normal');
      c2.setVerticalAlignment('middle');
      c2.setFontColor(C.dark);

      // Cible
      var c3 = kpiSheet.getRange(ROW, 3);
      c3.setValue(k.target);
      c3.setBackground(bg);
      c3.setFontSize(9);
      c3.setFontFamily('Arial');
      c3.setHorizontalAlignment('center');
      c3.setVerticalAlignment('middle');
      c3.setFontColor(C.grayDark);

      // Valeur actuelle
      var c4 = kpiSheet.getRange(ROW, 4);
      if (k.type === 'auto' && k.formula) {
        c4.setFormula(k.formula);
      }
      c4.setBackground(bg);
      c4.setFontSize(11);
      c4.setFontFamily('Arial');
      c4.setFontWeight('bold');
      c4.setHorizontalAlignment('center');
      c4.setVerticalAlignment('middle');
      c4.setFontColor(C.dark);

      // Tendance (saisie manuelle — liste déroulante)
      var c5 = kpiSheet.getRange(ROW, 5);
      c5.setBackground(bg);
      c5.setFontSize(12);
      c5.setHorizontalAlignment('center');
      c5.setVerticalAlignment('middle');

      // Statut (saisie manuelle — liste déroulante)
      var c6 = kpiSheet.getRange(ROW, 6);
      c6.setBackground(bg);
      c6.setFontSize(9);
      c6.setFontFamily('Arial');
      c6.setHorizontalAlignment('center');
      c6.setVerticalAlignment('middle');

      // Source
      var c7 = kpiSheet.getRange(ROW, 7);
      c7.setValue(k.source);
      c7.setBackground(k.type === 'auto' ? C.goldLight : bg);
      c7.setFontSize(8);
      c7.setFontFamily('Arial');
      c7.setFontStyle('italic');
      c7.setFontColor(k.type === 'auto' ? '#7A5C2E' : C.grayDark);
      c7.setHorizontalAlignment('center');
      c7.setVerticalAlignment('middle');

      // Bordure inférieure légère
      kpiSheet.getRange(ROW, 1, 1, 7).setBorder(
        false, false, true, false, false, false,
        C.border, SpreadsheetApp.BorderStyle.SOLID
      );

      alternateIndex++;
    }

    ROW++;
  }

  var lastDataRow = ROW - 1;

  // ── Validation données : Tendance (colonne E) ──
  var tendanceRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['↑ Hausse', '→ Stable', '↓ Baisse', ''], true)
    .setAllowInvalid(false)
    .build();
  kpiSheet.getRange(firstDataRow, 5, lastDataRow - firstDataRow + 1, 1)
    .setDataValidation(tendanceRule);

  // ── Validation données : Statut (colonne F) ──
  var statutRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['✅ OK', '⚠️ À surveiller', '🔴 Action requise', ''], true)
    .setAllowInvalid(false)
    .build();
  kpiSheet.getRange(firstDataRow, 6, lastDataRow - firstDataRow + 1, 1)
    .setDataValidation(statutRule);

  // ── Mise en forme conditionnelle : Statut ──
  var statusRange = kpiSheet.getRange(firstDataRow, 6, lastDataRow - firstDataRow + 1, 1);
  var rules = [];

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('✅ OK')
    .setBackground(C.greenLight)
    .setFontColor(C.green)
    .setBold(true)
    .setRanges([statusRange])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('⚠️')
    .setBackground(C.orangeLight)
    .setFontColor(C.orange)
    .setBold(true)
    .setRanges([statusRange])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('🔴')
    .setBackground(C.redLight)
    .setFontColor(C.red)
    .setBold(true)
    .setRanges([statusRange])
    .build());

  // ── Mise en forme conditionnelle : Tendance ──
  var trendRange = kpiSheet.getRange(firstDataRow, 5, lastDataRow - firstDataRow + 1, 1);

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('↑')
    .setFontColor(C.green)
    .setBold(true)
    .setRanges([trendRange])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('↓')
    .setFontColor(C.red)
    .setBold(true)
    .setRanges([trendRange])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('→')
    .setFontColor(C.grayDark)
    .setRanges([trendRange])
    .build());

  kpiSheet.setConditionalFormatRules(rules);

  // ── Bordure extérieure générale ──
  kpiSheet.getRange(3, 1, lastDataRow - 2, 7).setBorder(
    true, true, true, true, false, false,
    C.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // ── Figer les 3 premières lignes ──
  kpiSheet.setFrozenRows(3);

  // ── Masquer la grille de fond ──
  // (non disponible en Apps Script — cosmétique seulement)

  var allSheets = ss.getSheets();
  var sheetNames = allSheets.map(function(s){ return s.getName(); }).join(', ');
  SpreadsheetApp.getUi().alert(
    '✅  Tableau KPI créé !\n\n' +
    'Onglets présents dans ce fichier : ' + sheetNames + '\n\n' +
    '• Les lignes "Enquête – auto" se calculent automatiquement.\n' +
    '• Pour les autres lignes : saisir la valeur dans "Valeur actuelle",\n' +
    '  puis choisir Tendance et Statut via les listes déroulantes.'
  );
}
