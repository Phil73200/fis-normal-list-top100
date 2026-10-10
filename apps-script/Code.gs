/* ============================================================
 *  TOP 30 / 60 / 100 MONDIAL – FIS
 *
 *  DATA_FIS  = Normal List courante
 *  BASIC_REF = Basic List de référence
 *
 *  AUTOMATISATION
 *  ------------------------------------------------------------
 *  - Saison FIS calculée automatiquement.
 *  - Normal List courante importée automatiquement.
 *  - Basic List la plus récente détectée automatiquement.
 *  - BASIC_REF n'est remplacée que lorsqu'une Basic List
 *    plus récente et valide est réellement disponible.
 *
 *  Identifiants MySQL :
 *    Script Properties :
 *      DB_USER
 *      DB_PASS
 * ============================================================ */


/* ============================================================
 * CONFIGURATION
 * ============================================================ */

const DB_URL =
  'jdbc:mysql://db.fisski.com:3306/FISmainDB';

const SHEET_NAME =
  'DATA_FIS';

const BASIC_SHEET_NAME =
  'BASIC_REF';


/* ============================================================
 * SAISON FIS AUTOMATIQUE
 *
 * juillet 2026 -> saison 2027
 * avril 2027   -> saison 2027
 * juillet 2027 -> saison 2028
 * ============================================================ */

function getFisSeason_() {

  const now = new Date();

  const year =
    now.getFullYear();

  const month =
    now.getMonth() + 1;

  return month >= 7
    ? year + 1
    : year;

}


/* ============================================================
 * IDENTIFIANTS MYSQL
 * ============================================================ */

function getDbCredentials_() {

  const properties =
    PropertiesService
      .getScriptProperties();

  const user =
    properties.getProperty('DB_USER');

  const pass =
    properties.getProperty('DB_PASS');

  if (!user || !pass) {

    throw new Error(
      'Les propriétés DB_USER et/ou DB_PASS sont absentes.'
    );

  }

  return {
    user: user,
    pass: pass
  };

}


/* ============================================================
 * CONNEXION MYSQL
 * ============================================================ */

function getConnection_() {

  const credentials =
    getDbCredentials_();

  return Jdbc.getConnection(
    DB_URL,
    credentials.user,
    credentials.pass
  );

}


/* ============================================================
 * MENU
 * ============================================================ */

function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu('@ Admin')

    .addItem(
      'Mettre à jour Normal List',
      'importerTop100FIS'
    )

    .addItem(
      'Vérifier Basic List',
      'actualiserBasicReference'
    )

    .addSeparator()

    .addItem(
      'Vérifier JSON',
      'verifierJSON'
    )

    .addSeparator()

    .addItem(
      'Créer déclencheur hebdomadaire',
      'creerDeclencheurHebdomadaire'
    )

    .addItem(
      'Ajouter mercredi à 12 h',
      'ajouterDeclencheurMercredi'
    )

    .addToUi();

}


/* ============================================================
 * IMPORT NORMAL LIST COURANTE
 * ============================================================ */

function importerTop100FIS() {

  const startedAt = Date.now();
  const logStep = function(message) {
    console.log('[Normal List +' + Math.round((Date.now() - startedAt) / 1000) + ' s] ' + message);
  };

  logStep('Début import');

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(SHEET_NAME);

  if (!sheet) {

    sheet =
      ss.insertSheet(SHEET_NAME);

  }


  const season =
    getFisSeason_();


  const sql = `

    WITH CurrentList AS (

      SELECT
        Listid,
        Listname,
        Listnumber,
        Seasoncode,
        Validfrom

      FROM A_listdefal

      WHERE Seasoncode = ${season}

        AND Speciallist = 0

        AND Published = 1

        AND Validfrom <= CURDATE()

        AND Validto >= CURDATE()

      ORDER BY
        Validfrom DESC,
        Listid DESC

      LIMIT 1
    )

    SELECT

      c.Fiscode,

      TRIM(CONCAT_WS(' ', c.Lastname, c.Firstname)) AS Fullname,

      YEAR(c.Birthdate) AS Birthyear,

      c.Gender,

      r.Disciplinecode,

      c.Nationcode,

      r.Fispoints,

      r.Position,

      CONCAT(l.Listnumber, '.', RIGHT(CAST(l.Seasoncode AS CHAR), 2)) AS Liste_FIS

    FROM A_listresultal r

    JOIN CurrentList l
      ON l.Listid = r.Listid

    JOIN A_competitor c
      ON c.Competitorid = r.Competitorid

    WHERE

      r.Disciplinecode IN (
        'SL',
        'GS',
        'SG',
        'DH'
      )

      AND r.Position IS NOT NULL

      AND r.Position <= 100

      AND c.Gender IN (
        'M',
        'W'
      )

    ORDER BY

      c.Gender,

      r.Disciplinecode,

      r.Position,

      c.Nationcode,

      c.Lastname,

      c.Firstname

  `;


  logStep('Connexion FIS : début');

  const conn =
    getConnection_();

  logStep('Connexion FIS : établie');

  let stmt = null;
  let rs = null;

  const rows = [];


  try {

    stmt =
      conn.createStatement();

    logStep('Requête SQL : début');

    rs =
      stmt.executeQuery(sql);

    logStep('Requête SQL : résultats disponibles, début lecture');


    // Read in batches to avoid a separate Apps Script service call per cell.
    const getters =
      'getString(1),getString(2),getString(3),' +
      'getString(4),getString(5),getString(6),' +
      'getDouble(7),getInt(8),getString(9)';

    let batch;

    do {

      batch = rs.getRows(getters, 100);

      for (const row of batch) {

        rows.push([

          row[0] || '',

          row[1] || '',

          row[2] || '',

          row[3] || '',

          row[4] || '',

          row[5] || '',

          row[6],

          row[7],

          row[8] || ''

        ]);

      }

      if (batch.length) {
        logStep('Lecture : ' + rows.length + ' lignes');
      }

    } while (batch.length === 100);

    logStep('Lecture terminée : ' + rows.length + ' lignes');

  } finally {

    logStep('Fermeture connexion : début');

    if (rs) {
      rs.close();
    }

    if (stmt) {
      stmt.close();
    }

    conn.close();

    logStep('Fermeture connexion : terminée');

  }


  if (!rows.length) {

    throw new Error(
      'Aucune donnée Normal List trouvée.'
    );

  }


  logStep('Écriture DATA_FIS : début');

  sheet.clearContents();


  sheet
    .getRange(
      1,
      1,
      1,
      9
    )
    .setValues([[

      'Fiscode',

      'Fullname',

      'Birthyear',

      'Gender',

      'Discipline',

      'Nation',

      'Fispoints',

      'WorldRank',

      'Liste_FIS'

    ]]);


  sheet
    .getRange(
      2,
      1,
      rows.length,
      9
    )
    .setValues(rows);


  sheet
    .getRange('J1')
    .setValue('Mise à jour');


  sheet
    .getRange('J2')
    .setValue(new Date());


  sheet
    .getRange('K1')
    .setValue('Saison FIS');


  sheet
    .getRange('K2')
    .setValue(season);


  sheet.setFrozenRows(1);

  sheet.autoResizeColumns(1, 9);

  SpreadsheetApp.flush();

  logStep('Écriture DATA_FIS : terminée');


  /*
   * Vérification automatique de la Basic List.
   *
   * Une erreur de vérification Basic ne doit PAS empêcher
   * la mise à jour de la Normal List.
   */

  try {

    logStep('Vérification Basic List : début');

    actualiserBasicReference();

    logStep('Vérification Basic List : terminée');

  } catch (error) {

    console.log(
      'Vérification Basic List non effectuée : ' +
      error.message
    );

  }


  return (
    'Normal List mise à jour : ' +
    rows.length +
    ' lignes.'
  );

}


/* ============================================================
 * DÉTECTION DE LA BASIC LIST LA PLUS RÉCENTE
 *
 * MÉTHODE VOLONTAIREMENT SIMPLE :
 *
 * SELECT MAX(Seasoncode)
 *
 * On limite la recherche à :
 * saison FIS actuelle + 1
 *
 * Cela permet par exemple de détecter la Basic 2028
 * pendant la fin de saison 2027.
 * ============================================================ */

function detecterDerniereBasicSeason_() {

  const currentSeason =
    getFisSeason_();

  const maxAllowedSeason =
    currentSeason + 1;


  const sql = `

    SELECT
      MAX(Seasoncode)

    FROM A_listresultbaseal

    WHERE
      Seasoncode <= ${maxAllowedSeason}

      AND Disciplinecode IN (
        'SL',
        'GS',
        'SG',
        'DH'
      )

      AND Position IS NOT NULL

      AND Position <= 100

  `;


  const conn =
    getConnection_();

  let stmt = null;
  let rs = null;

  let season = null;


  try {

    stmt =
      conn.createStatement();

    rs =
      stmt.executeQuery(sql);


    if (rs.next()) {

      season =
        Number(
          rs.getInt(1)
        );

    }

  } finally {

    if (rs) {
      rs.close();
    }

    if (stmt) {
      stmt.close();
    }

    conn.close();

  }


  if (
    !season ||
    !Number.isFinite(season)
  ) {

    throw new Error(
      'Impossible de détecter une Basic List.'
    );

  }


  return season;

}


/* ============================================================
 * SAISON ACTUELLEMENT ENREGISTRÉE DANS BASIC_REF
 * ============================================================ */

function getBasicReferenceSeason_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      BASIC_SHEET_NAME
    );


  if (!sheet) {

    return null;

  }


  const value =
    sheet
      .getRange('H2')
      .getValue();


  const season =
    Number(value);


  if (
    !season ||
    !Number.isFinite(season)
  ) {

    return null;

  }


  return season;

}


/* ============================================================
 * CONTRÔLE D'UNE BASIC LIST
 *
 * On vérifie qu'elle contient des données Top 100
 * pour :
 *
 * HOMMES :
 * SL / GS / SG / DH
 *
 * DAMES :
 * SL / GS / SG / DH
 *
 * Donc 8 combinaisons attendues.
 * ============================================================ */

function verifierBasicSeason_(season) {

  const sql = `

    SELECT

      c.Gender,

      b.Disciplinecode,

      COUNT(*)

    FROM A_listresultbaseal b

    JOIN A_competitor c
      ON c.Competitorid = b.Competitorid

    WHERE

      b.Seasoncode = ${season}

      AND b.Disciplinecode IN (
        'SL',
        'GS',
        'SG',
        'DH'
      )

      AND b.Position IS NOT NULL

      AND b.Position <= 100

      AND c.Gender IN (
        'M',
        'W'
      )

    GROUP BY

      c.Gender,

      b.Disciplinecode

  `;


  const conn =
    getConnection_();

  let stmt = null;
  let rs = null;

  const found =
    new Set();


  try {

    stmt =
      conn.createStatement();

    rs =
      stmt.executeQuery(sql);


    while (rs.next()) {

      const gender =
        String(
          rs.getString(1) || ''
        )
        .trim()
        .toUpperCase();


      const discipline =
        String(
          rs.getString(2) || ''
        )
        .trim()
        .toUpperCase();


      const count =
        Number(
          rs.getInt(3)
        );


      if (count > 0) {

        found.add(
          gender +
          '|' +
          discipline
        );

      }

    }

  } finally {

    if (rs) {
      rs.close();
    }

    if (stmt) {
      stmt.close();
    }

    conn.close();

  }


  const expected = [

    'M|SL',
    'M|GS',
    'M|SG',
    'M|DH',

    'W|SL',
    'W|GS',
    'W|SG',
    'W|DH'

  ];


  return expected.every(
    function(key) {

      return found.has(key);

    }
  );

}


/* ============================================================
 * VÉRIFICATION / BASCULE AUTOMATIQUE BASIC
 * ============================================================ */

function actualiserBasicReference() {

  const detectedSeason =
    detecterDerniereBasicSeason_();


  const storedSeason =
    getBasicReferenceSeason_();


  /*
   * CAS NORMAL AUJOURD'HUI :
   *
   * détectée = 2027
   * stockée  = 2027
   *
   * => aucune modification.
   */

  if (
    storedSeason === detectedSeason
  ) {

    return (
      'Basic List ' +
      detectedSeason +
      ' déjà à jour.'
    );

  }


  /*
   * Sécurité :
   * ne jamais revenir automatiquement
   * vers une Basic plus ancienne.
   */

  if (
    storedSeason &&
    detectedSeason < storedSeason
  ) {

    return (
      'Basic List ' +
      storedSeason +
      ' conservée.'
    );

  }


  /*
   * Avant de toucher à BASIC_REF,
   * validation de la nouvelle saison.
   */

  const valid =
    verifierBasicSeason_(
      detectedSeason
    );


  if (!valid) {

    return (
      'Basic List ' +
      detectedSeason +
      ' détectée mais encore incomplète. ' +
      'Basic List ' +
      (storedSeason || '') +
      ' conservée.'
    );

  }


  /*
   * La nouvelle Basic est complète.
   */

  importerBasicReference_(
    detectedSeason
  );


  return (
    'Basic List ' +
    detectedSeason +
    ' installée.'
  );

}


/* ============================================================
 * IMPORT MANUEL / AUTOMATIQUE BASIC
 *
 * Cette fonction publique permet aussi de forcer
 * une vérification depuis Apps Script.
 * ============================================================ */

function importerBasicReference() {

  return actualiserBasicReference();

}


/* ============================================================
 * IMPORT RÉEL D'UNE BASIC LIST
 *
 * Fonction interne.
 *
 * IMPORTANT :
 * BASIC_REF n'est effacée qu'APRÈS avoir récupéré
 * toutes les nouvelles données.
 * ============================================================ */

function importerBasicReference_(season) {

  const sql = `

    SELECT

      c.Nationcode,

      c.Gender,

      b.Disciplinecode,

      SUM(
        CASE
          WHEN b.Position <= 30
          THEN 1
          ELSE 0
        END
      ),

      SUM(
        CASE
          WHEN b.Position <= 60
          THEN 1
          ELSE 0
        END
      ),

      SUM(
        CASE
          WHEN b.Position <= 100
          THEN 1
          ELSE 0
        END
      )

    FROM A_listresultbaseal b

    JOIN A_competitor c
      ON c.Competitorid = b.Competitorid

    WHERE

      b.Seasoncode = ${season}

      AND b.Disciplinecode IN (
        'SL',
        'GS',
        'SG',
        'DH'
      )

      AND b.Position IS NOT NULL

      AND b.Position <= 100

      AND c.Gender IN (
        'M',
        'W'
      )

    GROUP BY

      c.Nationcode,

      c.Gender,

      b.Disciplinecode

    ORDER BY

      c.Gender,

      b.Disciplinecode,

      c.Nationcode

  `;


  const conn =
    getConnection_();

  let stmt = null;
  let rs = null;

  const rows = [];


  try {

    stmt =
      conn.createStatement();

    rs =
      stmt.executeQuery(sql);


    while (rs.next()) {

      rows.push([

        rs.getString(1) || '',

        rs.getString(2) || '',

        rs.getString(3) || '',

        rs.getInt(4),

        rs.getInt(5),

        rs.getInt(6)

      ]);

    }

  } finally {

    if (rs) {
      rs.close();
    }

    if (stmt) {
      stmt.close();
    }

    conn.close();

  }


  /*
   * Si rien n'a été récupéré,
   * BASIC_REF actuelle reste intacte.
   */

  if (!rows.length) {

    throw new Error(
      'Aucune donnée trouvée pour la Basic List ' +
      season +
      '. BASIC_REF conservée.'
    );

  }


  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(
      BASIC_SHEET_NAME
    );


  if (!sheet) {

    sheet =
      ss.insertSheet(
        BASIC_SHEET_NAME
      );

  }


  /*
   * On dispose maintenant de toutes les nouvelles données.
   * On peut remplacer BASIC_REF.
   */

  sheet.clearContents();


  sheet
    .getRange(
      1,
      1,
      1,
      6
    )
    .setValues([[

      'Nation',

      'Gender',

      'Discipline',

      'Top30',

      'Top60',

      'Top100'

    ]]);


  sheet
    .getRange(
      2,
      1,
      rows.length,
      6
    )
    .setValues(rows);


  sheet
    .getRange('H1')
    .setValue(
      'Saison Basic List'
    );


  sheet
    .getRange('H2')
    .setValue(
      season
    );


  sheet
    .getRange('I1')
    .setValue(
      'Créée le'
    );


  sheet
    .getRange('I2')
    .setValue(
      new Date()
    );


  sheet.setFrozenRows(1);

  sheet.autoResizeColumns(1, 9);

  SpreadsheetApp.flush();

}


/* ============================================================
 * LECTURE BASIC_REF
 * ============================================================ */

function lireBasicReference_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      BASIC_SHEET_NAME
    );


  if (!sheet) {

    return {
      season: null,
      data: []
    };

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return {
      season: null,
      data: []
    };

  }


  const season =
    Number(
      sheet
        .getRange('H2')
        .getValue()
    ) || null;


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getValues();


  const data = [];


  for (const row of values) {

    const nation =
      String(
        row[0] || ''
      )
      .trim()
      .toUpperCase();


    const gender =
      String(
        row[1] || ''
      )
      .trim()
      .toUpperCase();


    const discipline =
      String(
        row[2] || ''
      )
      .trim()
      .toUpperCase();


    if (
      !nation ||
      !['M', 'W'].includes(gender) ||
      !['SL', 'GS', 'SG', 'DH'].includes(discipline)
    ) {

      continue;

    }


    data.push({

      Nation:
        nation,

      Gender:
        gender,

      Discipline:
        discipline,

      Top30:
        Number(row[3]) || 0,

      Top60:
        Number(row[4]) || 0,

      Top100:
        Number(row[5]) || 0

    });

  }


  return {

    season:
      season,

    data:
      data

  };

}


/* ============================================================
 * WEB APP
 * ============================================================ */

function doGet() {

  try {

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const sheet =
      ss.getSheetByName(
        SHEET_NAME
      );


    if (!sheet) {

      throw new Error(
        'Feuille DATA_FIS absente.'
      );

    }


    const lastRow =
      sheet.getLastRow();


    if (lastRow < 2) {

      throw new Error(
        'DATA_FIS ne contient aucune donnée.'
      );

    }


    const values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          9
        )
        .getValues();


    const athletes = [];


    for (const row of values) {

      const gender =
        String(
          row[3] || ''
        )
        .trim()
        .toUpperCase();


      const discipline =
        String(
          row[4] || ''
        )
        .trim()
        .toUpperCase();


      const nation =
        String(
          row[5] || ''
        )
        .trim()
        .toUpperCase();


      const worldRank =
        Number(row[7]);


      if (
        !['M', 'W'].includes(gender) ||
        !['SL', 'GS', 'SG', 'DH'].includes(discipline) ||
        !Number.isFinite(worldRank) ||
        worldRank > 100
      ) {

        continue;

      }


      athletes.push({

        Fiscode:
          String(
            row[0] || ''
          ),

        Fullname:
          String(
            row[1] || ''
          ).trim(),

        Birthyear:
          Number(row[2]) || null,

        Gender:
          gender,

        Discipline:
          discipline,

        Nation:
          nation,

        Fispoints:
          Number(row[6]),

        WorldRank:
          worldRank,

        Liste_FIS:
          String(
            row[8] || ''
          ).trim()

      });

    }


    if (!athletes.length) {

      throw new Error(
        'Aucun athlète valide dans DATA_FIS.'
      );

    }


    const basic =
      lireBasicReference_();


    const list =
      String(
        athletes[0].Liste_FIS || ''
      );


    const updatedCell =
      sheet
        .getRange('J2')
        .getValue();


    let updated = '';


    if (updatedCell instanceof Date) {

      updated =
        updatedCell.toISOString();

    } else if (updatedCell) {

      updated =
        String(updatedCell);

    }


    const payload = {

      list:
        list,

      count:
        athletes.length,

      updated:
        updated,

      athletes:
        athletes,

      basicSeason:
        basic.season,

      basic:
        basic.data

    };


    return ContentService

      .createTextOutput(
        JSON.stringify(payload)
      )

      .setMimeType(
        ContentService.MimeType.JSON
      );


  } catch (error) {

    return ContentService

      .createTextOutput(
        JSON.stringify({

          error:
            String(
              error.message ||
              error
            )

        })
      )

      .setMimeType(
        ContentService.MimeType.JSON
      );

  }

}


/* ============================================================
 * DIAGNOSTIC STRUCTURE A_competitor
 *
 * Recherche les colonnes dont le nom évoque
 * la naissance / date / année.
 * Lecture seule : aucune donnée n'est modifiée.
 * ============================================================ */

function diagnostiquerColonnesNaissance() {

  const sql = `

    SELECT
      COLUMN_NAME,
      DATA_TYPE

    FROM INFORMATION_SCHEMA.COLUMNS

    WHERE
      TABLE_SCHEMA = DATABASE()

      AND TABLE_NAME = 'A_competitor'

      AND (
        LOWER(COLUMN_NAME) LIKE '%birth%'
        OR LOWER(COLUMN_NAME) LIKE '%born%'
        OR LOWER(COLUMN_NAME) LIKE '%date%'
        OR LOWER(COLUMN_NAME) LIKE '%year%'
      )

    ORDER BY
      ORDINAL_POSITION

  `;

  const conn = getConnection_();

  let stmt = null;
  let rs = null;

  const columns = [];

  try {

    stmt = conn.createStatement();

    rs = stmt.executeQuery(sql);

    while (rs.next()) {

      columns.push({
        name: rs.getString(1) || '',
        type: rs.getString(2) || ''
      });

    }

  } finally {

    if (rs) {
      rs.close();
    }

    if (stmt) {
      stmt.close();
    }

    conn.close();

  }

  Logger.log(
    JSON.stringify(
      columns,
      null,
      2
    )
  );

  SpreadsheetApp
    .getUi()
    .alert(
      'Colonnes naissance A_competitor',
      columns.length
        ? JSON.stringify(columns, null, 2)
        : 'Aucune colonne correspondante trouvée.',
      SpreadsheetApp
        .getUi()
        .ButtonSet.OK
    );

  return columns;

}


/* ============================================================
 * VÉRIFICATION JSON
 * ============================================================ */

function verifierJSON() {

  const basic =
    lireBasicReference_();


  const fra =
    basic.data.find(
      function(item) {

        return (
          item.Nation === 'FRA' &&
          item.Gender === 'M' &&
          item.Discipline === 'GS'
        );

      }
    );


  const nor =
    basic.data.find(
      function(item) {

        return (
          item.Nation === 'NOR' &&
          item.Gender === 'M' &&
          item.Discipline === 'GS'
        );

      }
    );


  const result = {

    basicSeason:
      basic.season,

    nbLignesBasic:
      basic.data.length,

    FRA_M_GS:
      fra || null,

    NOR_M_GS:
      nor || null

  };


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  SpreadsheetApp
    .getUi()
    .alert(
      'Vérification JSON',
      JSON.stringify(
        result,
        null,
        2
      ),
      SpreadsheetApp
        .getUi()
        .ButtonSet.OK
    );


  return result;

}


/* ============================================================
 * DÉCLENCHEUR HEBDOMADAIRE
 *
 * Chaque mardi vers 16 h :
 *
 * 1. Normal List actualisée
 * 2. Basic List vérifiée automatiquement
 * 3. si nouvelle Basic complète -> bascule automatique
 * ============================================================ */

// Compatibilité avec le menu existant : ajoute le mardi sans supprimer
// les autres déclencheurs, notamment celui du mercredi.
function creerDeclencheurHebdomadaire() {
  const handler = 'importerTop100FISMardi';
  const exists = ScriptApp.getProjectTriggers().some(function(trigger) {
    const name = trigger.getHandlerFunction();
    return trigger.getEventType() === ScriptApp.EventType.CLOCK &&
      (name === handler || name === 'importerTop100FIS');
  });

  if (exists) {
    console.log(
      'Un déclencheur de l’import existe déjà. Il est conservé ; aucun déclencheur du mardi supplémentaire n’a été créé.'
    );
    return;
  }

  ScriptApp.newTrigger(handler)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(16)
    .inTimezone('Europe/Paris')
    .create();

  console.log(
    'Mise à jour ajoutée le mardi entre 16 h et 17 h (heure de Paris). Les autres déclencheurs sont conservés.'
  );
}

function importerTop100FISMardi() {
  return importerTop100FIS();
}

// Ajout du mercredi : aucun déclencheur existant n’est supprimé.
function importerTop100FISMercredi() {
  return importerTop100FIS();
}

function ajouterDeclencheurMercredi() {
  const handler = 'importerTop100FISMercredi';
  const exists = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === handler &&
      trigger.getEventType() === ScriptApp.EventType.CLOCK;
  });

  if (exists) {
    console.log('Le déclencheur du mercredi existe déjà. Aucun changement.');
    return;
  }

  ScriptApp.newTrigger(handler)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
    .atHour(12)
    .inTimezone('Europe/Paris')
    .create();

  console.log(
    'Mise à jour ajoutée le mercredi entre 12 h et 13 h (heure de Paris). Les déclencheurs existants sont conservés.'
  );
}
