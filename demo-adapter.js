/* CHTL – Navigateur de feuilles : classeur FICTIF pour l'aperçu hors Excel (?demo=1).
   Aucune donnée réelle : noms génériques uniquement. Même interface que ExcelAdapter. */
(function (root) {
  'use strict';
  var ROUGE = '#C00000', BLEU = '#2E9CCA', VERT = '#00B050', ORANGE = '#ED7D31';
  var defs = [['Accueil', ''], ['a. Famille A >>>', ROUGE], ['Synthèse A', ROUGE], ['Détail A', ROUGE], ['Paramètres A', ROUGE], ['Base cachée', ROUGE],
    ['b. Famille B >>>', BLEU], ['Suivi 2026', BLEU], ['Suivi 2025', BLEU], ['Suivi 2024', BLEU], ['Balance', VERT], ['Graphique balance', VERT],
    ['Données D', ORANGE], ['Notes', ''], ['Synthèse B', BLEU]];
  var seq = 0, nomClasseur = 'Suivi budgétaire (démo)';
  var feuilles = defs.map(function (d, i) { return { id: '{demo-' + (++seq) + '}', name: d[0], visibility: d[0] === 'Base cachée' ? 'Hidden' : 'Visible', position: i, tabColor: d[1] }; });
  var actif = feuilles[0].id, cb = {};
  function copie() { return feuilles.map(function (f) { return Object.assign({}, f); }); }
  function renum() { feuilles.forEach(function (f, i) { f.position = i; }); }

  var DemoAdapter = {
    mode: 'demo',
    init: function () { return Promise.resolve({ ok: true, plateforme: 'demo' }); },
    nomClasseur: function () { return nomClasseur + '.xlsx'; },
    lister: function () { return Promise.resolve({ actifId: actif, classeur: nomClasseur + '.xlsx', feuilles: copie() }); },
    activer: function (id) {
      var f = feuilles.filter(function (x) { return x.id === id; })[0];
      if (!f) return Promise.resolve({ ok: false, raison: 'absente' });
      if (f.visibility !== 'Visible') return Promise.resolve({ ok: false, raison: 'masquee', nom: f.name });
      actif = id; cb.activation && cb.activation(id);
      return Promise.resolve({ ok: true, nom: f.name });
    },
    abonner: function (c) { cb = c; return ['demo']; },
    _auto: false,
    lireOuvertureAuto: function () { return this._auto; },
    ecrireOuvertureAuto: function (v) { this._auto = !!v; return Promise.resolve({ ok: true }); }
  };

  /* Simulation de modifications faites dans Excel (tests manuels). */
  DemoAdapter.simuler = {
    renommer: function (ancien, nouveau) { feuilles.forEach(function (f) { if (f.name === ancien) f.name = nouveau; }); },
    supprimer: function (nom) { feuilles = feuilles.filter(function (f) { return f.name !== nom; }); renum(); },
    ajouter: function (nom, apres, couleur) {
      var i = feuilles.map(function (f) { return f.name; }).indexOf(apres);
      feuilles.splice(i < 0 ? feuilles.length : i + 1, 0, { id: '{demo-' + (++seq) + '}', name: nom, visibility: 'Visible', position: 0, tabColor: couleur || '' }); renum();
    },
    colorier: function (nom, couleur) { feuilles.forEach(function (f) { if (f.name === nom) f.tabColor = couleur; }); },
    renommerClasseur: function (nom) { nomClasseur = nom; },
    activerDansExcel: function (nom) { var f = feuilles.filter(function (x) { return x.name === nom; })[0]; if (f) { actif = f.id; cb.activation && cb.activation(f.id); } }
  };

  root.DemoAdapter = DemoAdapter;
})(typeof self !== 'undefined' ? self : this);
