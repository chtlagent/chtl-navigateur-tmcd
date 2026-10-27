/* CHTL – Navigateur TMCD : classeur FICTIF pour l'aperçu hors Excel (?demo=1).
   Aucune donnée réelle : noms génériques uniquement. Même interface que ExcelAdapter. */
(function (root) {
  'use strict';
  var noms = ['Accueil', 'a. Famille A >>>', 'Synthèse A', 'Détail A', 'Paramètres A', 'Base cachée', 'b. Famille B >>>', 'Suivi 2026', 'Suivi 2025',
    'Suivi 2024', 'c. Famille C >>>', 'Balance', 'Graphique balance', 'd. Famille D >>>', 'Données D'];
  var seq = 0;
  var feuilles = noms.map(function (n, i) { return { id: '{demo-' + (++seq) + '}', name: n, visibility: n === 'Base cachée' ? 'Hidden' : 'Visible', position: i, tabColor: /Suivi/.test(n) ? '#2E9CCA' : /Balance|Graph/.test(n) ? '#C00000' : '' }; });
  var actif = feuilles[0].id;
  var cb = {};
  function copie() { return feuilles.map(function (f) { return Object.assign({}, f); }); }
  function renum() { feuilles.forEach(function (f, i) { f.position = i; }); }

  var DemoAdapter = {
    mode: 'demo',
    init: function () { return Promise.resolve({ ok: true, plateforme: 'demo' }); },
    nomClasseur: function () { return 'Classeur de démonstration (fictif)'; },
    lister: function () { return Promise.resolve({ actifId: actif, feuilles: copie() }); },
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

  /* Simulation de modifications faites dans Excel (utilisé pour les tests manuels). */
  DemoAdapter.simuler = {
    renommer: function (ancien, nouveau) { feuilles.forEach(function (f) { if (f.name === ancien) f.name = nouveau; }); },
    supprimer: function (nom) { feuilles = feuilles.filter(function (f) { return f.name !== nom; }); renum(); },
    ajouter: function (nom, apres) {
      var i = feuilles.map(function (f) { return f.name; }).indexOf(apres);
      feuilles.splice(i < 0 ? feuilles.length : i + 1, 0, { id: '{demo-' + (++seq) + '}', name: nom, visibility: 'Visible', position: 0, tabColor: '' }); renum();
    },
    activerDansExcel: function (nom) { var f = feuilles.filter(function (x) { return x.name === nom; })[0]; if (f) { actif = f.id; cb.activation && cb.activation(f.id); } }
  };

  root.DemoAdapter = DemoAdapter;
})(typeof self !== 'undefined' ? self : this);
