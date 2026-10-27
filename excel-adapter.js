/* CHTL – Navigateur TMCD : pont réel vers Excel (Office.js).
   Lecture seule sur le contenu du classeur : seules opérations d'écriture = worksheet.activate()
   et, sur demande explicite de l'utilisateur, le paramètre de document « ouverture automatique du volet ». */
(function (root) {
  'use strict';

  function erreurLisible(e) {
    var code = e && e.code ? String(e.code) : '';
    var msg = (e && e.message) ? String(e.message) : String(e);
    if (/InvalidOperation|Busy|EditMode|cell edit/i.test(code + ' ' + msg))
      return 'Excel est occupé ou en mode édition de cellule. Validez (Entrée) ou annulez (Échap) puis réessayez.';
    if (/ItemNotFound|InvalidArgument/i.test(code)) return 'Feuille introuvable (renommée ou supprimée).';
    return msg;
  }

  function apiOk(version) {
    try { return Office.context.requirements.isSetSupported('ExcelApi', version); } catch (e) { return false; }
  }

  var ExcelAdapter = {
    mode: 'excel',

    /* Résout true si l'hôte est Excel, false sinon (page ouverte dans un navigateur). */
    init: function () {
      return new Promise(function (resolve) {
        if (typeof Office === 'undefined' || !Office.onReady) { resolve({ ok: false, raison: 'Office.js non chargé' }); return; }
        var fini = false;
        var t = setTimeout(function () { if (!fini) { fini = true; resolve({ ok: false, raison: 'Délai dépassé : hôte Office non détecté' }); } }, 6000);
        Office.onReady(function (info) {
          if (fini) return; fini = true; clearTimeout(t);
          if (info && info.host === Office.HostType.Excel) resolve({ ok: true, plateforme: info.platform });
          else resolve({ ok: false, raison: 'Ce volet doit être ouvert dans Excel.' });
        });
      });
    },

    nomClasseur: function () {
      try {
        var u = Office.context.document.url || '';
        return decodeURIComponent(u.split(/[\\/]/).pop().split('?')[0]) || '';
      } catch (e) { return ''; }
    },

    lister: function () {
      return Excel.run(function (ctx) {
        var coll = ctx.workbook.worksheets;
        var champs = 'items/id,items/name,items/visibility,items/position';
        var tab = apiOk('1.7');
        if (tab) champs += ',items/tabColor';
        coll.load(champs);
        var actif = coll.getActiveWorksheet();
        actif.load('id');
        return ctx.sync().then(function () {
          return {
            actifId: actif.id,
            feuilles: coll.items.map(function (w) {
              return { id: w.id, name: w.name, visibility: w.visibility, position: w.position, tabColor: tab ? (w.tabColor || '') : '' };
            })
          };
        });
      });
    },

    /* Active réellement la feuille (identifiée par son id : survit aux renommages). */
    activer: function (id) {
      return Excel.run(function (ctx) {
        var ws = ctx.workbook.worksheets.getItemOrNullObject(id);
        ws.load('isNullObject,name,visibility');
        return ctx.sync().then(function () {
          if (ws.isNullObject) return { ok: false, raison: 'absente' };
          if (ws.visibility !== 'Visible') return { ok: false, raison: 'masquee', nom: ws.name };
          ws.activate();
          return ctx.sync().then(function () { return { ok: true, nom: ws.name }; });
        });
      }).catch(function (e) { return { ok: false, raison: 'erreur', message: erreurLisible(e) }; });
    },

    /* Abonnements aux évènements Excel ; chaque abonnement est isolé car tous ne sont pas dispo selon la version. */
    abonner: function (cb) {
      var evenements = [
        ['onActivated', function (e) { cb.activation && cb.activation(e.worksheetId); }],
        ['onAdded', function () { cb.structure && cb.structure(); }],
        ['onDeleted', function () { cb.structure && cb.structure(); }],
        ['onNameChanged', function () { cb.structure && cb.structure(); }],
        ['onVisibilityChanged', function () { cb.structure && cb.structure(); }],
        ['onMoved', function () { cb.structure && cb.structure(); }]
      ];
      var actifs = [];
      evenements.reduce(function (p, ev) {
        return p.then(function () {
          return Excel.run(function (ctx) {
            var src = ctx.workbook.worksheets[ev[0]];
            if (!src || !src.add) return;
            var h = src.add(function (e) { ev[1](e); return Promise.resolve(); });
            return ctx.sync().then(function () { actifs.push(ev[0]); });
          }).catch(function () { /* évènement non supporté : le sondage prend le relais */ });
        });
      }, Promise.resolve());
      return actifs;
    },

    /* Ouverture automatique du volet avec ce classeur (paramètre de document). */
    lireOuvertureAuto: function () {
      try { return !!Office.context.document.settings.get('Office.AutoShowTaskpaneWithDocument'); } catch (e) { return false; }
    },
    ecrireOuvertureAuto: function (valeur) {
      return new Promise(function (resolve) {
        try {
          Office.context.document.settings.set('Office.AutoShowTaskpaneWithDocument', !!valeur);
          Office.context.document.settings.saveAsync(function (r) {
            resolve(r.status === Office.AsyncResultStatus.Succeeded ? { ok: true } : { ok: false, message: (r.error && r.error.message) || 'Enregistrement impossible' });
          });
        } catch (e) { resolve({ ok: false, message: erreurLisible(e) }); }
      });
    }
  };

  root.ExcelAdapter = ExcelAdapter;
  if (typeof module === 'object' && module.exports) module.exports = ExcelAdapter;
})(typeof self !== 'undefined' ? self : this);
