/* CHTL – Navigateur TMCD : logique pure (sans Office.js, testable sous Node et navigateur). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.NavCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var COULEURS = ['bleu', 'vert', 'orange'];

  var DEFAUTS = {
    motifSeparateur: '^\\s*([A-Za-z])\\s*[.)]\\s+(.+?)\\s*(?:>{2,}|»+)?\\s*$',
    libelleDebut: 'Général',
    raccourcis: []
  };

  /* Minuscule + suppression des accents, longueur de chaîne conservée (utile pour la surbrillance). */
  function plie(s) {
    return Array.from(String(s == null ? '' : s)).map(function (c) {
      var d = c.normalize('NFD')[0];
      return d.toLowerCase();
    }).join('');
  }
  function compact(s) { return plie(s).replace(/[^a-z0-9]+/g, ''); }

  function estVisible(f) { return !f.visibility || f.visibility === 'Visible'; }

  /* Regroupe les feuilles visibles en familles. Une feuille « x. Titre >>> » ouvre une famille. */
  function grouper(feuilles, cfg) {
    cfg = Object.assign({}, DEFAUTS, cfg || {});
    var re;
    try { re = new RegExp(cfg.motifSeparateur); } catch (e) { re = new RegExp(DEFAUTS.motifSeparateur); }
    var triees = feuilles.slice().sort(function (a, b) { return a.position - b.position; });
    var familles = [];
    var courante = null;
    var masquees = 0;
    triees.forEach(function (f) {
      if (!estVisible(f)) { masquees++; return; }
      var m = re.exec(f.name);
      if (m) {
        courante = { cle: 'f' + familles.length + ':' + plie(m[1]), lettre: m[1].toUpperCase(), libelle: (m[2] || f.name).trim(), separateur: f, feuilles: [] };
        familles.push(courante);
      } else {
        if (!courante) {
          courante = { cle: 'debut', lettre: '', libelle: cfg.libelleDebut, separateur: null, feuilles: [] };
          familles.push(courante);
        }
        courante.feuilles.push(f);
      }
    });
    familles.forEach(function (fa, i) { fa.couleur = COULEURS[i % COULEURS.length]; });
    return { familles: familles, masquees: masquees };
  }

  function toutesFeuilles(groupes) {
    var out = [];
    groupes.familles.forEach(function (fa) {
      if (fa.separateur) out.push({ feuille: fa.separateur, famille: fa, separateur: true });
      fa.feuilles.forEach(function (f) { out.push({ feuille: f, famille: fa, separateur: false }); });
    });
    return out;
  }

  /* Recherche multi-mots, insensible à la casse/accents ; renvoie les plages à surligner. */
  function rechercher(groupes, requete) {
    var jetons = plie(requete).split(/\s+/).filter(Boolean);
    if (!jetons.length) return [];
    var res = [];
    toutesFeuilles(groupes).forEach(function (e) {
      var nom = plie(e.feuille.name);
      var nomCompact = compact(e.feuille.name);
      var score = 0, plages = [], ok = true;
      jetons.forEach(function (j) {
        var i = nom.indexOf(j);
        if (i >= 0) {
          plages.push([i, i + j.length]);
          score += (i === 0 ? 3 : (/[^a-z0-9]/.test(nom[i - 1]) ? 2 : 1));
        } else if (compact(j) && nomCompact.indexOf(compact(j)) >= 0) {
          score += 0.5;
        } else ok = false;
      });
      if (ok) {
        /* préférer les noms courts à score égal */
        res.push({ feuille: e.feuille, famille: e.famille, separateur: e.separateur, plages: fusion(plages), score: score - e.feuille.name.length / 1000 });
      }
    });
    res.sort(function (a, b) { return b.score - a.score || a.feuille.position - b.feuille.position; });
    return res;
  }

  function fusion(plages) {
    var p = plages.slice().sort(function (a, b) { return a[0] - b[0]; });
    var out = [];
    p.forEach(function (r) {
      var d = out[out.length - 1];
      if (d && r[0] <= d[1]) d[1] = Math.max(d[1], r[1]); else out.push([r[0], r[1]]);
    });
    return out;
  }

  /* Raccourcis définis dans config.json, résolus sur les noms réels (donc robustes aux années qui changent). */
  function resoudreRaccourcis(feuilles, defs, annee) {
    var visibles = feuilles.filter(estVisible).sort(function (a, b) { return a.position - b.position; });
    return (defs || []).map(function (d) {
      var re;
      try { re = new RegExp(String(d.motif || '').replace(/\{annee\}/g, String(annee || new Date().getFullYear())), 'i'); }
      catch (e) { return { def: d, feuille: null }; }
      var cand = visibles.filter(function (f) { return re.test(f.name); });
      var pick = null;
      if (cand.length) {
        if (d.choix === 'dernier') pick = cand[cand.length - 1];
        else if (d.choix === 'plusRecent') {
          var num = function (f) { var m = /(\d{4})/.exec(f.name); return m ? parseInt(m[1], 10) : -1; };
          pick = cand.reduce(function (a, b) { return num(b) > num(a) ? b : a; });
        } else pick = cand[0];
      }
      return { def: d, feuille: pick };
    });
  }

  /* Historique pour le bouton Retour. */
  function creerHistorique(max) {
    max = max || 50;
    var pile = [], courant = null;
    return {
      visiter: function (id) {
        if (id == null || id === courant) return;
        if (courant != null) { pile.push(courant); if (pile.length > max) pile.shift(); }
        courant = id;
      },
      retour: function (existe) {
        while (pile.length) {
          var id = pile.pop();
          if (id !== courant && existe(id)) { courant = id; return id; }
        }
        return null;
      },
      peutRevenir: function (existe) { return pile.some(function (id) { return id !== courant && existe(id); }); },
      courant: function () { return courant; }
    };
  }

  function signature(feuilles) {
    return feuilles.map(function (f) { return [f.id, f.name, f.visibility, f.position, f.tabColor || ''].join('|'); }).join('\n');
  }

  return { plie: plie, compact: compact, grouper: grouper, toutesFeuilles: toutesFeuilles, rechercher: rechercher,
    resoudreRaccourcis: resoudreRaccourcis, creerHistorique: creerHistorique, signature: signature, DEFAUTS: DEFAUTS, COULEURS: COULEURS };
});
