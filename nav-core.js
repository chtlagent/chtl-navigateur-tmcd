/* CHTL – Navigateur de feuilles : logique pure (sans Office.js, testable sous Node et navigateur).
   Les familles sont déduites de la COULEUR D'ONGLET : mêmes couleurs = même famille, quel que soit le classeur. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.NavCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var GRIS_AUCUNE = '#8A93A6';

  /* Minuscule + suppression des accents, longueur de chaîne conservée (utile pour la surbrillance). */
  function plie(s) {
    return Array.from(String(s == null ? '' : s)).map(function (c) { return c.normalize('NFD')[0].toLowerCase(); }).join('');
  }
  function compact(s) { return plie(s).replace(/[^a-z0-9]+/g, ''); }
  function estVisible(f) { return !f.visibility || f.visibility === 'Visible'; }

  /* Couleur d'onglet Excel : n'accepte que #RRGGBB (jamais de texte libre dans un style). */
  function couleurOnglet(c) { return /^#[0-9a-fA-F]{6}$/.test(c || '') ? c.toUpperCase() : ''; }

  function hsl(hex) {
    var r = parseInt(hex.substr(1, 2), 16) / 255, g = parseInt(hex.substr(3, 2), 16) / 255, b = parseInt(hex.substr(5, 2), 16) / 255;
    var max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2, d = max - min, h = 0, s = 0;
    if (d) {
      s = d / (1 - Math.abs(2 * l - 1));
      if (max === r) h = ((g - b) / d) % 6; else if (max === g) h = (b - r) / d + 2; else h = (r - g) / d + 4;
      h = (h * 60 + 360) % 360;
    }
    return { h: h, s: s, l: l };
  }

  /* Nom français approximatif d'une couleur d'onglet (renommable par l'utilisateur dans le volet). */
  function nomCouleur(hex) {
    var c = hsl(hex), n;
    if (c.l < 0.12) return 'Noir';
    if (c.l > 0.94) return 'Blanc';
    if (c.s < 0.14) return c.l < 0.5 ? 'Gris foncé' : 'Gris';
    var h = c.h;
    if (h < 15 || h >= 345) n = 'Rouge';
    else if (h < 45) n = (c.l < 0.4 ? 'Marron' : 'Orange');
    else if (h < 70) n = 'Jaune';
    else if (h < 165) n = 'Vert';
    else if (h < 205) n = 'Bleu ciel';
    else if (h < 255) n = 'Bleu';
    else if (h < 290) n = 'Violet';
    else n = 'Rose';
    if (n === 'Marron') return n;
    if (c.l > 0.78) return n + ' clair';
    if (c.l < 0.28) return n + ' foncé';
    return n;
  }

  /* Regroupe les feuilles visibles en familles selon la couleur d'onglet.
     cfg.noms : libellés personnalisés par couleur ({ "#C00000": "TMCD" }). */
  function grouper(feuilles, cfg) {
    cfg = cfg || {};
    var noms = cfg.noms || {};
    var triees = feuilles.slice().sort(function (a, b) { return a.position - b.position; });
    var familles = [], index = {}, masquees = 0;
    triees.forEach(function (f) {
      if (!estVisible(f)) { masquees++; return; }
      var hex = couleurOnglet(f.tabColor);
      var cle = hex || 'aucune';
      var fa = index[cle];
      if (!fa) {
        fa = index[cle] = { cle: cle, couleurHex: hex, couleurAffichee: hex || GRIS_AUCUNE, libelleAuto: hex ? nomCouleur(hex) : 'Sans couleur', feuilles: [] };
        familles.push(fa);
      }
      fa.feuilles.push(f);
    });
    /* deux nuances différentes de même nom : « Bleu », « Bleu (2) » */
    var vus = {};
    familles.forEach(function (fa) {
      vus[fa.libelleAuto] = (vus[fa.libelleAuto] || 0) + 1;
      if (vus[fa.libelleAuto] > 1) fa.libelleAuto += ' (' + vus[fa.libelleAuto] + ')';
      var perso = noms[fa.cle];
      fa.personnalise = !!(perso && String(perso).trim());
      fa.libelle = fa.personnalise ? String(perso).trim() : fa.libelleAuto;
    });
    return { familles: familles, masquees: masquees };
  }

  function toutesFeuilles(groupes) {
    var out = [];
    groupes.familles.forEach(function (fa) { fa.feuilles.forEach(function (f) { out.push({ feuille: f, famille: fa }); }); });
    return out;
  }

  /* Recherche multi-mots, insensible à la casse/accents ; renvoie les plages à surligner. */
  function rechercher(groupes, requete) {
    var jetons = plie(requete).split(/\s+/).filter(Boolean);
    if (!jetons.length) return [];
    var res = [];
    toutesFeuilles(groupes).forEach(function (e) {
      var nom = plie(e.feuille.name), nomCompact = compact(e.feuille.name), famille = plie(e.famille.libelle);
      var score = 0, plages = [], ok = true;
      jetons.forEach(function (j) {
        var i = nom.indexOf(j);
        if (i >= 0) {
          plages.push([i, i + j.length]);
          score += (i === 0 ? 3 : (/[^a-z0-9]/.test(nom[i - 1]) ? 2 : 1));
        } else if (compact(j) && nomCompact.indexOf(compact(j)) >= 0) score += 0.5;
        else if (famille.indexOf(j) >= 0) score += 0.3; /* le mot désigne la famille */
        else ok = false;
      });
      if (ok) res.push({ feuille: e.feuille, famille: e.famille, plages: fusion(plages), score: score - e.feuille.name.length / 1000 });
    });
    res.sort(function (a, b) { return b.score - a.score || a.feuille.position - b.feuille.position; });
    return res;
  }

  function fusion(plages) {
    var p = plages.slice().sort(function (a, b) { return a[0] - b[0]; }), out = [];
    p.forEach(function (r) {
      var d = out[out.length - 1];
      if (d && r[0] <= d[1]) d[1] = Math.max(d[1], r[1]); else out.push([r[0], r[1]]);
    });
    return out;
  }

  /* Raccourcis définis dans config.json, résolus sur les noms réels. Introuvable = feuille null (masqué dans le volet). */
  function resoudreRaccourcis(feuilles, defs, annee) {
    var visibles = feuilles.filter(estVisible).sort(function (a, b) { return a.position - b.position; });
    return (defs || []).map(function (d) {
      var re;
      try { re = new RegExp(String(d.motif || '').replace(/\{annee\}/g, String(annee || new Date().getFullYear())), 'i'); }
      catch (e) { return { def: d, feuille: null }; }
      var cand = visibles.filter(function (f) { return re.test(f.name); }), pick = null;
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

  /* Historique Retour / Avant (comme les boutons latéraux d'une souris). */
  function creerHistorique(max) {
    max = max || 50;
    var arriere = [], avant = [], courant = null;
    function cherche(pile, existe) {
      while (pile.length) { var id = pile.pop(); if (id !== courant && existe(id)) return id; }
      return null;
    }
    return {
      visiter: function (id) {
        if (id == null || id === courant) return;
        if (courant != null) { arriere.push(courant); if (arriere.length > max) arriere.shift(); }
        avant = []; courant = id;
      },
      retour: function (existe) {
        var id = cherche(arriere, existe);
        if (id == null) return null;
        if (courant != null) avant.push(courant);
        courant = id; return id;
      },
      avancer: function (existe) {
        var id = cherche(avant, existe);
        if (id == null) return null;
        if (courant != null) arriere.push(courant);
        courant = id; return id;
      },
      precedent: function (existe) { for (var i = arriere.length - 1; i >= 0; i--) if (arriere[i] !== courant && existe(arriere[i])) return arriere[i]; return null; },
      suivant: function (existe) { for (var i = avant.length - 1; i >= 0; i--) if (avant[i] !== courant && existe(avant[i])) return avant[i]; return null; },
      peutRevenir: function (existe) { return this.precedent(existe) != null; },
      courant: function () { return courant; }
    };
  }

  /* Feuilles consultées récemment (la plus récente d'abord), sans doublon. */
  function creerRecents(max) {
    max = max || 8;
    var liste = [];
    return {
      visiter: function (id) {
        if (id == null) return;
        liste = liste.filter(function (x) { return x !== id; });
        liste.unshift(id);
        if (liste.length > max) liste.length = max;
      },
      charger: function (ids) { liste = (ids || []).slice(0, max); },
      ids: function () { return liste.slice(); },
      lister: function (existe, exclure, n) { return liste.filter(function (id) { return id !== exclure && existe(id); }).slice(0, n || 5); }
    };
  }

  function compterVisibles(feuilles) { return feuilles.filter(estVisible).length; }

  function famillePourId(groupes, id) {
    for (var i = 0; i < groupes.familles.length; i++) if (groupes.familles[i].feuilles.some(function (f) { return f.id === id; })) return groupes.familles[i];
    return null;
  }

  function signature(feuilles) {
    return feuilles.map(function (f) { return [f.id, f.name, f.visibility, f.position, f.tabColor || ''].join('|'); }).join('\n');
  }

  /* « 12. TMCD PILOTAGE v4.xlsm » -> « 12. TMCD PILOTAGE v4 » */
  function titreClasseur(nom) {
    var n = String(nom || '').replace(/\.(xlsx|xlsm|xlsb|xls|xltx|xltm)$/i, '').trim();
    return n || 'Classeur Excel';
  }

  return { plie: plie, compact: compact, grouper: grouper, toutesFeuilles: toutesFeuilles, rechercher: rechercher,
    resoudreRaccourcis: resoudreRaccourcis, creerHistorique: creerHistorique, creerRecents: creerRecents, signature: signature,
    compterVisibles: compterVisibles, couleurOnglet: couleurOnglet, nomCouleur: nomCouleur, famillePourId: famillePourId,
    titreClasseur: titreClasseur, GRIS_AUCUNE: GRIS_AUCUNE };
});
