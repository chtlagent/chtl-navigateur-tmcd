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

  /* ---------- Nom automatique des familles : analyse des feuilles du classeur ouvert ----------
     Ordre : 1) feuille séparatrice « x. Titre » / « Titre >>> » ; 2) famille d'une seule feuille = nom de la feuille ;
     3) mot dominant dans les noms (« P&L », « BCA », « Matrice HR », « ICR 24-25 ») ;
     4) section du classeur (dernière feuille séparatrice située avant) ; 5) liste des feuilles. */
  var RE_SEPARATEUR = /^\s*(?:[A-Za-z]\s*[.)]\s+(.+?)|(.+?)\s*(?:>{2,}|»+))\s*(?:>{2,}|»+)?\s*$/;
  var MOTS_VIDES = { de: 1, du: 1, des: 1, la: 1, le: 1, les: 1, et: 1, en: 1, au: 1, aux: 1, sur: 1, par: 1, pour: 1, and: 1, the: 1, of: 1, feuil: 1, feuille: 1, sheet: 1 };

  function titreSeparateur(nom) {
    var m = RE_SEPARATEUR.exec(nom);
    if (!m) return '';
    var t = (m[1] || m[2] || '').trim();
    return t.length >= 2 ? t : '';
  }
  function motsDe(nom) { return String(nom).split(/[\s_\-–—:.,()\/]+/).filter(Boolean); }
  function motUtile(m) { return m.length >= 2 && !/\d/.test(m) && !MOTS_VIDES[plie(m)]; }

  /* Mot présent dans au moins 60 % des feuilles (et ≥ 2), étendu aux mots suivants communs, + suffixe numérique commun. */
  function nomParMotDominant(feuilles) {
    var n = feuilles.length;
    if (n < 2) return null;
    var stats = {}, ordre = [];
    feuilles.forEach(function (f, idx) {
      var vus = {};
      motsDe(f.name).forEach(function (m, i) {
        if (!motUtile(m)) return;
        var k = plie(m);
        if (vus[k]) return; vus[k] = true;
        if (!stats[k]) { stats[k] = { affiche: m, nb: 0, pos: 0 }; ordre.push(k); }
        stats[k].nb++; stats[k].pos += i;
      });
    });
    var seuil = Math.max(2, Math.ceil(n * 0.6)), best = null;
    ordre.forEach(function (k) {
      var s = stats[k];
      if (s.nb < seuil) return;
      if (!best || s.nb > stats[best].nb || (s.nb === stats[best].nb && s.pos < stats[best].pos)) best = k;
    });
    if (!best) return null;
    var avec = feuilles.filter(function (f) { return motsDe(f.name).some(function (m) { return plie(m) === best; }); });
    var nom = [stats[best].affiche];
    for (;;) { /* prolonge : « Matrice » -> « Matrice HR » si tous les noms continuent par le même mot */
      var suivants = avec.map(function (f) {
        var w = motsDe(f.name), i = w.map(plie).indexOf(best); i = -1 === i ? -1 : i + nom.length;
        return i >= 0 && i < w.length ? w[i] : null;
      });
      if (suivants[0] && motUtile(suivants[0]) && suivants.every(function (s) { return s && plie(s) === plie(suivants[0]); })) nom.push(suivants[0]); else break;
    }
    var suf = avec.map(function (f) { var m = /(\d[\d\-\/]*\d|\d)\s*$/.exec(f.name); return m ? m[1] : null; });
    var res = nom.join(' ');
    if (avec.length >= 2 && suf[0] && suf.every(function (s) { return s === suf[0]; })) res += ' ' + suf[0];
    return res;
  }

  function nomListe(feuilles) {
    var s = '';
    for (var i = 0; i < feuilles.length; i++) {
      var ajout = (s ? ' · ' : '') + feuilles[i].name;
      if (s && (s + ajout).length > 30) return s + ' …';
      s += ajout;
    }
    return s;
  }

  function sectionDe(f, seps) {
    var cur = null;
    for (var i = 0; i < seps.length; i++) { if (seps[i].position < f.position) cur = seps[i]; else break; }
    return cur;
  }

  function nomAuto(fa, seps, utilises) {
    var fs = fa.feuilles, titres = [];
    fs.forEach(function (f) { var t = titreSeparateur(f.name); if (t) titres.push(t); });
    if (titres.length === 1) return titres[0];
    if (titres.length > 1 && titres.length >= fs.length / 2) return 'Sections';
    if (fs.length === 1 && fa.couleurHex) return fs[0].name;
    var mot = nomParMotDominant(fs);
    if (mot) return mot;
    var sec = fs.map(function (f) { return sectionDe(f, seps); });
    if (sec[0] && sec.every(function (s) { return s === sec[0]; }) && !utilises[sec[0].titre]) return sec[0].titre;
    return fa.couleurHex ? nomListe(fs) : 'Autres feuilles';
  }

  /* Regroupe les feuilles visibles en familles selon la couleur d'onglet.
     cfg.noms : libellés personnalisés par couleur ({ "#C00000": "TMCD" }). */
  function grouper(feuilles, cfg) {
    cfg = cfg || {};
    var noms = cfg.noms || {};
    var triees = feuilles.slice().sort(function (a, b) { return a.position - b.position; });
    var familles = [], index = {}, masquees = 0, seps = [];
    triees.forEach(function (f) {
      if (!estVisible(f)) { masquees++; return; }
      var t = titreSeparateur(f.name);
      if (t) seps.push({ position: f.position, titre: t });
      var hex = couleurOnglet(f.tabColor);
      var cle = hex || 'aucune';
      var fa = index[cle];
      if (!fa) {
        fa = index[cle] = { cle: cle, couleurHex: hex, couleurAffichee: hex || GRIS_AUCUNE, couleurNom: hex ? nomCouleur(hex) : 'Sans couleur', feuilles: [] };
        familles.push(fa);
      }
      fa.feuilles.push(f);
    });
    var vus = {}, utilises = {};
    familles.forEach(function (fa) {
      var auto = nomAuto(fa, seps, utilises) || fa.couleurNom;
      utilises[auto] = true;
      vus[auto] = (vus[auto] || 0) + 1;
      fa.libelleAuto = vus[auto] > 1 ? auto + ' (' + vus[auto] + ')' : auto;
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
    creerHistorique: creerHistorique, signature: signature,
    compterVisibles: compterVisibles, couleurOnglet: couleurOnglet, nomCouleur: nomCouleur, famillePourId: famillePourId,
    titreClasseur: titreClasseur, GRIS_AUCUNE: GRIS_AUCUNE };
});
