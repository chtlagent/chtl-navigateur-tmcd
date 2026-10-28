/* CHTL – Navigateur de feuilles : interface du volet (universelle, tout classeur Excel). */
(function () {
  'use strict';
  var N = window.NavCore;
  var $ = function (id) { return document.getElementById(id); };
  var demo = /[?&]demo=1\b/.test(location.search);
  var adapter = null, cfg = {}, feuilles = [], groupes = { familles: [], masquees: 0 }, actifId = null, sig = '', classeur = '';
  var histo = N.creerHistorique(50), recents = N.creerRecents(8);
  var ferme = {}, favoris = [], noms = {}, defilerVersActif = false, edition = null;
  var stockCle = 'chtl.nav.v2', cleClasseur = 'defaut';

  /* ---------- stockage local (par utilisateur, jamais dans le classeur) ---------- */
  function lire() { try { return JSON.parse(localStorage.getItem(stockCle + ':' + cleClasseur) || '{}'); } catch (e) { return {}; } }
  function ecrire() {
    try { localStorage.setItem(stockCle + ':' + cleClasseur, JSON.stringify({ ferme: ferme, favoris: favoris, recents: recents.ids(), noms: noms })); }
    catch (e) { /* stockage indisponible */ }
  }

  /* ---------- utilitaires ---------- */
  function el(tag, cls, txt) { var e = document.createElement(tag); if (cls) e.className = cls; if (txt != null) e.textContent = txt; return e; }
  function surligner(parent, texte, plages) {
    if (!plages || !plages.length) { parent.textContent = texte; return; }
    var chars = Array.from(texte), pos = 0;
    plages.forEach(function (r) {
      if (r[0] > pos) parent.appendChild(document.createTextNode(chars.slice(pos, r[0]).join('')));
      parent.appendChild(el('mark', null, chars.slice(r[0], r[1]).join('')));
      pos = r[1];
    });
    if (pos < chars.length) parent.appendChild(document.createTextNode(chars.slice(pos).join('')));
  }
  var toastT;
  function toast(msg) { var t = $('toast'); t.textContent = msg; t.hidden = false; clearTimeout(toastT); toastT = setTimeout(function () { t.hidden = true; }, 4500); }
  function avis(msg, erreur) { var a = $('avis'); if (!msg) { a.hidden = true; return; } a.textContent = msg; a.className = 'avis' + (erreur ? ' erreur' : ''); a.hidden = false; }
  function etat(msg, cls) { $('etat').textContent = msg; $('pastille').className = 'pastille ' + (cls || ''); }
  function existe(id) { return feuilles.some(function (f) { return f.id === id; }); }
  function feuilleDe(id) { return feuilles.filter(function (x) { return x.id === id; })[0] || null; }
  function couleurFam(fa) { return fa ? fa.couleurAffichee : N.GRIS_AUCUNE; }

  /* Toute arrivée sur une feuille (clic dans le volet ou action dans Excel) passe ici. */
  function noterVisite(id) {
    histo.visiter(id); recents.visiter(id); actifId = id; defilerVersActif = true;
    var fa = N.famillePourId(groupes, id);
    if (fa && ferme[fa.cle]) delete ferme[fa.cle];
    ecrire();
  }

  /* ---------- navigation ---------- */
  function ouvrir(f) {
    adapter.activer(f.id).then(function (r) {
      if (r.ok) return;
      if (r.raison === 'absente') toast('« ' + f.name + ' » n’existe plus (supprimée ou renommée). Liste actualisée.');
      else if (r.raison === 'masquee') toast('« ' + (r.nom || f.name) + ' » est masquée dans le classeur.');
      else toast(r.message || 'Activation impossible.');
      rafraichir(true);
    });
  }

  function aller(sens) {
    var id = sens < 0 ? histo.retour(existe) : histo.avancer(existe);
    if (id == null) { toast(sens < 0 ? 'Aucune feuille précédente.' : 'Aucune feuille suivante.'); majBoutons(); return; }
    adapter.activer(id).then(function (r) { if (!r.ok) { toast('Cette feuille n’est plus accessible.'); rafraichir(true); } });
    recents.visiter(id); actifId = id; defilerVersActif = true; ecrire(); render();
  }

  function majBoutons() {
    var p = histo.precedent(existe), s = histo.suivant(existe), r = $('retour'), a = $('avant');
    r.disabled = p == null; a.disabled = s == null;
    r.title = p == null ? 'Aucune feuille précédente' : 'Revenir à « ' + (feuilleDe(p) || {}).name + ' »';
    a.title = s == null ? 'Aucune feuille suivante' : 'Aller à « ' + (feuilleDe(s) || {}).name + ' »';
  }

  /* ---------- rendu ---------- */
  function estFav(f) { return favoris.some(function (x) { return x.id === f.id; }); }
  function basculerFav(f) {
    if (estFav(f)) favoris = favoris.filter(function (x) { return x.id !== f.id; }); else favoris.push({ id: f.id, name: f.name });
    ecrire(); render();
  }

  function ligne(f, fa, plages, avecFamille) {
    var li = el('div', 'ligne' + (f.id === actifId ? ' active' : ''));
    li.style.setProperty('--fc', couleurFam(fa));
    var b = el('button', 'feuille'); b.type = 'button'; b.dataset.id = f.id;
    if (f.id === actifId) b.setAttribute('aria-current', 'true');
    b.title = 'Ouvrir la feuille « ' + f.name + ' »';
    var nom = el('span', 'nom'); surligner(nom, f.name, plages); b.appendChild(nom);
    if (avecFamille && fa) b.appendChild(el('span', 'tag', fa.libelle));
    b.addEventListener('click', function () { ouvrir(f); });
    li.appendChild(b);
    var e = el('button', 'etoile' + (estFav(f) ? ' on' : ''), estFav(f) ? '★' : '☆');
    e.type = 'button'; e.title = estFav(f) ? 'Retirer des favoris' : 'Ajouter aux favoris'; e.setAttribute('aria-label', e.title);
    e.addEventListener('click', function (ev) { ev.stopPropagation(); basculerFav(f); });
    li.appendChild(e);
    return li;
  }

  function puces(titre, items) {
    if (!items.length) return null;
    var wrap = el('section'); wrap.appendChild(el('h2', 'section-titre', titre));
    var box = el('div', 'raccourcis');
    items.forEach(function (it) {
      var b = el('button', 'puce ' + it.cls + (it.f.id === actifId ? ' actif' : ''), it.txt); b.type = 'button';
      b.title = 'Ouvrir « ' + it.f.name + ' »'; b.addEventListener('click', function () { ouvrir(it.f); });
      box.appendChild(b);
    });
    wrap.appendChild(box); return wrap;
  }

  function sectionRaccourcis() {
    var items = [];
    N.resoudreRaccourcis(feuilles, cfg.raccourcis, new Date().getFullYear()).forEach(function (r) {
      if (r.feuille) items.push({ cls: '', f: r.feuille, txt: r.feuille.name !== r.def.libelle ? r.def.libelle + ' · ' + r.feuille.name : r.def.libelle });
    });
    favoris.forEach(function (fv) {
      var f = feuilleDe(fv.id) || feuilles.filter(function (x) { return x.name === fv.name; })[0];
      if (f && f.visibility === 'Visible') items.push({ cls: 'fav', f: f, txt: '★ ' + f.name });
    });
    return puces('Raccourcis', items);
  }
  function sectionRecents() {
    return puces('Récents', recents.lister(existe, actifId, 5).map(function (id) { var f = feuilleDe(id); return { cls: 'recent', f: f, txt: f.name }; }));
  }

  function enteteNom(fa) {
    if (edition === fa.cle) {
      var inp = el('input', 'fam-edit'); inp.type = 'text'; inp.value = fa.libelle; inp.maxLength = 40; inp.setAttribute('aria-label', 'Nom de la famille');
      var fin = function (ok) {
        if (edition !== fa.cle) return; edition = null;
        if (ok) { var v = inp.value.trim(); if (v && v !== fa.libelleAuto) noms[fa.cle] = v; else delete noms[fa.cle]; ecrire(); groupes = N.grouper(feuilles, { noms: noms }); }
        render();
      };
      inp.addEventListener('keydown', function (e) { e.stopPropagation(); if (e.key === 'Enter') fin(true); else if (e.key === 'Escape') fin(false); });
      inp.addEventListener('blur', function () { fin(true); });
      setTimeout(function () { inp.focus(); inp.select(); }, 0);
      return inp;
    }
    return el('span', 'fam-nom', fa.libelle);
  }

  function famille(fa) {
    var ouvert = !ferme[fa.cle];
    var box = el('section', 'famille' + (ouvert ? '' : ' ferme')); box.style.setProperty('--fc', couleurFam(fa));
    var tete = el('div', 'fam-tete');
    var bas = el(edition === fa.cle ? 'div' : 'button', 'fam-bascule'); /* pas de champ de saisie dans un bouton */
    if (edition !== fa.cle) { bas.type = 'button'; bas.setAttribute('aria-expanded', String(ouvert)); }
    bas.innerHTML = '<svg class="chev" viewBox="0 0 10 10" aria-hidden="true"><path d="M1 3l4 4 4-4" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>';
    var sw = el('span', 'swatch' + (fa.couleurHex ? '' : ' vide')); sw.setAttribute('aria-hidden', 'true'); bas.appendChild(sw);
    if (edition === fa.cle) { var wrapE = el('span', 'fam-nom'); wrapE.appendChild(enteteNom(fa)); bas.appendChild(wrapE); }
    else bas.appendChild(enteteNom(fa));
    bas.appendChild(el('span', 'nombre', String(fa.feuilles.length)));
    bas.addEventListener('click', function (e) { if (edition === fa.cle) return; ferme[fa.cle] = ouvert; ecrire(); render(); });
    tete.appendChild(bas);
    var ed = el('button', 'ouvrir-sep', '✎'); ed.type = 'button'; ed.title = 'Renommer cette famille (affichage uniquement)'; ed.setAttribute('aria-label', ed.title);
    ed.addEventListener('click', function () { edition = fa.cle; render(); });
    tete.appendChild(ed);
    box.appendChild(tete);
    var corps = el('div', 'fam-corps');
    fa.feuilles.forEach(function (f) { corps.appendChild(ligne(f, fa, null, false)); });
    box.appendChild(corps);
    return box;
  }

  function render() {
    var c = $('contenu'), q = $('q').value.trim(), pos = c.scrollTop;
    c.textContent = '';
    $('effacer').hidden = !q;
    if (q) {
      var res = N.rechercher(groupes, q);
      if (!res.length) c.appendChild(el('div', 'vide', 'Aucune feuille ne correspond à « ' + q + ' ».'));
      else {
        var box = el('div', 'resultats');
        res.forEach(function (r) { box.appendChild(ligne(r.feuille, r.famille, r.plages, true)); });
        c.appendChild(el('h2', 'section-titre', res.length + ' résultat' + (res.length > 1 ? 's' : '')));
        c.appendChild(box);
      }
    } else {
      var s = sectionRaccourcis(); if (s) c.appendChild(s);
      var rc = sectionRecents(); if (rc) c.appendChild(rc);
      c.appendChild(el('h2', 'section-titre', 'Familles (couleurs d’onglets)'));
      groupes.familles.forEach(function (fa) { c.appendChild(famille(fa)); });
      if (!groupes.familles.length) c.appendChild(el('div', 'vide', 'Aucune feuille visible.'));
      if (groupes.masquees) c.appendChild(el('p', 'note', groupes.masquees + ' feuille' + (groupes.masquees > 1 ? 's masquées ne sont' : ' masquée n’est') + ' pas affichée' + (groupes.masquees > 1 ? 's' : '') + '.'));
    }
    c.scrollTop = pos;
    if (defilerVersActif) { defilerVersActif = false; var a = c.querySelector('.ligne.active'); if (a && a.scrollIntoView) a.scrollIntoView({ block: 'nearest' }); }
    var nb = N.compterVisibles(feuilles);
    etat((adapter.mode === 'demo' ? 'Aperçu hors Excel · ' : 'Connecté à Excel · ') + nb + ' feuille' + (nb > 1 ? 's' : '') + ' visible' + (nb > 1 ? 's' : ''), 'ok');
    majBoutons();
  }

  function majTitre(nom) {
    var t = N.titreClasseur(nom || adapter.nomClasseur());
    if (t !== classeur) { classeur = t; $('titre').textContent = t; $('titre').title = nom || t; document.title = 'CHTL – ' + t; }
  }

  /* ---------- synchronisation avec le classeur ---------- */
  function rafraichir(force) {
    return adapter.lister().then(function (r) {
      majTitre(r.classeur);
      var nouvelle = N.signature(r.feuilles), changeActif = r.actifId !== actifId, change = force || nouvelle !== sig || changeActif;
      if (change) { sig = nouvelle; feuilles = r.feuilles; groupes = N.grouper(feuilles, { noms: noms }); }
      if (changeActif) noterVisite(r.actifId);
      if (change && edition == null) render();
    }).catch(function (e) { etat('Erreur de lecture du classeur', 'ko'); avis('Lecture des feuilles impossible : ' + (e && e.message ? e.message : e), true); });
  }
  function surActivation(id) {
    if (id === actifId) return;
    noterVisite(id);
    if (!existe(id)) rafraichir(true); else if (edition == null) render();
  }

  /* ---------- clavier et souris ---------- */
  function boutonsNav() { return Array.prototype.slice.call(document.querySelectorAll('#contenu .feuille, #contenu .puce')); }
  document.addEventListener('keydown', function (e) {
    var q = $('q');
    if (e.altKey && e.key === 'ArrowLeft') { aller(-1); e.preventDefault(); return; }
    if (e.altKey && e.key === 'ArrowRight') { aller(1); e.preventDefault(); return; }
    if (e.key === 'Escape' && q.value) { q.value = ''; render(); q.focus(); e.preventDefault(); return; }
    if (e.key === '/' && document.activeElement !== q && document.activeElement.tagName !== 'INPUT') { q.focus(); q.select(); e.preventDefault(); return; }
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      var l = boutonsNav(); if (!l.length) return;
      var i = l.indexOf(document.activeElement);
      if (document.activeElement === q) i = -1; else if (i < 0) return;
      var n = e.key === 'ArrowDown' ? i + 1 : i - 1;
      if (n < 0) q.focus(); else if (n < l.length) l[n].focus();
      e.preventDefault();
    }
    if (e.key === 'Enter' && document.activeElement === q) {
      var premier = document.querySelector('#contenu .feuille'); if (premier && q.value.trim()) premier.click();
    }
  });
  /* boutons latéraux de la souris : 3 = précédent, 4 = suivant */
  function souris(e) {
    if (e.button === 3) { e.preventDefault(); if (e.type === 'mouseup') aller(-1); }
    else if (e.button === 4) { e.preventDefault(); if (e.type === 'mouseup') aller(1); }
  }
  document.addEventListener('mousedown', souris); document.addEventListener('mouseup', souris);

  /* ---------- démarrage ---------- */
  function brancher() {
    $('q').addEventListener('input', render);
    $('effacer').addEventListener('click', function () { $('q').value = ''; render(); $('q').focus(); });
    $('retour').addEventListener('click', function () { aller(-1); });
    $('avant').addEventListener('click', function () { aller(1); });
    $('actualiser').addEventListener('click', function () { rafraichir(true).then(function () { toast('Liste actualisée.'); }); });
    $('replier').addEventListener('click', function () { groupes.familles.forEach(function (f) { ferme[f.cle] = true; }); ecrire(); render(); });
    $('deplier').addEventListener('click', function () { ferme = {}; ecrire(); render(); });
    $('auto').addEventListener('change', function () {
      var v = $('auto').checked;
      adapter.ecrireOuvertureAuto(v).then(function (r) {
        if (r.ok) toast(v ? 'Enregistrez le classeur pour que l’ouverture automatique soit conservée.' : 'Ouverture automatique désactivée. Enregistrez le classeur.');
        else { $('auto').checked = !v; toast('Réglage impossible : ' + (r.message || 'classeur en lecture seule ?')); }
      });
    });
  }

  function demarrer() {
    brancher();
    fetch('config.json', { cache: 'no-cache' }).then(function (r) { return r.ok ? r.json() : {}; }).catch(function () { return {}; })
      .then(function (c) { cfg = c || {}; adapter = demo ? window.DemoAdapter : window.ExcelAdapter; return adapter.init(); })
      .then(function (r) {
        if (!r.ok) {
          etat('Hors Excel', 'ko');
          avis('Ce volet fonctionne dans Excel (' + (r.raison || 'hôte non détecté') + '). Pour un aperçu avec un classeur fictif, ajoutez ?demo=1 à l’adresse.', true);
          $('contenu').appendChild(el('div', 'vide', 'Ouvrez un classeur dans Excel, puis cliquez sur « Navigateur CHTL » dans l’onglet Accueil.'));
          return;
        }
        cleClasseur = adapter.nomClasseur() || 'defaut';
        var s = lire(); ferme = s.ferme || {}; favoris = s.favoris || []; noms = s.noms || {}; recents.charger(s.recents);
        majTitre();
        try { if (Office.context.officeTheme && Office.context.officeTheme.isDarkTheme) document.documentElement.classList.add('sombre'); } catch (e) { /* thème par défaut */ }
        if (adapter.mode === 'demo') {
          avis('Aperçu : classeur FICTIF, aucune donnée hospitalière.');
          var outils = el('div', 'demo-outils');
          [['Renommer un onglet', function () { adapter.simuler.renommer('Suivi 2026', 'Suivi 2026 v2'); }],
           ['Supprimer « Détail A »', function () { adapter.simuler.supprimer('Détail A'); }],
           ['Colorier « Notes » en vert', function () { adapter.simuler.colorier('Notes', '#00B050'); }],
           ['Renommer le classeur', function () { adapter.simuler.renommerClasseur('Autre fichier'); }]].forEach(function (d) {
            var b = el('button', null, d[0]); b.type = 'button'; b.addEventListener('click', d[1]); outils.appendChild(b);
          });
          $('avis').appendChild(outils);
        }
        $('autoLabel').hidden = false; $('auto').checked = adapter.lireOuvertureAuto();
        adapter.abonner({ activation: surActivation, structure: function () { rafraichir(false); } });
        return rafraichir(true).then(function () {
          /* Filet de sécurité : sondage léger si un évènement Excel n'est pas disponible (couleurs d'onglet, noms…). */
          var sec = Math.max(2, +cfg.actualisationSecondes || 5);
          setInterval(function () { if (!document.hidden) rafraichir(false); }, sec * 1000);
          document.addEventListener('visibilitychange', function () { if (!document.hidden) rafraichir(false); });
          window.addEventListener('focus', function () { rafraichir(false); });
        });
      });
  }

  window.CHTL = { demarrer: demarrer };
  document.addEventListener('DOMContentLoaded', demarrer);
})();
