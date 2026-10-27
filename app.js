/* CHTL – Navigateur TMCD : interface du volet. */
(function () {
  'use strict';
  var N = window.NavCore;
  var $ = function (id) { return document.getElementById(id); };
  var demo = /[?&]demo=1\b/.test(location.search);
  var adapter = null, cfg = {}, feuilles = [], groupes = { familles: [], masquees: 0 }, actifId = null, sig = '';
  var histo = N.creerHistorique(50);
  var recents = N.creerRecents(8);
  var ferme = {}, favoris = [], defilerVersActif = false;
  var stockCle = 'chtl.tmcd.v1';

  /* ---------- stockage local (par utilisateur, jamais dans le classeur) ---------- */
  function lire() { try { return JSON.parse(localStorage.getItem(stockCle + ':' + wbCle()) || '{}'); } catch (e) { return {}; } }
  function ecrire() { try { localStorage.setItem(stockCle + ':' + wbCle(), JSON.stringify({ ferme: ferme, favoris: favoris, recents: recents.ids() })); } catch (e) { /* stockage indisponible */ } }

  /* Toute arrivée sur une feuille (clic dans le volet ou action dans Excel) passe ici. */
  function noterVisite(id) {
    histo.visiter(id); recents.visiter(id); actifId = id; defilerVersActif = true;
    var fa = N.famillePourId(groupes, id);
    if (fa && ferme[fa.cle]) delete ferme[fa.cle];
    ecrire();
  }
  function wbCle() { return (adapter && adapter.nomClasseur && adapter.nomClasseur()) || 'defaut'; }

  /* ---------- utilitaires d'affichage ---------- */
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
  function toast(msg) {
    var t = $('toast'); t.textContent = msg; t.hidden = false;
    clearTimeout(toastT); toastT = setTimeout(function () { t.hidden = true; }, 4500);
  }
  function avis(msg, erreur) {
    var a = $('avis'); if (!msg) { a.hidden = true; return; }
    a.textContent = msg; a.className = 'avis' + (erreur ? ' erreur' : ''); a.hidden = false;
  }
  function etat(msg, cls) { $('etat').textContent = msg; $('pastille').className = 'pastille ' + (cls || ''); }
  function existe(id) { return feuilles.some(function (f) { return f.id === id; }); }

  /* ---------- navigation ---------- */
  function ouvrir(f) {
    adapter.activer(f.id).then(function (r) {
      if (r.ok) return;
      if (r.raison === 'absente') { toast('« ' + f.name + ' » n’existe plus (supprimée ou renommée). Liste actualisée.'); }
      else if (r.raison === 'masquee') { toast('« ' + (r.nom || f.name) + ' » est masquée dans le classeur.'); }
      else toast(r.message || 'Activation impossible.');
      rafraichir(true);
    });
  }

  function retour() {
    var id = histo.retour(existe);
    if (id == null) { toast('Aucune feuille précédente.'); majBoutons(); return; }
    adapter.activer(id).then(function (r) {
      if (!r.ok) { toast('La feuille précédente n’est plus accessible.'); rafraichir(true); }
    });
    recents.visiter(id); actifId = id; defilerVersActif = true; ecrire(); render();
  }

  function nomDe(id) { var f = feuilles.filter(function (x) { return x.id === id; })[0]; return f ? f.name : ''; }
  function majBoutons() {
    var p = histo.precedent(existe), b = $('retour');
    b.disabled = p == null;
    b.title = p == null ? 'Aucune feuille précédente' : 'Revenir à « ' + nomDe(p) + ' »';
  }

  /* ---------- rendu ---------- */
  function ligne(f, fam, plages, avecFamille) {
    var li = el('div', 'ligne' + (f.id === actifId ? ' active' : ''));
    var b = el('button', 'feuille'); b.type = 'button'; b.dataset.id = f.id;
    if (f.id === actifId) b.setAttribute('aria-current', 'true');
    b.title = 'Ouvrir la feuille « ' + f.name + ' »';
    var coul = N.couleurOnglet(f.tabColor);
    if (coul) { var pt = el('span', 'pastille-tab'); pt.style.background = coul; pt.setAttribute('aria-hidden', 'true'); b.appendChild(pt); }
    var nom = el('span', 'nom'); surligner(nom, f.name, plages); b.appendChild(nom);
    if (avecFamille && fam) b.appendChild(el('span', 'tag t-' + (fam.couleur === 'bleu' ? 'bleu' : fam.couleur), fam.lettre ? fam.lettre + '. ' + fam.libelle : fam.libelle));
    b.addEventListener('click', function () { ouvrir(f); });
    li.appendChild(b);
    var e = el('button', 'etoile' + (estFav(f) ? ' on' : ''), estFav(f) ? '★' : '☆');
    e.type = 'button'; e.title = estFav(f) ? 'Retirer des favoris' : 'Ajouter aux favoris'; e.setAttribute('aria-label', e.title);
    e.addEventListener('click', function (ev) { ev.stopPropagation(); basculerFav(f); });
    li.appendChild(e);
    return li;
  }

  function estFav(f) { return favoris.some(function (x) { return x.id === f.id; }); }
  function basculerFav(f) {
    if (estFav(f)) favoris = favoris.filter(function (x) { return x.id !== f.id; });
    else favoris.push({ id: f.id, name: f.name });
    ecrire(); render();
  }

  function sectionRaccourcis() {
    var wrap = el('section');
    var defs = N.resoudreRaccourcis(feuilles, cfg.raccourcis, new Date().getFullYear());
    var items = defs.map(function (r) { return { libelle: r.def.libelle, feuille: r.feuille, fav: false }; });
    /* favoris : résolus par id (survit au renommage) puis par nom */
    favoris.forEach(function (fv) {
      var f = feuilles.filter(function (x) { return x.id === fv.id; })[0] || feuilles.filter(function (x) { return x.name === fv.name; })[0];
      if (f && f.visibility === 'Visible') items.push({ libelle: f.name, feuille: f, fav: true });
    });
    if (!items.length) return null;
    wrap.appendChild(el('h2', 'section-titre', 'Raccourcis'));
    var box = el('div', 'raccourcis');
    items.forEach(function (it) {
      var b = el('button', 'puce' + (it.fav ? ' fav' : '') + (it.feuille && it.feuille.id === actifId ? ' actif' : '') + (it.feuille ? '' : ' absent'));
      b.type = 'button';
      b.textContent = (it.fav ? '★ ' : '') + (it.feuille && !it.fav ? it.feuille.name !== it.libelle ? it.libelle + ' · ' + it.feuille.name : it.libelle : it.libelle);
      if (it.feuille) { b.title = 'Ouvrir « ' + it.feuille.name + ' »'; b.addEventListener('click', function () { ouvrir(it.feuille); }); }
      else { b.disabled = true; b.title = 'Feuille introuvable dans ce classeur (renommée ou supprimée).'; }
      box.appendChild(b);
    });
    wrap.appendChild(box);
    return wrap;
  }

  function sectionRecents() {
    var ids = recents.lister(existe, actifId, 5);
    if (!ids.length) return null;
    var wrap = el('section');
    wrap.appendChild(el('h2', 'section-titre', 'Récents'));
    var box = el('div', 'raccourcis');
    ids.forEach(function (id) {
      var f = feuilles.filter(function (x) { return x.id === id; })[0];
      var b = el('button', 'puce recent', f.name); b.type = 'button'; b.title = 'Revenir à « ' + f.name + ' »';
      b.addEventListener('click', function () { ouvrir(f); });
      box.appendChild(b);
    });
    wrap.appendChild(box);
    return wrap;
  }

  function famille(fa) {
    var ouvert = !ferme[fa.cle];
    var box = el('section', 'famille ' + (fa.couleur === 'bleu' ? '' : fa.couleur) + (ouvert ? '' : ' ferme'));
    var tete = el('div', 'fam-tete');
    var bas = el('button', 'fam-bascule'); bas.type = 'button'; bas.setAttribute('aria-expanded', String(ouvert));
    bas.innerHTML = '<svg class="chev" viewBox="0 0 10 10" aria-hidden="true"><path d="M1 3l4 4 4-4" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>';
    if (fa.lettre) bas.appendChild(el('span', 'lettre', fa.lettre));
    bas.appendChild(el('span', 'fam-nom', fa.libelle));
    bas.appendChild(el('span', 'nombre', String(fa.feuilles.length)));
    bas.addEventListener('click', function () { ferme[fa.cle] = ouvert; ecrire(); render(); });
    tete.appendChild(bas);
    if (fa.separateur) {
      var o = el('button', 'ouvrir-sep', '→'); o.type = 'button';
      o.title = 'Ouvrir la feuille « ' + fa.separateur.name + ' »'; o.setAttribute('aria-label', o.title);
      o.addEventListener('click', function () { ouvrir(fa.separateur); });
      tete.appendChild(o);
    }
    box.appendChild(tete);
    var corps = el('div', 'fam-corps');
    fa.feuilles.forEach(function (f) { corps.appendChild(ligne(f, fa, null, false)); });
    box.appendChild(corps);
    return box;
  }

  function render() {
    var c = $('contenu'), q = $('q').value.trim();
    var pos = c.scrollTop;
    c.textContent = '';
    $('effacer').hidden = !q;
    if (q) {
      var res = N.rechercher(groupes, q);
      var box = el('div', 'resultats');
      if (!res.length) c.appendChild(el('div', 'vide', 'Aucune feuille ne correspond à « ' + q + ' ».'));
      else {
        res.forEach(function (r) { box.appendChild(ligne(r.feuille, r.famille, r.plages, true)); });
        c.appendChild(el('h2', 'section-titre', res.length + ' résultat' + (res.length > 1 ? 's' : '')));
        c.appendChild(box);
      }
    } else {
      var s = sectionRaccourcis(); if (s) c.appendChild(s);
      var rc = sectionRecents(); if (rc) c.appendChild(rc);
      c.appendChild(el('h2', 'section-titre', 'Familles'));
      groupes.familles.forEach(function (fa) { c.appendChild(famille(fa)); });
      if (!groupes.familles.length) c.appendChild(el('div', 'vide', 'Aucune feuille visible.'));
      if (groupes.masquees) c.appendChild(el('p', 'note', groupes.masquees + ' feuille' + (groupes.masquees > 1 ? 's masquées ne sont' : ' masquée n’est') + ' pas affichée' + (groupes.masquees > 1 ? 's' : '') + '.'));
    }
    c.scrollTop = pos;
    if (defilerVersActif) {
      defilerVersActif = false;
      var a = c.querySelector('.ligne.active'); if (a && a.scrollIntoView) a.scrollIntoView({ block: 'nearest' });
    }
    var nb = N.compterVisibles(feuilles);
    etat((adapter.mode === 'demo' ? 'Aperçu hors Excel · ' : 'Connecté à Excel · ') + nb + ' feuille' + (nb > 1 ? 's' : '') + ' visible' + (nb > 1 ? 's' : ''), 'ok');
    majBoutons();
  }

  /* ---------- synchronisation avec le classeur ---------- */
  function rafraichir(force) {
    return adapter.lister().then(function (r) {
      var nouvelle = N.signature(r.feuilles);
      var changeActif = r.actifId !== actifId;
      var change = force || nouvelle !== sig || changeActif;
      if (change) { sig = nouvelle; feuilles = r.feuilles; groupes = N.grouper(feuilles, cfg.familles); }
      if (changeActif) noterVisite(r.actifId);
      if (change) render();
    }).catch(function (e) { etat('Erreur de lecture du classeur', 'ko'); avis('Lecture des feuilles impossible : ' + (e && e.message ? e.message : e), true); });
  }

  function surActivation(id) {
    if (id === actifId) return;
    noterVisite(id);
    if (!existe(id)) rafraichir(true); else render();
  }

  /* ---------- clavier ---------- */
  function boutonsNav() { return Array.prototype.slice.call(document.querySelectorAll('#contenu .feuille, #contenu .puce:not(:disabled)')); }
  document.addEventListener('keydown', function (e) {
    var q = $('q');
    if (e.key === 'Escape' && q.value) { q.value = ''; render(); q.focus(); e.preventDefault(); return; }
    if (e.key === '/' && document.activeElement !== q) { q.focus(); q.select(); e.preventDefault(); return; }
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      var l = boutonsNav(); if (!l.length) return;
      var i = l.indexOf(document.activeElement);
      if (document.activeElement === q) i = -1;
      else if (i < 0) return;
      var n = e.key === 'ArrowDown' ? i + 1 : i - 1;
      if (n < 0) { q.focus(); } else if (n < l.length) l[n].focus();
      e.preventDefault();
    }
    if (e.key === 'Enter' && document.activeElement === q) {
      var premier = document.querySelector('#contenu .feuille');
      if (premier && q.value.trim()) premier.click();
    }
  });

  /* ---------- démarrage ---------- */
  function brancher() {
    $('q').addEventListener('input', render);
    $('effacer').addEventListener('click', function () { $('q').value = ''; render(); $('q').focus(); });
    $('retour').addEventListener('click', retour);
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
      .then(function (c) { cfg = c || {}; })
      .then(function () {
        adapter = demo ? window.DemoAdapter : window.ExcelAdapter;
        return adapter.init();
      }).then(function (r) {
        if (!r.ok) {
          etat('Hors Excel', 'ko');
          avis('Ce volet fonctionne dans Excel (' + (r.raison || 'hôte non détecté') + '). Pour un aperçu avec un classeur fictif, ajoutez ?demo=1 à l’adresse.', true);
          $('contenu').appendChild(el('div', 'vide', 'Ouvrez le classeur TMCD dans Excel, puis cliquez sur « CHTL – Pilotage TMCD » dans le ruban.'));
          return;
        }
        var s = lire(); ferme = s.ferme || {}; favoris = s.favoris || []; recents.charger(s.recents);
        try { if (Office.context.officeTheme && Office.context.officeTheme.isDarkTheme) document.documentElement.classList.add('sombre'); } catch (e) { /* thème par défaut */ }
        $('classeur').textContent = adapter.nomClasseur();
        if (adapter.mode === 'demo') {
          avis('Aperçu : classeur FICTIF, aucune donnée hospitalière.');
          var outils = el('div', 'demo-outils');
          [['Renommer « Suivi 2026 »', function () { adapter.simuler.renommer('Suivi 2026', 'Suivi 2026 v2'); }],
           ['Supprimer « Détail A »', function () { adapter.simuler.supprimer('Détail A'); }],
           ['Ajouter une feuille', function () { adapter.simuler.ajouter('Nouvelle feuille', 'Balance'); }]].forEach(function (d) {
            var b = el('button', null, d[0]); b.type = 'button'; b.addEventListener('click', d[1]); outils.appendChild(b);
          });
          $('avis').appendChild(outils);
        }
        $('autoLabel').hidden = false;
        $('auto').checked = adapter.lireOuvertureAuto();
        adapter.abonner({ activation: surActivation, structure: function () { rafraichir(false); } });
        return rafraichir(true).then(function () {
          /* Filet de sécurité : sondage léger (noms/visibilité/position) si un évènement n'est pas disponible. */
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
