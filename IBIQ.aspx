<%@ Page Language="C#" MasterPageFile="~masterurl/default.master" Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<asp:Content ContentPlaceHolderID="PlaceHolderMain" runat="server">

<style>
/* ── VARIABLES ── */
:root {
  --bg:      #f3f4f6;
  --surface: #ffffff;
  --surface2:#e9ebee;
  --border:  #d1d5db;
  --navy:    #1a4f8a;
  --navydk:  #0d2d52;
  --text:    #1f2937;
  --muted:   #6b7280;
  --accent:  #e8a838;
  --ppt:     #c43e1c;
  --word:    #1a56aa;
  --app:     #059669;
  --wiki:    #7c3aed;
  --red:     #dc2626;
}

* { margin:0; padding:0; box-sizing:border-box; }

#ibiq-root {
  font-family: 'Segoe UI', Arial, sans-serif;
  background: var(--bg);
  color: var(--text);
  min-height: 100vh;
}

/* ── HEADER ── */
.ib-header {
  background: var(--navydk);
  padding: 0 2rem;
  display: flex; align-items: center; justify-content: space-between;
  height: 58px;
  border-bottom: 3px solid var(--accent);
}
.ib-logo { font-size: 1.5rem; font-weight: 700; color: var(--accent); letter-spacing: -1px; }
.ib-logo span { color: #fff; }
.ib-team { font-size: .8rem; color: rgba(255,255,255,.6); }
.ib-team strong { color: #fff; }

/* ── ONGLETS ── */
.ib-tabs {
  background: var(--navy);
  padding: 0 2rem;
  display: flex; align-items: flex-end;
  gap: .25rem;
  position: sticky; top: 0; z-index: 100;
  box-shadow: 0 2px 8px rgba(0,0,0,.15);
}
.ib-tab {
  padding: .65rem 1.2rem;
  font-size: .82rem; font-weight: 600;
  color: rgba(255,255,255,.6);
  background: none; border: none;
  border-top: 3px solid transparent;
  border-radius: 4px 4px 0 0;
  cursor: pointer;
  transition: color .15s, background .15s, border-color .15s;
  white-space: nowrap;
}
.ib-tab:hover { color: #fff; background: rgba(255,255,255,.08); }
.ib-tab.active {
  color: var(--navydk);
  background: var(--bg);
  border-top-color: var(--accent);
  font-weight: 700;
}
.ib-tab-sep { width:1px; height:20px; background:rgba(255,255,255,.2); margin: auto .25rem; }

/* ── HERO ── */
.ib-hero {
  background: linear-gradient(135deg, #dbeafe, #eff6ff, #e0e7ff);
  border-bottom: 1px solid var(--border);
  padding: 2rem;
}
.ib-hero-grid {
  max-width: 1100px; margin: 0 auto;
  display: grid; grid-template-columns: 1fr 1fr; gap: 2rem; align-items: start;
}
.ib-hero h1 { font-size: 1.9rem; font-weight: 700; line-height: 1.15; margin-bottom: .4rem; }
.ib-hero h1 span { color: var(--navy); }
.ib-hero p { color: var(--muted); font-size: .88rem; line-height: 1.6; }
.ib-refs { margin-top: .8rem; display:flex; flex-wrap:wrap; gap:.4rem; }
.ib-chip {
  background: var(--surface2); border: 1px solid var(--border);
  padding: 2px 10px; border-radius: 20px; font-size: .73rem; color: var(--muted);
}
.ib-chip a { color: var(--word); text-decoration:none; font-weight:600; }

/* ── ACTUALITÉS ── */
.ib-news {
  background: var(--surface); border: 1px solid var(--border);
  border-radius: 10px; padding: 1rem;
}
.ib-news h3 { font-size:.68rem; font-weight:700; text-transform:uppercase; letter-spacing:2px; color:var(--muted); margin-bottom:.7rem; }
.ib-news-item { display:flex; gap:.7rem; padding:.5rem 0; border-bottom:1px solid var(--border); }
.ib-news-item:last-child { border-bottom:none; }
.ib-news-date { font-size:.68rem; color:var(--navy); font-weight:700; white-space:nowrap; margin-top:2px; }
.ib-news-text { font-size:.78rem; color:var(--muted); line-height:1.4; }

/* ── MAIN ── */
.ib-main { max-width: 1100px; margin: 0 auto; padding: 1.5rem 2rem; }
.ib-panel { display:none; }
.ib-panel.active { display:block; }

/* ── PART HEADER ── */
.ib-part-label {
  display:flex; align-items:center; gap:.8rem;
  padding: .8rem 1.2rem; border-radius: 10px;
  color: #fff; margin-bottom: 1.5rem;
}
.ib-part-label-icon { font-size:1.4rem; }
.ib-part-label-title { font-size:1rem; font-weight:700; }
.ib-part-label-desc { font-size:.78rem; opacity:.8; margin-top:.1rem; }

/* ── LÉGENDE ── */
.ib-legend {
  display:flex; flex-wrap:wrap; gap:.6rem; align-items:center;
  padding:.8rem 1rem; background:var(--surface);
  border:1px solid var(--border); border-radius:8px; margin-bottom:1.5rem;
}
.ib-legend-label { font-size:.73rem; font-weight:700; color:var(--muted); margin-right:.25rem; }
.ib-legend-item { display:flex; align-items:center; gap:.3rem; font-size:.73rem; color:var(--muted); }
.ib-legend-dot { width:8px; height:8px; border-radius:50%; }

/* ── SECTION ── */
.ib-section { margin-bottom:2rem; }
.ib-section-header {
  display:flex; align-items:center; gap:.6rem;
  margin-bottom:1rem; padding-bottom:.7rem;
  border-bottom:1px solid var(--border);
}
.ib-section-icon {
  width:30px; height:30px; border-radius:7px;
  display:flex; align-items:center; justify-content:center; font-size:.9rem;
}
.ib-section-title { font-size:.95rem; font-weight:700; }
.ib-section-refs { margin-left:auto; font-size:.7rem; color:var(--muted); }

/* ── CARD GRID ── */
.ib-grid {
  display:grid;
  grid-template-columns: repeat(auto-fill, minmax(250px,1fr));
  gap:.9rem;
}
.ib-card {
  background:var(--surface); border:1px solid var(--border);
  border-radius:10px; padding:.9rem 1rem;
  display:flex; align-items:flex-start; gap:.8rem;
  text-decoration:none; color:inherit;
  transition:border-color .2s, background .2s, transform .15s;
  position:relative; overflow:hidden;
}
.ib-card:hover {
  border-color: rgba(26,79,138,.4);
  background: #f0f4ff;
  transform: translateY(-2px);
  box-shadow: 0 4px 12px rgba(26,79,138,.1);
}
.ib-card-icon {
  width:36px; height:36px; border-radius:7px;
  display:flex; align-items:center; justify-content:center;
  font-size:1rem; flex-shrink:0; font-weight:700;
}
.ib-icon-ppt  { background:rgba(196,62,28,.1);  color:var(--ppt);  border:1px solid rgba(196,62,28,.2); }
.ib-icon-word { background:rgba(26,86,170,.1);  color:var(--word); border:1px solid rgba(26,86,170,.2); }
.ib-icon-app  { background:rgba(5,150,105,.1);  color:var(--app);  border:1px solid rgba(5,150,105,.2); }
.ib-icon-wiki { background:rgba(124,58,237,.1); color:var(--wiki); border:1px solid rgba(124,58,237,.2); }
.ib-icon-vid  { background:rgba(220,38,38,.1);  color:var(--red);  border:1px solid rgba(220,38,38,.2); }
.ib-card-body { flex:1; min-width:0; }
.ib-card-title { font-size:.85rem; font-weight:600; margin-bottom:.2rem; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.ib-card-desc  { font-size:.73rem; color:var(--muted); line-height:1.4; }
.ib-badge {
  position:absolute; top:.6rem; right:.7rem;
  font-size:.6rem; font-weight:700; text-transform:uppercase;
  letter-spacing:1px; padding:2px 5px; border-radius:4px;
}
.ib-b-ppt  { background:rgba(196,62,28,.1);  color:var(--ppt); }
.ib-b-word { background:rgba(26,86,170,.1);  color:var(--word); }
.ib-b-app  { background:rgba(5,150,105,.1);  color:var(--app); }
.ib-b-wiki { background:rgba(124,58,237,.1); color:var(--wiki); }
.ib-b-vid  { background:rgba(220,38,38,.1);  color:var(--red); }
.ib-b-obs  { background:rgba(100,100,100,.1);color:#9ca3af; }
.ib-card.obsolete { opacity:.4; }
.ib-card.obsolete:hover { opacity:.65; border-color:var(--border); }

/* ── FOOTER ── */
.ib-footer {
  border-top:1px solid var(--border);
  padding:1.2rem 2rem; text-align:center;
  font-size:.73rem; color:var(--muted); margin-top:1rem;
}

@media(max-width:768px){
  .ib-hero-grid { grid-template-columns:1fr; }
  .ib-grid { grid-template-columns:1fr 1fr; }
  .ib-tabs { overflow-x:auto; }
}
@media(max-width:480px){ .ib-grid { grid-template-columns:1fr; } }
</style>

<div id="ibiq-root">

  <!-- HEADER -->
  <div class="ib-header">
    <div class="ib-logo">IBI<span>Q</span></div>
    <div class="ib-team"><strong>Équipe A340</strong> &nbsp;·&nbsp; Qualité des Données</div>
  </div>

  <!-- ONGLETS -->
  <div class="ib-tabs">
    <button class="ib-tab active" onclick="ibiqTab('pod',this)">🗂️ POD</button>
    <div class="ib-tab-sep"></div>
    <button class="ib-tab" onclick="ibiqTab('projet',this)">📁 Projet</button>
    <button class="ib-tab" onclick="ibiqTab('apps',this)">⚙️ Applications</button>
    <button class="ib-tab" onclick="ibiqTab('docs',this)">📚 Documentation</button>
  </div>

  <!-- HERO -->
  <div class="ib-hero">
    <div class="ib-hero-grid">
      <div>
        <h1>Portail <span>IBIQ</span><br>Qualité des données</h1>
        <p>Périmètre : qualité des données, contrôles automatiques et data lineage.</p>
        <div class="ib-refs">
          <span class="ib-chip">Réf. : <a href="#">Q. Zimmermann</a></span>
          <span class="ib-chip"><a href="#">A. Mamecier</a></span>
          <span class="ib-chip"><a href="#">A. Mazier</a></span>
        </div>
      </div>
      <div class="ib-news">
        <h3>📢 Actualités</h3>
        <div class="ib-news-item">
          <span class="ib-news-date">29/08/2024</span>
          <span class="ib-news-text">Initialisation du projet POD</span>
        </div>
        <div class="ib-news-item">
          <span class="ib-news-date">04/06/2024</span>
          <span class="ib-news-text">Nouvelle SFD : Comment remplir la partie F5</span>
        </div>
      </div>
    </div>
  </div>

  <!-- MAIN -->
  <div class="ib-main">

    <!-- LÉGENDE -->
    <div class="ib-legend">
      <span class="ib-legend-label">Types :</span>
      <div class="ib-legend-item"><div class="ib-legend-dot" style="background:var(--ppt)"></div> PowerPoint</div>
      <div class="ib-legend-item"><div class="ib-legend-dot" style="background:var(--word)"></div> Word</div>
      <div class="ib-legend-item"><div class="ib-legend-dot" style="background:var(--app)"></div> Application</div>
      <div class="ib-legend-item"><div class="ib-legend-dot" style="background:var(--wiki)"></div> Wiki</div>
      <div class="ib-legend-item"><div class="ib-legend-dot" style="background:var(--red)"></div> Vidéo</div>
    </div>

    <!-- ══ ONGLET POD ══ -->
    <div class="ib-panel active" id="panel-pod">
      <div class="ib-part-label" style="background:var(--navy)">
        <div class="ib-part-label-icon">🗂️</div>
        <div>
          <div class="ib-part-label-title">POD — Portail Organisation Données</div>
          <div class="ib-part-label-desc">Point d'entrée du projet — toujours accessible</div>
        </div>
      </div>
      <div class="ib-grid">
        <a class="ib-card" href="#">
          <div class="ib-card-icon ib-icon-ppt">P</div>
          <div class="ib-card-body">
            <div class="ib-card-title">Présentation projet</div>
            <div class="ib-card-desc">Les enjeux du projet, son historique et ses objectifs</div>
          </div>
          <span class="ib-badge ib-b-ppt">PPT</span>
        </a>
        <a class="ib-card" href="#">
          <div class="ib-card-icon ib-icon-wiki">W</div>
          <div class="ib-card-body">
            <div class="ib-card-title">POD — Principe et utilisation</div>
            <div class="ib-card-desc">Guide d'utilisation du portail organisation données</div>
          </div>
          <span class="ib-badge ib-b-wiki">WIKI</span>
        </a>
      </div>
    </div>

    <!-- ══ ONGLET PROJET ══ -->
    <div class="ib-panel" id="panel-projet">
      <div class="ib-part-label" style="background:#374151">
        <div class="ib-part-label-icon">📁</div>
        <div>
          <div class="ib-part-label-title">Projet</div>
          <div class="ib-part-label-desc">Feuille de route et organisation du projet QDD</div>
        </div>
      </div>
      <div class="ib-section">
        <div class="ib-section-header">
          <div class="ib-section-icon" style="background:rgba(208,74,2,.12);border:1px solid rgba(208,74,2,.3)">📊</div>
          <div class="ib-section-title">Feuille de route QDD</div>
        </div>
        <div class="ib-grid">
          <a class="ib-card obsolete" href="#">
            <div class="ib-card-icon ib-icon-ppt">P</div>
            <div class="ib-card-body">
              <div class="ib-card-title">Vision Macro</div>
              <div class="ib-card-desc">Vue d'ensemble macro de la feuille de route</div>
            </div>
            <span class="ib-badge ib-b-obs">OBSOLÈTE</span>
          </a>
        </div>
      </div>
    </div>

    <!-- ══ ONGLET APPLICATIONS ══ -->
    <div class="ib-panel" id="panel-apps">
      <div class="ib-part-label" style="background:#065f46">
        <div class="ib-part-label-icon">⚙️</div>
        <div>
          <div class="ib-part-label-title">Applications</div>
          <div class="ib-part-label-desc">Outils de contrôle, revue de code et indicateurs</div>
        </div>
      </div>

      <!-- SIMC -->
      <div class="ib-section">
        <div class="ib-section-header">
          <div class="ib-section-icon" style="background:rgba(16,185,129,.12);border:1px solid rgba(16,185,129,.3)">🏭</div>
          <div class="ib-section-title">Contrôler ses données — SIMC</div>
          <div class="ib-section-refs">Réf. : A. Mazier · A. Mamecier · Y. Mourelon</div>
        </div>
        <div class="ib-grid">
          <a class="ib-card" href="#">
            <div class="ib-card-icon ib-icon-app">⚙</div>
            <div class="ib-card-body">
              <div class="ib-card-title">Rapport BD KPI UoC</div>
              <div class="ib-card-desc">Tableau de bord des indicateurs clés de qualité</div>
            </div>
            <span class="ib-badge ib-b-app">APP</span>
          </a>
          <a class="ib-card" href="#">
            <div class="ib-card-icon ib-icon-app">🏭</div>
            <div class="ib-card-body">
              <div class="ib-card-title">SIMC (Usine à Contrôle)</div>
              <div class="ib-card-desc">Application de contrôle automatique des données</div>
            </div>
            <span class="ib-badge ib-b-app">APP</span>
          </a>
          <a class="ib-card" href="#">
            <div class="ib-card-icon ib-icon-ppt">P</div>
            <div class="ib-card-body">
              <div class="ib-card-title">Présentation Générale</div>
              <div class="ib-card-desc">Synthèse des attendus ACM, blocs développés, interactions projets</div>
            </div>
            <span class="ib-badge ib-b-ppt">PPT</span>
          </a>
          <a class="ib-card" href="#">
            <div class="ib-card-icon ib-icon-word">W</div>
            <div class="ib-card-body">
              <div class="ib-card-title">SFD — Remplir la fonction F5</div>
              <div class="ib-card-desc">Guide pour compléter la spécification fonctionnelle</div>
            </div>
            <span class="ib-badge ib-b-word">WORD</span>
          </a>
        </div>
      </div>

      <!-- MAYA -->
      <div class="ib-section">
        <div class="ib-section-header">
          <div class="ib-section-icon" style="background:rgba(239,68,68,.12);border:1px solid rgba(239,68,68,.3)">🔍</div>
          <div class="ib-section-title">MAYA — Revue de code</div>
        </div>
        <div class="ib-grid">
          <a class="ib-card" href="#">
            <div class="ib-card-icon ib-icon-app">🧑‍💻</div>
            <div class="ib-card-body">
              <div class="ib-card-title">MAYA — Utilitaire de revue de code</div>
              <div class="ib-card-desc">Outil d'analyse et de revue automatisée du code</div>
            </div>
            <span class="ib-badge ib-b-app">APP</span>
          </a>
        </div>
      </div>
    </div>

    <!-- ══ ONGLET DOCUMENTATION ══ -->
    <div class="ib-panel" id="panel-docs">
      <div class="ib-part-label" style="background:#5b21b6">
        <div class="ib-part-label-icon">📚</div>
        <div>
          <div class="ib-part-label-title">Documentation</div>
          <div class="ib-part-label-desc">Modèles de données, dictionnaire META et lignage</div>
        </div>
      </div>

      <!-- MPD -->
      <div class="ib-section">
        <div class="ib-section-header">
          <div class="ib-section-icon" style="background:rgba(139,92,246,.12);border:1px solid rgba(139,92,246,.3)">📐</div>
          <div class="ib-section-title">Documenter son modèle de données</div>
        </div>
        <div class="ib-grid">
          <a class="ib-card" href="#">
            <div class="ib-card-icon ib-icon-wiki">W</div>
            <div class="ib-card-body">
              <div class="ib-card-title">Créer son MPD depuis les tables</div>
              <div class="ib-card-desc">Procédure de création du modèle physique de données</div>
            </div>
            <span class="ib-badge ib-b-wiki">WIKI</span>
          </a>
          <a class="ib-card" href="#">
            <div class="ib-card-icon ib-icon-vid">▶</div>
            <div class="ib-card-body">
              <div class="ib-card-title">Tutoriel vidéo — Projet PowerDesigner</div>
              <div class="ib-card-desc">Prise en main et utilisation de PowerDesigner</div>
            </div>
            <span class="ib-badge ib-b-vid">VIDÉO</span>
          </a>
        </div>
      </div>

      <!-- META -->
      <div class="ib-section">
        <div class="ib-section-header">
          <div class="ib-section-icon" style="background:rgba(232,168,56,.12);border:1px solid rgba(232,168,56,.3)">🔖</div>
          <div class="ib-section-title">META — Dictionnaire de données</div>
          <div class="ib-section-refs">Réf. : A. Mazier · A. Mamecier · H. Charlot</div>
        </div>
        <div class="ib-grid">
          <a class="ib-card" href="#">
            <div class="ib-card-icon ib-icon-wiki">W</div>
            <div class="ib-card-body">
              <div class="ib-card-title">MPD vers META</div>
              <div class="ib-card-desc">Procédure d'alimentation du dictionnaire depuis le MPD</div>
            </div>
            <span class="ib-badge ib-b-wiki">WIKI</span>
          </a>
          <a class="ib-card" href="#">
            <div class="ib-card-icon ib-icon-wiki">W</div>
            <div class="ib-card-body">
              <div class="ib-card-title">Utilisation du dictionnaire META</div>
              <div class="ib-card-desc">Guide d'utilisation et de consultation du dictionnaire</div>
            </div>
            <span class="ib-badge ib-b-wiki">WIKI</span>
          </a>
        </div>
      </div>

      <!-- LIGNAGE -->
      <div class="ib-section">
        <div class="ib-section-header">
          <div class="ib-section-icon" style="background:rgba(59,130,246,.12);border:1px solid rgba(59,130,246,.3)">🔗</div>
          <div class="ib-section-title">Lignage de données</div>
        </div>
        <div class="ib-grid">
          <a class="ib-card" href="#">
            <div class="ib-card-icon ib-icon-wiki">W</div>
            <div class="ib-card-body">
              <div class="ib-card-title">Lignage de données</div>
              <div class="ib-card-desc">Documentation et traçabilité du cycle de vie des données</div>
            </div>
            <span class="ib-badge ib-b-wiki">WIKI</span>
          </a>
        </div>
      </div>
    </div>

  </div><!-- /ib-main -->

  <div class="ib-footer">
    IBIQ · Équipe A340 · Portail Qualité des Données — Dernière mise à jour 29/08/2024
  </div>

</div><!-- /ibiq-root -->

<script type="text/javascript">
function ibiqTab(name, btn) {
  // Désactiver tous les onglets et panneaux
  var tabs   = document.querySelectorAll('.ib-tab');
  var panels = document.querySelectorAll('.ib-panel');
  for (var i = 0; i < tabs.length; i++)   tabs[i].classList.remove('active');
  for (var i = 0; i < panels.length; i++) panels[i].classList.remove('active');
  // Activer l'onglet cliqué
  btn.classList.add('active');
  var panel = document.getElementById('panel-' + name);
  if (panel) panel.classList.add('active');
}
</script>

</asp:Content>
