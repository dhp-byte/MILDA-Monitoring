################################################################################
# TABLEAU DE BORD AVANCÉ - Monitorage externe MILDA
# Version Corrigée - Compatible Excel et KoBo
################################################################################

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
import scipy
from streamlit_folium import st_folium
import folium
import requests
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
import io
import zipfile
import re
import math
import json
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import warnings
warnings.filterwarnings('ignore')

# Bibliothèques avancées
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from scipy import stats
    from sklearn.preprocessing import StandardScaler
    STATS_AVAILABLE = True
except ImportError:
    STATS_AVAILABLE = False

################################################################################
# CONFIGURATION GLOBALE
################################################################################

# Configuration de la page
st.set_page_config(
    page_title="MILDA Dashboard",
    page_icon="🦟",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': 'https://www.example.com/help',
        'Report a bug': "https://www.example.com/bug",
        'About': "# MILDA Dashboard v2.0\nTableau de bord pour le monitorage de la distribution des moustiquaires au Tchad"
    }
)

# Thème et styles personnalisés
CUSTOM_CSS = """
<style>
    /* En-tête principal */
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    /* Cartes KPI */
    .kpi-card {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border-left: 4px solid #667eea;
        margin-bottom: 1rem;
    }
    
    .kpi-value {
        font-size: 2.5rem;
        font-weight: bold;
        color: #667eea;
        margin: 0;
    }
    
    .kpi-label {
        font-size: 0.9rem;
        color: #666;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    .kpi-trend {
        font-size: 0.8rem;
        margin-top: 0.5rem;
    }
    
    .trend-up { color: #10b981; }
    .trend-down { color: #ef4444; }
    .trend-neutral { color: #6b7280; }
    
    /* Alertes */
    .alert-box {
        padding: 1rem;
        border-radius: 8px;
        margin: 1rem 0;
    }
    
    .alert-success { background-color: #d1fae5; border-left: 4px solid #10b981; }
    .alert-warning { background-color: #fef3c7; border-left: 4px solid #f59e0b; }
    .alert-danger { background-color: #fee2e2; border-left: 4px solid #ef4444; }
    .alert-info { background-color: #dbeafe; border-left: 4px solid #3b82f6; }
    
    /* Filtres */
    .filter-section {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 10px;
        margin-bottom: 1.5rem;
    }
    
    /* Tables améliorées */
    .dataframe {
        font-size: 0.9rem !important;
    }
    
    /* Badges */
    .badge {
        display: inline-block;
        padding: 0.25rem 0.75rem;
        border-radius: 12px;
        font-size: 0.8rem;
        font-weight: 600;
    }
    
    .badge-success { background-color: #d1fae5; color: #065f46; }
    .badge-warning { background-color: #fef3c7; color: #92400e; }
    .badge-danger { background-color: #fee2e2; color: #991b1b; }
    
    /* Animation de chargement */
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.5; }
    }
    
    .loading-pulse {
        animation: pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite;
    }
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

################################################################################
# CLASSES ET STRUCTURES DE DONNÉES
################################################################################

class DataProcessor:
    """Classe pour le traitement avancé des données"""
    
    @staticmethod
    def normalize_yes_no(value) -> Optional[str]:
        """Normalise les réponses Oui/Non avec gestion robuste"""
        if pd.isna(value):
            return None
        
        value_str = str(value).strip().lower()
        yes_values = ['oui', 'yes', 'y', '1', 'true', 'o', '1.0']
        no_values = ['non', 'no', 'n', '0', 'false', '0.0']
        
        if value_str in yes_values:
            return 'Oui'
        elif value_str in no_values:
            return 'Non'
        return None
    
    @staticmethod
    def calculate_expected_milda(n_persons: float) -> int:
        """Calcule le nombre de MILDA attendues (1 pour 2 personnes)"""
        try:
            if pd.isna(n_persons) or n_persons <= 0:
                return 0
            return math.ceil(float(n_persons) / 2)
        except:
            return 0
    
    @staticmethod
    def clean_column_names(df: pd.DataFrame) -> pd.DataFrame:
        """Nettoie et normalise les noms de colonnes"""
        df = df.copy()
        df.columns = df.columns.str.strip()
        return df
    
    @staticmethod
    def detect_outliers(series: pd.Series, method='iqr', threshold=1.5) -> pd.Series:
        """Détecte les valeurs aberrantes"""
        if method == 'iqr':
            Q1 = series.quantile(0.25)
            Q3 = series.quantile(0.75)
            IQR = Q3 - Q1
            lower_bound = Q1 - threshold * IQR
            upper_bound = Q3 + threshold * IQR
            return (series < lower_bound) | (series > upper_bound)
        elif method == 'zscore' and STATS_AVAILABLE:
            z_scores = np.abs(stats.zscore(series.dropna()))
            return z_scores > threshold
        return pd.Series([False] * len(series))


class MetricsCalculator:
    """Classe pour calculer les métriques et indicateurs"""
    
    @staticmethod
    def calculate_coverage_metrics(df: pd.DataFrame) -> Dict:
        """Calcule les métriques de couverture"""
        total_households = len(df)
        
        if total_households == 0:
            return {
                'total_menages': 0,
                'menages_servis': 0,
                'menages_correct': 0,
                'menages_marques': 0,
                'menages_informes': 0,
                'pct_servis': 0,
                'pct_correct': 0,
                'pct_marques': 0,
                'pct_informes': 0
            }
        
        served = (df['indic_servi'] == 1).sum() if 'indic_servi' in df.columns else 0
        correctly_served = (df['indic_correct'] == 1).sum() if 'indic_correct' in df.columns else 0
        marked = (df['indic_marque'] == 1).sum() if 'indic_marque' in df.columns else 0
        informed = (df['indic_info'] == 1).sum() if 'indic_info' in df.columns else 0
        
        return {
            'total_menages': total_households,
            'menages_servis': served,
            'menages_correct': correctly_served,
            'menages_marques': marked,
            'menages_informes': informed,
            'pct_servis': round(100 * served / total_households, 2) if total_households > 0 else 0,
            'pct_correct': round(100 * correctly_served / served, 2) if served > 0 else 0,
            'pct_marques': round(100 * marked / served, 2) if served > 0 else 0,
            'pct_informes': round(100 * informed / total_households, 2) if total_households > 0 else 0
        }
    
    @staticmethod
    def calculate_distribution_accuracy(df: pd.DataFrame) -> Dict:
        """Calcule la précision de la distribution"""
        df_served = df[df['menage_servi'] == 'Oui'].copy()
        
        if len(df_served) == 0:
            return {'precision': 0, 'sur_distribution': 0, 'sous_distribution': 0, 'ecart_moyen': 0}
        
        if 'nb_milda_recues' not in df_served.columns or 'nb_milda_attendues' not in df_served.columns:
            return {'precision': 0, 'sur_distribution': 0, 'sous_distribution': 0, 'ecart_moyen': 0}
        
        df_served['ecart'] = df_served['nb_milda_recues'] - df_served['nb_milda_attendues']
        
        correct = (df_served['ecart'] == 0).sum()
        over = (df_served['ecart'] > 0).sum()
        under = (df_served['ecart'] < 0).sum()
        
        return {
            'precision': round(100 * correct / len(df_served), 2),
            'sur_distribution': round(100 * over / len(df_served), 2),
            'sous_distribution': round(100 * under / len(df_served), 2),
            'ecart_moyen': round(df_served['ecart'].mean(), 2)
        }
    
    @staticmethod
    def calculate_quality_score(metrics: Dict) -> float:
        """Calcule un score de qualité global (0-100)"""
        weights = {
            'pct_servis': 0.25,
            'pct_correct': 0.30,
            'pct_marques': 0.20,
            'pct_informes': 0.25
        }
        
        score = sum(metrics.get(k, 0) * w for k, w in weights.items())
        return round(score, 2)


class VisualizationEngine:
    """Classe pour créer des visualisations avancées"""
    
    COLOR_PALETTE = {
        'primary': '#667eea',
        'secondary': '#764ba2',
        'success': '#10b981',
        'warning': '#f59e0b',
        'danger': '#ef4444',
        'info': '#3b82f6'
    }
    
    @classmethod
    def create_kpi_gauge(cls, value: float, title: str, max_value: float = 100,
                        threshold_good: float = 80, threshold_medium: float = 60) -> go.Figure:
        """Crée un gauge KPI interactif"""
        
        # Déterminer la couleur selon les seuils
        if value >= threshold_good:
            color = cls.COLOR_PALETTE['success']
        elif value >= threshold_medium:
            color = cls.COLOR_PALETTE['warning']
        else:
            color = cls.COLOR_PALETTE['danger']
        
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=value,
            title={'text': title, 'font': {'size': 16}},
            delta={'reference': threshold_good, 'increasing': {'color': cls.COLOR_PALETTE['success']}},
            gauge={
                'axis': {'range': [None, max_value], 'tickwidth': 1},
                'bar': {'color': color},
                'bgcolor': "white",
                'borderwidth': 2,
                'bordercolor': "gray",
                'steps': [
                    {'range': [0, threshold_medium], 'color': '#fee2e2'},
                    {'range': [threshold_medium, threshold_good], 'color': '#fef3c7'},
                    {'range': [threshold_good, max_value], 'color': '#d1fae5'}
                ],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': threshold_good
                }
            }
        ))
        
        fig.update_layout(
            height=250,
            margin=dict(l=20, r=20, t=50, b=20),
            paper_bgcolor="white",
            font={'color': "#333", 'family': "Arial"}
        )
        
        return fig
    
    @classmethod
    def create_stacked_bar_chart(cls, df: pd.DataFrame, x_col: str, 
                                 y_cols: List[str], title: str) -> go.Figure:
        """Crée un graphique à barres empilées avec annotations"""
        
        fig = go.Figure()
        
        colors = [cls.COLOR_PALETTE['primary'], cls.COLOR_PALETTE['success'], 
                 cls.COLOR_PALETTE['warning'], cls.COLOR_PALETTE['info']]
        
        for idx, col in enumerate(y_cols):
            if col in df.columns:
                fig.add_trace(go.Bar(
                    name=col.replace('pct_', '').title(),
                    x=df[x_col],
                    y=df[col],
                    marker_color=colors[idx % len(colors)],
                    text=df[col].apply(lambda x: f'{x:.1f}%' if pd.notna(x) else ''),
                    textposition='inside',
                    textfont=dict(color='white', size=12),
                    hovertemplate='<b>%{x}</b><br>' + col + ': %{y:.1f}%<extra></extra>'
                ))
        
        fig.update_layout(
            title={
                'text': f'<b>{title}</b>',
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 20}
            },
            barmode='stack',
            xaxis_title="",
            yaxis_title="Pourcentage (%)",
            height=500,
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="center",
                x=0.5
            ),
            template="plotly_white",
            hovermode='x unified'
        )
        
        return fig
    
    @classmethod
    def create_comparison_chart(cls, df: pd.DataFrame, categories: str, 
                               metric: str, title: str) -> go.Figure:
        """Crée un graphique de comparaison horizontal avec gradient"""
        
        if len(df) == 0 or metric not in df.columns:
            fig = go.Figure()
            fig.add_annotation(text="Pas de données disponibles", 
                             xref="paper", yref="paper",
                             x=0.5, y=0.5, showarrow=False)
            return fig
        
        df_sorted = df.sort_values(metric)
        
        # Créer un gradient de couleurs
        colors = [cls.COLOR_PALETTE['danger'] if v < 60 else 
                 cls.COLOR_PALETTE['warning'] if v < 80 else 
                 cls.COLOR_PALETTE['success'] for v in df_sorted[metric]]
        
        fig = go.Figure(go.Bar(
            y=df_sorted[categories],
            x=df_sorted[metric],
            orientation='h',
            marker=dict(
                color=colors,
                line=dict(color='rgba(0,0,0,0.3)', width=1)
            ),
            text=df_sorted[metric].apply(lambda x: f'{x:.1f}%'),
            textposition='outside',
            textfont=dict(size=12),
            hovertemplate='<b>%{y}</b><br>Valeur: %{x:.1f}%<extra></extra>'
        ))
        
        fig.update_layout(
            title={'text': f'<b>{title}</b>', 'x': 0.5, 'xanchor': 'center'},
            xaxis=dict(title="Pourcentage (%)", range=[0, 105]),
            yaxis=dict(title=""),
            height=max(400, len(df) * 40),
            template="plotly_white",
            showlegend=False
        )
        
        # Ajouter une ligne de référence à 80%
        fig.add_vline(x=80, line_dash="dash", line_color="red", 
                     annotation_text="Objectif 80%", annotation_position="top")
        
        return fig


################################################################################
# FONCTIONS DE TRAITEMENT DES DONNÉES
################################################################################

@st.cache_data(ttl=3600, show_spinner=False)
def load_and_process_data(uploaded_file, sheet_name: str = None) -> Tuple[pd.DataFrame, Dict]:
    """Charge et traite les données avec mise en cache et mapping robuste"""
    
    try:
        # Lecture du fichier
        if sheet_name:
            try:
                data = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            except:
                data = pd.read_excel(uploaded_file, sheet_name=0)
        else:
            data = pd.read_excel(uploaded_file, sheet_name=0)
        
        # Nettoyer les noms de colonnes
        data = DataProcessor.clean_column_names(data)
        
        # Mapping complet et robuste des colonnes
        # IMPORTANT : Les colonnes KoBo ont le préfixe 'gr_1/'
        column_mapping = {
            'province': ['province', 'Province', 'S0Q04'],
            'district': ['district', 'district sanitaire', 'District sanitaire de :', 'S0Q05'],
            'centre_sante': ['centre_sante', 'centre de santé', 'Centre de santé', 'S0Q06'],
            'village': ['village', 'Village/Avenue/Quartier', 'S0Q07'],
            'date_enquete': ['date_enquete', 'date_enquête', 'Date enquête', 'Date', 'Date de l'enquête', 'S0Q01'],
            'heure_interview': ['heure_interview', 'Heure', 'time', 'heure', 'end'],
            'agent_name': ['agent_name', "Nom de l'enquêteur", 'Enquêteur', 'Username', 'S0Q03'],
            
            # CORRECTION CRITIQUE : mapping correct pour KoBo
            'menage_servi': [
                'menage_servi',
                'Est-ce que le ménage a-t-il été servi en MILDA lors de la campagne de distribution de masse ?',
                'gr_1/S1Q17'  # Colonne KoBo correcte
            ],
            'nb_personnes': [
                'nb_personnes',
                'Nombre des personnes qui habitent dans le ménage',
                'gr_1/S1Q19'  # Colonne KoBo correcte
            ],
            'nb_milda_recues': [
                'nb_milda_recues',
                'Combien de MILDA avez-vous reçues ?',
                'gr_1/S1Q20'  # Colonne KoBo correcte
            ],
            'verif_cle': [
                'verif_cle',
                'gr_1/verif_cle'  # Colonne KoBo correcte
            ],
            'menage_marque': [
                'menage_marque',
                'Est-ce que le ménage a  été marqué comme un ménage ayant reçu de MILDA?',
                'gr_1/S1Q22'  # Colonne KoBo correcte
            ],
            'sensibilise': [
                'sensibilise',
                'Avez-vous été sensibilisé sur l'utilisation correcte du MILDA par les relais communautaires ?',
                'gr_1/S1Q23'  # Colonne KoBo correcte
            ],
            'information': [
                'information',
                'Étiez-vous informé qu'il y aurait une campagne de distribution de moustiquaires et que des agents visiteraient les ménages ?',
                'gr_1/information'  # Colonne KoBo correcte
            ],
            'latitude': ['latitude', '_LES COORDONNEES GEOGRAPHIQUES_latitude', 'geo_location_latitude'],
            'longitude': ['longitude', '_LES COORDONNEES GEOGRAPHIQUES_longitude', 'geo_location_longitude']
        }
        
        # Appliquer le mapping
        rename_dict = {}
        for target, sources in column_mapping.items():
            for source in sources:
                if source in data.columns:
                    rename_dict[source] = target
                    break
        
        data = data.rename(columns=rename_dict)
        
        # LOG : Afficher les colonnes trouvées pour debug
        st.sidebar.info(f"✅ Colonnes détectées:\n" + "\n".join([f"• {col}" for col in data.columns[:10]]))
        
        # Normalisation des colonnes Oui/Non
        yes_no_cols = ['menage_servi', 'verif_cle', 'menage_marque', 'sensibilise', 'information']
        for col in yes_no_cols:
            if col in data.columns:
                data[col] = data[col].apply(DataProcessor.normalize_yes_no)
        
        # Conversion des valeurs numériques
        if 'nb_personnes' in data.columns:
            data['nb_personnes'] = pd.to_numeric(data['nb_personnes'], errors='coerce')
        if 'nb_milda_recues' in data.columns:
            data['nb_milda_recues'] = pd.to_numeric(data['nb_milda_recues'], errors='coerce')
        
        # Calcul des MILDA attendues
        if 'nb_personnes' in data.columns:
            data['nb_milda_attendues'] = data['nb_personnes'].apply(DataProcessor.calculate_expected_milda)
        else:
            data['nb_milda_attendues'] = 0
        
        # Calcul de l'écart de distribution
        if 'nb_milda_attendues' in data.columns and 'nb_milda_recues' in data.columns:
            data['ecart_distribution'] = data['nb_milda_recues'] - data['nb_milda_attendues']
        else:
            data['ecart_distribution'] = 0
        
        # Indicateurs binaires (avec vérification d'existence)
        data['indic_servi'] = (data['menage_servi'] == 'Oui').astype(int) if 'menage_servi' in data.columns else 0
        
        if 'menage_servi' in data.columns and 'verif_cle' in data.columns:
            data['indic_correct'] = ((data['menage_servi'] == 'Oui') & 
                                     (data['verif_cle'].str.contains('Oui', na=False))).astype(int)
        else:
            data['indic_correct'] = 0
        
        if 'menage_servi' in data.columns and 'menage_marque' in data.columns:
            data['indic_marque'] = ((data['menage_servi'] == 'Oui') & 
                                     (data['menage_marque'] == 'Oui')).astype(int)
        else:
            data['indic_marque'] = 0
        
        data['indic_info'] = (data['sensibilise'] == 'Oui').astype(int) if 'sensibilise' in data.columns else 0
        
        # Conversion des dates
        if 'date_enquete' in data.columns:
            data['date_enquete'] = pd.to_datetime(data['date_enquete'], errors='coerce')
        
        # Statistiques de base
        stats = {
            'total_rows': len(data),
            'total_provinces': data['province'].nunique() if 'province' in data.columns else 0,
            'total_districts': data['district'].nunique() if 'district' in data.columns else 0,
            'date_range': (
                data['date_enquete'].min().strftime('%Y-%m-%d') if 'date_enquete' in data.columns and data['date_enquete'].notna().any() else 'N/A',
                data['date_enquete'].max().strftime('%Y-%m-%d') if 'date_enquete' in data.columns and data['date_enquete'].notna().any() else 'N/A'
            ),
            'colonnes_trouvees': list(data.columns)
        }
        
        return data, stats
        
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement des données : {str(e)}")
        st.error(f"Type d'erreur : {type(e).__name__}")
        import traceback
        st.code(traceback.format_exc())
        return pd.DataFrame(), {}


@st.cache_data(show_spinner=False)
def generate_analysis_tables(data: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Génère les tableaux d'analyse"""
    
    tables = {}
    
    try:
        # Table 0: Résumé par province
        if 'province' in data.columns:
            tables['resume_province'] = data.groupby('province').agg(
                menages_total=('province', 'count'),
                menages_servis=('indic_servi', 'sum'),
                menages_correct=('indic_correct', 'sum'),
                menages_marques=('indic_marque', 'sum'),
                menages_informes=('indic_info', 'sum'),
                pct_servis=('indic_servi', lambda x: round(100 * x.mean(), 1)),
                pct_correct=('indic_correct', lambda x: round(100 * x.mean(), 1)),
                pct_marques=('indic_marque', lambda x: round(100 * x.mean(), 1)),
                pct_informes=('indic_info', lambda x: round(100 * x.mean(), 1))
            ).reset_index()
        
        # Table 1: Détail par district
        if 'province' in data.columns and 'district' in data.columns:
            tables['detail_district'] = data.groupby(['province', 'district']).agg(
                menages_total=('district', 'count'),
                pct_servis=('indic_servi', lambda x: round(100 * x.mean(), 1)),
                pct_correct=('indic_correct', lambda x: round(100 * x.mean(), 1)),
                pct_marques=('indic_marque', lambda x: round(100 * x.mean(), 1)),
                pct_informes=('indic_info', lambda x: round(100 * x.mean(), 1))
            ).reset_index()
        
        # Table 2: Analyse de la distribution
        if 'ecart_distribution' in data.columns:
            distribution_df = data[data['menage_servi'] == 'Oui'].copy()
            if len(distribution_df) > 0:
                tables['analyse_distribution'] = distribution_df.groupby('province').agg(
                    total=('province', 'count'),
                    distribution_exacte=('ecart_distribution', lambda x: round(100 * (x == 0).mean(), 1)),
                    sur_distribution=('ecart_distribution', lambda x: round(100 * (x > 0).mean(), 1)),
                    sous_distribution=('ecart_distribution', lambda x: round(100 * (x < 0).mean(), 1)),
                    ecart_moyen=('ecart_distribution', lambda x: round(x.mean(), 2))
                ).reset_index()
        
        # Table 3: Performance par enquêteur
        if 'agent_name' in data.columns:
            tables['performance_enqueteur'] = data.groupby('agent_name').agg(
                nombre_enquetes=('agent_name', 'count'),
                pct_servis=('indic_servi', lambda x: round(100 * x.mean(), 1)),
                pct_correct=('indic_correct', lambda x: round(100 * x.mean(), 1)),
                qualite_score=('indic_correct', lambda x: round(100 * x.mean(), 1))
            ).reset_index().sort_values('qualite_score', ascending=False)
    
    except Exception as e:
        st.warning(f"⚠️ Erreur lors de la génération des tableaux : {str(e)}")
    
    return tables


################################################################################
# INTERFACE UTILISATEUR - FONCTIONS DE RENDU
################################################################################

def render_header():
    """Affiche l'en-tête principal"""
    st.markdown("""
        <div class="main-header">
            <h1>🦟 MILDA Dashboard</h1>
            <p style="font-size: 1.2rem; margin-top: 0.5rem;">
                Système de monitorage et d'analyse de la distribution des moustiquaires au Tchad 2026
            </p>
        </div>
    """, unsafe_allow_html=True)


def render_kpi_cards(metrics: Dict):
    """Affiche les cartes KPI principales"""
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
            <div class="kpi-card">
                <p class="kpi-label">Ménages Servis</p>
                <p class="kpi-value">{metrics.get('pct_servis', 0):.1f}%</p>
                <p class="kpi-trend trend-{'up' if metrics.get('pct_servis', 0) >= 80 else 'down'}">
                    {'✓ Objectif atteint' if metrics.get('pct_servis', 0) >= 80 else '⚠ Sous objectif'}
                </p>
            </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
            <div class="kpi-card">
                <p class="kpi-label">Distribution Correcte</p>
                <p class="kpi-value">{metrics.get('pct_correct', 0):.1f}%</p>
                <p class="kpi-trend trend-{'up' if metrics.get('pct_correct', 0) >= 80 else 'down'}">
                    {'✓ Objectif atteint' if metrics.get('pct_correct', 0) >= 80 else '⚠ Sous objectif'}
                </p>
            </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
            <div class="kpi-card">
                <p class="kpi-label">Ménages Marqués</p>
                <p class="kpi-value">{metrics.get('pct_marques', 0):.1f}%</p>
                <p class="kpi-trend trend-{'up' if metrics.get('pct_marques', 0) >= 80 else 'down'}">
                    {'✓ Objectif atteint' if metrics.get('pct_marques', 0) >= 80 else '⚠ Sous objectif'}
                </p>
            </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
            <div class="kpi-card">
                <p class="kpi-label">Ménages Informés</p>
                <p class="kpi-value">{metrics.get('pct_informes', 0):.1f}%</p>
                <p class="kpi-trend trend-{'up' if metrics.get('pct_informes', 0) >= 80 else 'down'}">
                    {'✓ Objectif atteint' if metrics.get('pct_informes', 0) >= 80 else '⚠ Sous objectif'}
                </p>
            </div>
        """, unsafe_allow_html=True)


def render_alerts(metrics: Dict):
    """Affiche les alertes basées sur les seuils"""
    
    alerts = []
    
    if metrics.get('pct_servis', 0) < 70:
        alerts.append(('danger', f"Taux de ménages servis critique: {metrics['pct_servis']:.1f}% (objectif: 80%)"))
    elif metrics.get('pct_servis', 0) < 80:
        alerts.append(('warning', f"Taux de ménages servis sous l'objectif: {metrics['pct_servis']:.1f}% (objectif: 80%)"))
    else:
        alerts.append(('success', f"Excellent taux de ménages servis: {metrics['pct_servis']:.1f}%"))
    
    if metrics.get('pct_correct', 0) < 70:
        alerts.append(('danger', f"Précision de distribution critique: {metrics['pct_correct']:.1f}%"))
    
    if metrics.get('pct_informes', 0) < 60:
        alerts.append(('warning', "Sensibilisation insuffisante sur l'utilisation des MILDA"))
    
    for alert_type, message in alerts:
        st.markdown(f"""
            <div class="alert-box alert-{alert_type}">
                <strong>{'🔴' if alert_type == 'danger' else '⚠️' if alert_type == 'warning' else '✅'}</strong> {message}
            </div>
        """, unsafe_allow_html=True)


################################################################################
# PAGES DU DASHBOARD
################################################################################

def page_dashboard(data: pd.DataFrame, tables: Dict[str, pd.DataFrame]):
    """Page principale du dashboard"""
    
    st.markdown("## 📊 Vue d'ensemble")
    
    # Calcul des métriques
    metrics = MetricsCalculator.calculate_coverage_metrics(data)
    quality_score = MetricsCalculator.calculate_quality_score(metrics)
    
    # KPIs principaux
    render_kpi_cards(metrics)
    
    st.markdown("---")
    
    # Alertes
    with st.expander("🔔 Alertes et Notifications", expanded=True):
        render_alerts(metrics)
    
    st.markdown("---")
    
    # Graphiques principaux
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📈 Indicateurs par Province")
        if 'resume_province' in tables and len(tables['resume_province']) > 0:
            fig = VisualizationEngine.create_comparison_chart(
                tables['resume_province'],
                'province',
                'pct_servis',
                'Taux de couverture par province'
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Données par province non disponibles")
    
    with col2:
        st.markdown("### 🎯 Score de qualité global")
        fig = VisualizationEngine.create_kpi_gauge(
            quality_score,
            "Score de qualité",
            max_value=100,
            threshold_good=80,
            threshold_medium=60
        )
        st.plotly_chart(fig, use_container_width=True)
    
    # Graphiques secondaires
    if 'resume_province' in tables and len(tables['resume_province']) > 0:
        st.markdown("### 📊 Comparaison des indicateurs")
        fig = VisualizationEngine.create_stacked_bar_chart(
            tables['resume_province'],
            'province',
            ['pct_servis', 'pct_correct', 'pct_marques', 'pct_informes'],
            'Indicateurs de qualité par province'
        )
        st.plotly_chart(fig, use_container_width=True)


def page_data_explorer(data: pd.DataFrame):
    """Page d'exploration des données"""
    
    st.markdown("## 🔍 Explorateur de données")
    
    st.markdown("### 📋 Données brutes")
    st.dataframe(data, use_container_width=True, height=600)
    
    st.markdown("### 📊 Statistiques descriptives")
    st.dataframe(data.describe(), use_container_width=True)
    
    # Export
    st.markdown("### 💾 Exporter les données")
    col1, col2 = st.columns(2)
    
    with col1:
        csv = data.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Télécharger CSV",
            data=csv,
            file_name=f"milda_data_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    
    with col2:
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            data.to_excel(writer, sheet_name='Données', index=False)
        excel_buffer.seek(0)
        
        st.download_button(
            label="📥 Télécharger Excel",
            data=excel_buffer,
            file_name=f"milda_data_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


################################################################################
# APPLICATION PRINCIPALE
################################################################################

def main():
    """Fonction principale de l'application"""
    
    # En-tête
    render_header()
    
    # Sidebar
    with st.sidebar:
        st.markdown("## ⚙️ Configuration")
        
        uploaded_file = st.file_uploader(
            "📂 Charger un fichier de données",
            type=['xlsx', 'xls'],
            help="Formats supportés: Excel (.xlsx, .xls)"
        )
        
        if uploaded_file:
            # Options de chargement
            sheet_name = st.text_input("📄 Nom de la feuille (optionnel)", value="")
            
            if st.button("🔄 Charger les données"):
                with st.spinner("Chargement en cours..."):
                    data, stats = load_and_process_data(
                        uploaded_file, 
                        sheet_name if sheet_name else None
                    )
                    
                    if not data.empty:
                        st.session_state['data'] = data
                        st.session_state['stats'] = stats
                        st.success("✅ Données chargées avec succès!")
                        
                        # Afficher les stats
                        st.metric("Lignes", stats['total_rows'])
                        st.metric("Provinces", stats['total_provinces'])
                        st.metric("Districts", stats['total_districts'])
                    else:
                        st.error("❌ Erreur lors du chargement")
    
    # Contenu principal
    if 'data' in st.session_state and not st.session_state['data'].empty:
        data = st.session_state['data']
        
        # Générer les tableaux d'analyse
        tables = generate_analysis_tables(data)
        
        # Navigation
        page = st.sidebar.radio(
            "📍 Navigation",
            ["Dashboard", "Explorateur de données"]
        )
        
        if page == "Dashboard":
            page_dashboard(data, tables)
        elif page == "Explorateur de données":
            page_data_explorer(data)
    
    else:
        st.info("👆 Veuillez charger un fichier de données pour commencer")
        
        st.markdown("""
        ### 📖 Instructions
        
        1. **Charger le fichier** : Cliquez sur le bouton de chargement dans la barre latérale
        2. **Sélectionner la source** : Excel ou KoBo
        3. **Analyser** : Explorez les différentes sections du dashboard
        
        ### 🎯 Fonctionnalités
        
        - ✅ Support Excel et KoBo
        - ✅ Indicateurs de performance clés
        - ✅ Visualisations interactives
        - ✅ Export des données
        - ✅ Alertes automatiques
        """)


if __name__ == "__main__":
    main()
