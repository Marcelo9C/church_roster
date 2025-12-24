import unicodedata
import streamlit as st
import pandas as pd
import re
from datetime import date, timedelta, datetime
from enum import Enum
import calendar
from fpdf import FPDF
import base64
import random
import json
import os

# --- 1. Persistência de Dados (JSON) ---
DATA_FILE = "church_data.json"
HISTORY_FILE = "history_scales.json"  # Arquivo de histórico de escalas


def load_history_scales():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}


def save_history_scales(history_dict):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history_dict, f, indent=4, ensure_ascii=False)


def save_data():
    data = {
        "volunteers": [v.to_dict() for v in st.session_state.volunteers],
        "events_config": st.session_state.events_config,
        "availability_exceptions": {
            k: {
                'full_absence': v['full_absence'],
                # Serializar datas para string ISO
                'blocked_days': [d.isoformat() for d in v['blocked_days']]
            } for k, v in st.session_state.availability_exceptions.items()
        }
    }
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)


def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)

            # Obreiros
            loaded_vols = []
            for v_data in data.get("volunteers", []):
                obj = Volunteer.from_dict(v_data)
                if obj:
                    loaded_vols.append(obj)
            st.session_state.volunteers = loaded_vols

            # Eventos
            if "events_config" in data:
                st.session_state.events_config = data["events_config"]

            # Exceções
            loaded_exc = {}
            for k, v in data.get("availability_exceptions", {}).items():
                loaded_exc[k] = {
                    'full_absence': v['full_absence'],
                    'blocked_days': [date.fromisoformat(d) for d in v['blocked_days']]
                }
            st.session_state.availability_exceptions = loaded_exc

        except Exception as e:
            st.error(f"Erro ao carregar dados: {e}")


def normalize_text(text):
    """Remove acentos e minúsculas para comparação segura"""
    if not isinstance(text, str):
        return str(text)
    return unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('ASCII').lower().strip()

# --- 2. Estrutura de Dados ---


class Role(Enum):
    # ... (Enum definitions stay same)
    PRESBITERO = "Presbítero"
    DIACONO = "Diácono"
    AUXILIAR = "Auxiliar"
    DIACONISA = "Diaconisa"
    JOVEM = "Jovem"
    MEMBRO = "Membro"


class Gender(Enum):
    MALE = "M"
    FEMALE = "F"


class Volunteer:
    def __init__(self, name, role, gender, active=True):
        self.name = name
        self.role = role
        self.gender = gender
        self.active = active

    def to_dict(self):
        return {
            "Nome": self.name,
            # Garantir que retorna o value do Enum exato
            "Cargo": self.role.value,
            "Gênero": self.gender.value,
            "Ativo": self.active
        }

    @staticmethod
    def from_dict(data):
        try:
            # Normalização BRUTA para garantir match
            raw_role = str(data["Cargo"])
            target = normalize_text(raw_role)

            role_enum = Role.MEMBRO  # Default

            for r in Role:
                if normalize_text(r.value) == target:
                    role_enum = r
                    break

            gender_val = str(data["Gênero"]).upper().strip()
            gender_enum = Gender.MALE if gender_val.startswith(
                "M") else Gender.FEMALE
            active = bool(data.get("Ativo", True))

            return Volunteer(data["Nome"], role_enum, gender_enum, active)
        except Exception:
            return None


# Inicialização do Session State
if 'volunteers' not in st.session_state:
    st.session_state.volunteers = []

if 'events_config' not in st.session_state:
    st.session_state.events_config = [
        {"name": "Culto da Palavra", "weekday": 1, "time": "18:30", "roles_needed": [
            "Responsável", "Portaria", "Recepção", "Estacionamento"]},  # Terça
        {"name": "Quinta Profética", "weekday": 3, "time": "18:30", "roles_needed": [
            "Responsável", "Portaria", "Recepção", "Recepção", "Estacionamento"]},  # Quinta
        {"name": "Escola Bíblica Dominical", "weekday": 6, "time": "08:30", "roles_needed": [
            "Responsável", "Portaria", "Recepção", "Estacionamento"]},  # Domingo Manhã
        {"name": "Culto de Adoração", "weekday": 6, "time": "17:30", "roles_needed": [
            "Responsável", "Portaria", "Recepção", "Recepção", "Estacionamento"]},  # Domingo Noite
    ]

if 'availability_exceptions' not in st.session_state:
    st.session_state.availability_exceptions = {}

# Carregar dados ao iniciar (se ainda vazios ou for reload)
if not st.session_state.volunteers and os.path.exists(DATA_FILE):
    load_data()


# --- 3. Classes Utilitárias (PDF) ---

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 16)
        self.set_text_color(0, 0, 0)
        self.cell(0, 10, 'ESCALA DE OBREIROS & VOLUNTÁRIOS', 0, 1, 'C')
        self.ln(2)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

    def draw_event_card(self, x, y, w, h, data):
        GOLD = (184, 134, 11)
        BLACK = (0, 0, 0)
        WHITE = (255, 255, 255)

        # 1. Barra de Título
        self.set_fill_color(*GOLD)
        self.rect(x, y, w, 8, 'F')

        self.set_xy(x, y)
        self.set_font('Arial', 'B', 12)
        self.set_text_color(*WHITE)

        # Cabeçalho: Nome do Evento - Horário
        # Cabeçalho: Nome do Evento (Esq) - Horário | Resp (Dir)
        # Layout Preview:
        # Left: Evento
        # Right: Time | Resp: Nome

        event_val = str(data.get('Evento', '')).upper()

        raw_time = data.get('Horário', '')
        if hasattr(raw_time, 'strftime'):
            time_val = raw_time.strftime('%H:%M')
        else:
            time_val = str(raw_time)[:5]

        resp_val = str(data.get('Responsável', 'TBD'))

        # Left align Event
        self.set_xy(x + 2, y)
        self.cell(w/2, 8, event_val, 0, 0, 'L')

        # Right align Meta
        meta_text = f"{time_val}  |  Resp: {resp_val}"
        self.set_xy(x + w/2, y)
        self.cell(w/2 - 2, 8, meta_text, 0, 1, 'R')

        body_h = h - 8
        body_y = y + 8

        # 2. Barra Lateral Esq (Dia Semanal)
        self.set_fill_color(245, 245, 245)
        self.set_draw_color(*GOLD)
        self.rect(x, body_y, 25, body_h, 'DF')

        self.set_xy(x, body_y + 5)
        self.set_font('Arial', 'B', 10)
        self.set_text_color(*BLACK)

        day_val = str(data.get('Dia', '')).upper()
        weekday_lines = day_val.split("-")
        for line in weekday_lines:
            self.set_x(x)
            self.cell(25, 4, line, 0, 1, 'C')

        # 3. Barra Lateral Dir (Data)
        self.rect(x + w - 25, body_y, 25, body_h, 'DF')
        self.set_xy(x + w - 25, body_y + body_h/2 - 3)
        date_str = str(data.get('DataStr', ''))
        self.cell(25, 6, date_str, 0, 0, 'C')

        # 4. Corpo (Lista Dinâmica de Funções)
        center_w = w - 50
        center_x = x + 25

        self.set_xy(center_x + 2, body_y + 3)
        self.set_font('Arial', '', 10)
        self.set_text_color(50, 50, 50)

        # Filtra chaves de metadados para não exibir
        # Removemos Responsável daqui pois já está no header
        ignore_keys = ['Data', 'Dia', 'Horário',
                       'Evento', 'DataStr', 'Responsável']

        # Ordenação preferencial
        priority_order = ["Portaria", "Estacionamento",
                          "Recepção 1", "Recepção 2", "Recepção 3", "Berçário"]

        # Coletar itens para exibir
        items_to_show = []

        # 1. Adicionar priorizados se existirem
        for key in priority_order:
            if key in data and data[key] and str(data[key]).lower() not in ["none", "vago", "nan", ""]:
                items_to_show.append((key, data[key]))

        # 2. Adicionar outros (Cargos Extras/Dinâmicos)
        for k, v in data.items():
            if k not in ignore_keys and k not in priority_order:
                if v and str(v).lower() not in ["none", "vago", "nan", ""]:
                    items_to_show.append((k, v))

        line_height = 6
        current_y_list = self.get_y()

        for label, val in items_to_show:
            if current_y_list > y + h - 5:
                break

            # Remover sufixos numéricos (ex: "Recepção 2" -> "Recepção") para visualização clean
            clean_label = re.sub(r' \d+$', '', label)

            self.set_x(center_x + 5)
            self.set_font('Arial', '', 10)
            # Margem label um pouco maior
            self.cell(35, line_height, f"{clean_label}:", 0, 0, 'L')

            self.set_x(center_x + 40)
            self.set_font('Arial', 'B', 10)

            # Truncar se muito longo
            val_str = str(val)
            if len(val_str) > 30:
                val_str = val_str[:28] + "..."
            self.cell(0, line_height, val_str, 0, 1, 'L')

            # Linha divisória suave
            self.set_draw_color(240, 240, 240)
            self.line(center_x + 2, self.get_y(),
                      center_x + center_w - 2, self.get_y())
            current_y_list += line_height


def create_pdf(schedule_df, title_text):
    pdf = PDF()
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()

    pdf.set_font('Arial', '', 12)
    pdf.cell(0, 10, title_text, 0, 1, 'C')
    pdf.ln(5)

    card_height = 55
    card_width = 190
    margin_x = 10
    start_y = pdf.get_y()

    current_y = start_y

    for index, row in schedule_df.iterrows():
        if current_y + card_height > 275:
            pdf.add_page()
            current_y = 20

        row_data = row.to_dict()
        # Verificar se a coluna Data é datetime, se não for (pq foi editado e virou string/object), tentar converter
        if pd.api.types.is_datetime64_any_dtype(schedule_df['Data']):
            row_data['DataStr'] = row['Data'].strftime('%d/%m')
        else:
            # Fallback se virou string no editor
            try:
                d = pd.to_datetime(row['Data']).strftime('%d/%m')
                row_data['DataStr'] = d
            except:
                row_data['DataStr'] = str(row['Data'])

        pdf.draw_event_card(margin_x, current_y,
                            card_width, card_height, row_data)
        current_y += card_height + 5

    return pdf.output(dest='S').encode('latin-1', errors='ignore')


# --- 4. Lógica de Geração ---

def get_days_in_month(year, month):
    cal = calendar.Calendar()
    days = []
    for day in cal.itermonthdays(year, month):
        if day != 0:
            days.append(date(year, month, day))
    return days


def week_day_name(weekday_idx):
    names = ["Segunda", "Terça", "Quarta",
             "Quinta", "Sexta", "Sábado", "Domingo"]
    try:
        return names[weekday_idx]
    except:
        return "Dia"


def generate_schedule_range(start_date, end_date):
    schedule = []
    volunteers = st.session_state.volunteers

    # Se não houver voluntários, tenta carregar de novo ou avisa
    if not volunteers:
        load_data()
        volunteers = st.session_state.volunteers

    # --- DEBUG INFO ---
    # Comparar por VALOR (String) para evitar erro de identidade de classe no reload do Streamlit
    debug_active = [v for v in volunteers if v.active]

    target_resp_values = [Role.PRESBITERO.value, Role.DIACONO.value]
    target_male_value = Gender.MALE.value

    debug_resp = [
        v.name for v in debug_active if v.role.value in target_resp_values]
    debug_male = [
        v.name for v in debug_active if v.gender.value == target_male_value]

    # Auto-expand se houver problemas
    has_issues = len(debug_resp) == 0 or len(debug_male) == 0

    with st.expander("🛠️ Ferramenta de Diagnóstico (Clique aqui)", expanded=has_issues):
        if not volunteers:
            st.error(
                "🛑 O banco de dados está vazio! Você clicou em '💾 Salvar Alterações da Tabela' na aba 'Cadastro'?")

        st.write(f"**Total Obreiros na Memória:** {len(volunteers)}")
        st.write(f"**Ativos:** {len(debug_active)}")

        st.markdown("---")
        st.write(
            "🕵️ **Detecção de Cargos (Como o sistema está 'lendo' cada pessoa):**")
        for v in debug_active:
            is_resp = v.role.value in target_resp_values
            role_status = "✅ RESP" if is_resp else "👤"
            st.text(f"{role_status} {v.name} -> {v.role.value}")
        st.markdown("---")

        st.write(
            f"**Candidatos a Responsável (Presb/Diac):** {len(debug_resp)}")
        st.write(f"**Homens Disponíveis (Portaria):** {len(debug_male)}")

        if not debug_resp:
            st.warning("⚠️ Nenhum 'Presbítero' ou 'Diácono' ativo encontrado.")
        if not debug_male:
            st.warning("⚠️ Nenhum homem ativo encontrado para Portaria.")
    # ------------------

    if not volunteers:
        return pd.DataFrame()

    delta = end_date - start_date
    current_days = [start_date + timedelta(days=i)
                    for i in range(delta.days + 1)]

    # Trackers for Fairness Logic

    # 1. Carregar Histórico Persistente (Arquivo JSON)
    persistent_history = load_history_scales()

    last_role_history = {}    # { 'Name': 'Role' }
    last_service_date = {}    # { 'Name': date_obj }

    # Pre-popular com dados do persistente
    for name, data in persistent_history.items():
        if 'last_role' in data:
            last_role_history[name] = data['last_role']
        if 'last_date' in data:
            try:
                last_service_date[name] = date.fromisoformat(data['last_date'])
            except:
                pass

    for current_date in current_days:
        weekday = current_date.weekday()
        day_events = [
            e for e in st.session_state.events_config if e['weekday'] == weekday]

        for event_conf in day_events:
            event_row = {
                "Data": current_date,  # Manter como objeto date para ordenação
                "Dia": week_day_name(weekday),
                "Horário": event_conf['time'],
                "Evento": event_conf['name']
            }

            # --- 1. Definição da Estrutura do Culto (Regras Dinâmicas) ---
            # Regra Culto Grande: Quintas/Sextas (Weekdays 3,4) > 18h OU Domingos (Weekday 6) > 17h
            hour_int = int(event_conf['time'].split(':')[0])
            is_thursday_friday_night = (weekday in [3, 4] and hour_int >= 18)
            is_sunday_night = (weekday == 6 and hour_int >= 17)
            is_big_service = is_thursday_friday_night or is_sunday_night

            # --- 2. Preparação do Pool ---
            # ... (código existente de disponibilidade removido para brevidade, mantendo lógica)

            # Disponibilidade Check (Simplificado - mantendo o anterior)
            available_pool = []
            year = current_date.year
            month = current_date.month
            for v in volunteers:
                if not v.active:
                    continue
                exc_key = f"{v.name}_{year}-{month}"
                exceptions = st.session_state.availability_exceptions.get(
                    exc_key, {})
                if exceptions.get('full_absence', False):
                    continue
                if current_date in exceptions.get('blocked_days', []):
                    continue
                available_pool.append(v)
            pool = available_pool if available_pool else [
                v for v in volunteers if v.active]

            # --- 3. Alocação de Papéis (DINÂMICO) ---

            used_names = set()  # Inicializar conjunto de nomes usados neste evento

            # Helper para selecionar candidato com base em regras de 'Cansaço' e 'Rodízio'
            def get_candidate_tiered(candidates, role_target, current_date_obj):
                # Filtra quem já está neste culto
                base_pool = [c for c in candidates if c.name not in used_names]
                if not base_pool:
                    return None

                # Tier 1: Ideal (Não trabalhou no anterior E Não fez esse papel na ultima vez)
                tier1 = []
                # Tier 2: Aceitável
                tier2 = []  # Fresco mas repetindo função OU Cansado mas mudando função
                tier3 = []  # Cansado e repetindo função (evitar ao máximo)

                for cand in base_pool:
                    last_date = last_service_date.get(cand.name)
                    last_role = last_role_history.get(cand.name)

                    is_tired = False
                    if last_date:
                        delta = (current_date_obj - last_date).days
                        # Regra User: "reversar nas escalas... sem repetir posições bem como dias seguidos"
                        if delta <= 1:
                            is_tired = True

                    is_same_role = (last_role == role_target)

                    if not is_tired and not is_same_role:
                        tier1.append(cand)
                    elif (is_tired and not is_same_role) or (not is_tired and is_same_role):
                        tier2.append(cand)
                    else:
                        tier3.append(cand)

                chosen = None
                if tier1:
                    chosen = random.choice(tier1)
                elif tier2:
                    chosen = random.choice(tier2)
                elif tier3:
                    chosen = random.choice(tier3)
                else:
                    chosen = random.choice(base_pool)

                if chosen:
                    used_names.add(chosen.name)
                    last_role_history[chosen.name] = role_target
                    last_service_date[chosen.name] = current_date_obj

                return chosen

            # Combinar Roles fixos com Extras
            # A lista "roles_needed" já vem do Configurações, onde o usuário pode adicionar manualmente.
            # Não precisamos de "extra_roles_map" separado se o usuário editar o Config antes de gerar.

            current_roles_needed = event_conf.get("roles_needed", []).copy()

            # Counter para sufixos de roles duplicadas (Ex: Recepção, Recepção -> Recepção, Recepção 2)
            role_counts = {}

            for role_name_raw in current_roles_needed:
                # Normalizar Nome para Chave Única
                base_name = role_name_raw.strip()
                if base_name not in role_counts:
                    role_counts[base_name] = 1
                    final_key = base_name
                else:
                    role_counts[base_name] += 1
                    final_key = f"{base_name} {role_counts[base_name]}"

                # Definir Restrições de Gênero/Cargo com base no nome original (sem número)
                role_norm = base_name.lower()

                candidates = []

                if "responsável" in role_norm:
                    # Apenas Presbíteros (Regra Estrita)
                    candidates = [
                        v for v in pool if v.role.value == Role.PRESBITERO.value]
                elif "portaria" in role_norm or "estacionamento" in role_norm:
                    # Apenas Homens (Excluindo Presbíteros, pois só podem ser Responsáveis)
                    candidates = [v for v in pool if v.gender.value ==
                                  Gender.MALE.value and v.role.value != Role.PRESBITERO.value]
                elif "recepção" in role_norm or "berçário" in role_norm:
                    # Apenas Mulheres (Excluindo Presbíteros caso existam mulheres presbítero no futuro, ou por segurança)
                    candidates = [v for v in pool if v.gender.value ==
                                  Gender.FEMALE.value and v.role.value != Role.PRESBITERO.value]
                else:
                    # GENÉRICO (Cargos Extras) - Excluir Presbíteros
                    candidates = [
                        v for v in pool if v.role.value != Role.PRESBITERO.value]

                chosen = get_candidate_tiered(
                    candidates, base_name, current_date)
                event_row[final_key] = chosen.name if chosen else "Vago"

            schedule.append(event_row)

    return pd.DataFrame(schedule)


# --- 5. Interface UI/CSS ---

st.set_page_config(page_title="Gerenciador eclesiástico",
                   layout="wide", page_icon="⛪")

st.markdown("""
<style>
    /* Global Styles */
    .stButton>button { width: 100%; border-radius: 6px; height: 3em; font-weight: 600; }
    div[data-testid="stExpander"] { border: 1px solid #4bb0b6; border-radius: 8px; }
    
    /* Card Preview Styles */
    .card-preview {
        border: 1px solid #B8860B;
        border-radius: 8px;
        margin-bottom: 15px;
        background-color: #2b2b2b;
        color: white;
        box-shadow: 0 4px 6px rgba(0,0,0,0.3);
        overflow: hidden;
    }
    .card-header {
        background: linear-gradient(90deg, #B8860B 0%, #daa520 100%);
        color: white;
        padding: 8px 15px;
        font-weight: bold;
        font-size: 1.1em;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .card-meta { font-size: 0.9em; opacity: 0.9; }
    .card-body { display: flex; padding: 0; }
    .card-sidebar { width: 80px; display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 10px; font-weight: bold; text-align: center; }
    .card-day { background-color: #333; border-right: 2px solid #B8860B; font-size: 0.9em; text-transform: uppercase; }
    .card-date { background-color: #333; border-left: 2px solid #B8860B; }
    .card-content { flex-grow: 1; padding: 10px 15px; background-color: #1e1e1e; }
    .role-row { display: flex; justify-content: space-between; border-bottom: 1px solid #333; padding: 6px 0; align-items: center; }
    .role-row:last-child { border-bottom: none; }
    .role-label { color: #aaa; font-size: 0.9em; }
    .role-name { font-weight: bold; color: #fff; }
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2664/2664627.png", width=80)
    st.title("ADHR SistemaS")
    st.caption("v1.3.0 - Persistência Ativa")
    st.markdown("---")
    st.info("👋 **Bem-vindo!** Seus dados agora são salvos automaticamente.")

    st.write("© 2025 ADHR")

st.title("⛪ Gerenciador de Escalas")

tab1, tab2, tab3, tab4 = st.tabs(
    ["📅 Gerar Escala", "🚫 Disponibilidade", "⚙️ Configurações", "👥 Cadastro (Editar/Excluir)"])

# --- TAB 1: GERAR ESCALA (Com Edição) ---
with tab1:
    with st.container():
        st.subheader("Painel de Geração")

        c_mode, c_date = st.columns([1, 2])
        with c_mode:
            gen_mode = st.radio("Tipo de Escala", [
                                "Mensal", "Semanal"], horizontal=True)

        start_d, end_d = None, None
        title_ref = ""

        with c_date:
            if gen_mode == "Mensal":
                col_y, col_m = st.columns(2)
                with col_y:
                    s_year = st.number_input("Ano", 2024, 2030, 2025)
                with col_m:
                    s_month = st.selectbox("Mês", range(1, 13))

                start_d = date(s_year, s_month, 1)
                last_day = calendar.monthrange(s_year, s_month)[1]
                end_d = date(s_year, s_month, last_day)
                title_ref = f"Mês de Referência: {calendar.month_name[s_month].capitalize()}/{s_year}"

            else:  # Semanal
                start_d = st.date_input(
                    "Início da Semana (Segunda)", value=date.today())
                end_d = start_d + timedelta(days=6)
                st.caption(f"Até: {end_d.strftime('%d/%m/%Y')}")
                title_ref = f"Escala Semanal: {start_d.strftime('%d/%m')} a {end_d.strftime('%d/%m/%Y')}"

        if st.button("🚀 Gerar Nova Escala", type="primary"):
            st.balloons()  # <--- ADICIONE ESTA LINHA AQUI
    # ... o resto do código continua aqui embaixo ...
            df = generate_schedule_range(start_d, end_d)
            if df.empty:
                st.warning(
                    "Nenhuma escala gerada. Verifique se há obreiros cadastrados e eventos configurados.")
            else:
                st.session_state['generated_schedule'] = df
                st.success("Escala Gerada! Edite abaixo se necessário.")

        # Se existir escala gerada na memória, mostrar editor e preview
        if 'generated_schedule' in st.session_state and isinstance(st.session_state['generated_schedule'], pd.DataFrame):
            edited_df = st.session_state['generated_schedule'].copy()

            # Converter colunas para tipos compatíveis com o Editor
            if "Horário" in edited_df.columns:
                try:
                    # Converte para string primeiro (garante que funciona se já for time object)
                    # Usa 'mixed' para aceitar HH:MM ou HH:MM:SS
                    series_str = edited_df["Horário"].astype(str)
                    edited_df["Horário"] = pd.to_datetime(
                        series_str, format='mixed').dt.time
                except Exception:
                    # print(f"Erro convertendo horário: {e}")
                    pass

            st.write("### ✏️ Editar Escala Gerada")
            st.info(
                "Você pode alterar nomes, horários ou eventos diretamente aqui antes de baixar o PDF.")

            # Editor da Escala
            edited_schedule = st.data_editor(
                edited_df,
                num_rows="dynamic",
                key="schedule_editor",
                use_container_width=True,
                column_config={
                    "Data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
                    "Dia": st.column_config.TextColumn("Dia"),
                    "Horário": st.column_config.TimeColumn("Hora", format="HH:mm"),
                    "Evento": st.column_config.TextColumn("Culto/Evento"),
                    "Responsável": st.column_config.TextColumn("Responsável", help="Presbítero Responsável"),
                    "Portaria": st.column_config.TextColumn("Portaria"),
                    "Estacionamento": st.column_config.TextColumn("Estacionamento"),
                    "Recepção": st.column_config.TextColumn("Recepção 1"),
                    "Recepção 2": st.column_config.TextColumn("Recepção 2"),
                    "Recepção 3": st.column_config.TextColumn("Recepção 3"),
                    "Berçário": st.column_config.TextColumn("Berçário"),
                    "Galeria": st.column_config.TextColumn("Galeria (Extra)"),
                    "Apoio": st.column_config.TextColumn("Apoio (Extra)"),
                },
                hide_index=True
            )

            # Botão PDF
            if st.button("🔄 Atualizar Prévia e PDF"):
                st.session_state['generated_schedule'] = edited_schedule
                st.success("Atualizado!")

            # Botão Salvar Histórico (Confirmar Escala)
            st.divider()
            col_save_hist, _ = st.columns([1.5, 2])
            with col_save_hist:
                st.info(
                    "👆 Se a escala estiver OK, clique abaixo para salvar no histórico e alimentar a inteligência do sistema.")
                if st.button("💾 Oficializar Escala (Salvar no Histórico)"):
                    # Ler o DataFrame FINAL (Editado)
                    # Atualizar o arquivo JSON com a última função/data de cada um

                    current_hist = load_history_scales()

                    # Iterar cronologicamente para garantir que pegamos o último evento
                    # Ordenar por data
                    final_df_sorted = final_df.sort_values(by="Data")

                    count_updates = 0
                    for _, row in final_df_sorted.iterrows():
                        d_str = row['DataStr']  # Já formatado ou raw
                        # Tentar pegar data real
                        d_obj = row['Data']
                        if not isinstance(d_obj, (date, datetime)):
                            # Tentar converter
                            try:
                                d_obj = pd.to_datetime(row['Data']).date()
                            except:
                                continue

                        d_iso = d_obj.isoformat()

                        # Mapear colunas de cargos
                        roles_cols = ["Responsável", "Portaria", "Estacionamento",
                                      "Recepção", "Recepção 2", "Recepção 3", "Berçário"]

                        for col in roles_cols:
                            if col in row:
                                person_name = str(row[col]).strip()
                                if person_name and person_name not in ["", "None", "Vago", "TBD"]:
                                    # Atualizar
                                    current_hist[person_name] = {
                                        # Remover " 2" etc
                                        "last_role": col.split(" ")[0],
                                        "last_date": d_iso
                                    }
                                    count_updates += 1

                    save_history_scales(current_hist)
                    st.success(
                        f"Histórico atualizado com sucesso! ({count_updates} registros processados)")

            # Usar o DF editado para o PDF e Preview
            final_df = edited_schedule

            pdf_bytes = create_pdf(final_df, title_ref)
            b64 = base64.b64encode(pdf_bytes).decode()
            fname = f"Escala_Oficial.pdf"

            col_pdf, col_prev = st.columns([1, 2])
            with col_pdf:
                href = f'<a href="data:application/octet-stream;base64,{b64}" download="{fname}">' \
                    f'<button style="background-color:#E74C3C;color:white;border:none;padding:15px;width:100%;border-radius:5px;cursor:pointer;font-size:1.1em;">' \
                    f'📄 BAIXAR PDF FINAL</button></a>'
                st.markdown(href, unsafe_allow_html=True)

            with col_prev:
                st.markdown("### Prévia do Documento")
                for mn, row in final_df.iterrows():
                    # Lógica de formatação segura
                    try:
                        data_str = row['Data'].strftime(
                            '%d/%m') if hasattr(row['Data'], 'strftime') else str(row['Data'])
                    except:
                        data_str = str(row.get('Data', ''))

                    dia_str = str(row.get('Dia', ''))

                    roles_html = ""
                    exclude_keys = ['Data', 'Dia', 'Horário',
                                    'Evento', 'Responsável', 'DataStr']
                    roles_dict = {k: v for k,
                                  v in row.items() if k not in exclude_keys}

                    # Ordenar Preview
                    prio = ["Portaria", "Estacionamento", "Recepção 1",
                            "Recepção 2", "Recepção 3", "Berçário"]

                    # 1. Prio
                    for k in prio:
                        if k in roles_dict:
                            val = roles_dict.pop(k)
                            if val and str(val) != "Vago" and str(val) != "nan":
                                roles_html += f'<div class="role-row"><span class="role-label">{k}</span><span class="role-name">{val}</span></div>'

                    # 2. Resto (Extras como Galeria)
                    for r, n in roles_dict.items():
                        if n and str(n) != "N/A" and str(n).lower() != "nan":
                            roles_html += f'<div class="role-row"><span class="role-label">{r}</span><span class="role-name">{n}</span></div>'

                    st.markdown(f"""
                    <div class="card-preview">
                        <div class="card-header">
                            <span>{row.get('Evento', '')}</span>
                            <span class="card-meta">⏰ {row.get('Horário', '')}  |  👤 Resp: {row.get('Responsável', '')}</span>
                        </div>
                        <div class="card-body">
                            <div class="card-sidebar card-day">
                                {dia_str.upper().replace('-', '<br>')}
                            </div>
                            <div class="card-content">
                                {roles_html}
                            </div>
                            <div class="card-sidebar card-date">
                                📅<br>{data_str}
                            </div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)


# --- TAB 2: DISPONIBILIDADE ---
with tab2:
    st.header("Gerenciador de Exceções")
    if not st.session_state.volunteers:
        st.warning("Sem cadastros.")
    else:
        col_filtro1, col_filtro2 = st.columns(2)
        with col_filtro1:
            d_year = st.number_input("Ano", 2024, 2030, 2025, key='exc_year')
        with col_filtro2:
            d_month = st.selectbox("Mês", range(1, 13), key='exc_month')

        volunteer_names = [v.name for v in st.session_state.volunteers]
        sel_name = st.selectbox("Buscar Obreiro:", volunteer_names)

        if sel_name:
            st.divider()
            exc_key = f"{sel_name}_{d_year}-{d_month}"
            curr = st.session_state.availability_exceptions.get(exc_key, {})

            with st.form("exc_form"):
                full = st.checkbox("Ausência Total no Mês",
                                   value=curr.get('full_absence', False))
                days = get_days_in_month(d_year, d_month)
                days_fmt = {
                    d: f"{d.strftime('%d/%m')} ({week_day_name(d.weekday())})" for d in days}

                blocked_dates_input = curr.get('blocked_days', [])
                # Garantir que são objetos date
                default_vals = []
                for b in blocked_dates_input:
                    if isinstance(b, date):
                        default_vals.append(b)
                    elif isinstance(b, str):
                        default_vals.append(date.fromisoformat(b))

                blocked = st.multiselect(
                    "Dias Indisponíveis specificos:",
                    days,
                    default=default_vals,
                    format_func=lambda x: days_fmt[x]
                )

                if st.form_submit_button("Salvar Restrição"):
                    st.session_state.availability_exceptions[exc_key] = {
                        'full_absence': full,
                        'blocked_days': blocked
                    }
                    save_data()  # SAVE AUTO
                    st.success("Salvo!")

# --- TAB 3: CONFIGURAÇÕES ---
with tab3:
    st.header("⚙️ Configuração dos Cultos")
    st.info("💡 **Dica:** Aqui você define a estrutura padrão. Para 'Cargos Extras' (ex: Galeria), basta digitar o nome do cargo na coluna 'Funções' (separado por vírgula) para o culto desejado.")
    st.caption("Exemplo: `Responsável, Portaria, Galeria`")

    events_data = []
    for e in st.session_state.events_config:
        events_data.append({
            "Nome": e["name"],
            "Dia Semana": e["weekday"],
            "Horário": e["time"],
            "Funções": ", ".join(e["roles_needed"])
        })

    WEEKDAY_MAP = {0: "Segunda", 1: "Terça", 2: "Quarta",
                   3: "Quinta", 4: "Sexta", 5: "Sábado", 6: "Domingo"}
    REVERSE_MAP = {v: k for k, v in WEEKDAY_MAP.items()}

    from datetime import datetime
    events_data_processed = []
    for row in events_data:
        new_row = row.copy()
        new_row["Dia Nome"] = WEEKDAY_MAP.get(
            row["Dia Semana"], "Desconhecido")
        try:
            # Se já for time object (pq veio do load), ok. Se for string, convert.
            if isinstance(row["Horário"], str):
                new_row["Horário"] = datetime.strptime(
                    row["Horário"], "%H:%M").time()
            else:
                new_row["Horário"] = row["Horário"]
        except:
            new_row["Horário"] = None
        events_data_processed.append(new_row)

    df_events = pd.DataFrame(events_data_processed)
    if df_events.empty:
        df_events = pd.DataFrame(
            columns=["Nome", "Dia Nome", "Horário", "Funções"])

    edited_df = st.data_editor(
        df_events,
        num_rows="dynamic",
        key="ev_edit",
        use_container_width=True,
        column_config={
            "Dia Nome": st.column_config.SelectboxColumn("Dia da Semana", options=list(WEEKDAY_MAP.values()), required=True),
            "Dia Semana": None,
            "Nome": st.column_config.TextColumn("Nome do Culto", required=True),
            "Horário": st.column_config.TimeColumn("Horário (HH:mm)", required=True, format="HH:mm"),
            "Funções": st.column_config.TextColumn("Cargos (separados por vírgula)")
        },
        column_order=["Nome", "Dia Nome", "Horário", "Funções"]
    )

    if st.button("💾 Salvar Configurações de Eventos"):
        new_c = []
        records = edited_df.to_dict('records')
        for row in records:
            if row["Nome"]:
                w_int = REVERSE_MAP.get(row.get("Dia Nome"), 0)
                roles_str = row.get("Funções", "")
                roles_list = [r.strip()
                              for r in roles_str.split(",") if r.strip()]

                time_val = row["Horário"]
                time_str = time_val.strftime("%H:%M") if time_val else "00:00"

                new_c.append({
                    "name": row["Nome"],
                    "weekday": w_int,
                    "time": time_str,
                    "roles_needed": roles_list if roles_list else ["Responsável"]
                })
        st.session_state.events_config = new_c
        save_data()  # SAVE AUTO
        st.success("Eventos atualizados com sucesso!")

# --- TAB 4: CADASTRO (DATA EDITOR) ---
with tab4:
    st.subheader("📂 Importar/Exportar Dados")

    # 1. Área de UPLOAD (Importar)
    uploaded_file = st.file_uploader(
        "Restaurar Backup ou Importar Lista (CSV)", type=["csv"])

    if uploaded_file is not None:
        try:
            # Lendo com o separador ';' para ser compatível com o arquivo que o Excel salva
            df_import = pd.read_csv(uploaded_file, sep=';')

            # Limpeza dos nomes das colunas (remove espaços extras que o Excel possa ter criado)
            df_import.columns = [c.strip() for c in df_import.columns]

            required_cols = ["Nome", "Cargo", "Gênero"]

            if all(col in df_import.columns for col in required_cols):

                st.info(
                    f"Arquivo carregado com {len(df_import)} registros. O que deseja fazer?")
                col_btn1, col_btn2 = st.columns(2)

                # Opção A: SUBSTITUIR (Zera e coloca o novo)
                if col_btn1.button("⚠️ SUBSTITUIR toda a lista atual"):
                    new_volunteers = []
                    for index, row in df_import.iterrows():
                        # Verifica se a coluna 'Ativo' existe, senão assume True
                        active_status = row["Ativo"] if "Ativo" in row else True

                        v = Volunteer.from_dict({
                            "Nome": row["Nome"],
                            "Cargo": row["Cargo"],
                            "Gênero": row["Gênero"],
                            "Ativo": active_status
                        })
                        if v:
                            new_volunteers.append(v)

                    st.session_state.volunteers = new_volunteers
                    save_data()
                    st.success(
                        "Lista substituída! Pressione 'R' para recarregar.")
                    st.rerun()

                # Opção B: MESCLAR (Mantém os atuais e soma os novos)
                if col_btn2.button("➕ ADICIONAR aos existentes"):
                    added_count = 0
                    # Cria lista de nomes atuais normalizada (minúscula) para evitar duplicatas
                    current_names = [v.name.lower().strip()
                                     for v in st.session_state.volunteers]

                    for index, row in df_import.iterrows():
                        name_clean = str(row["Nome"]).strip()
                        # Só adiciona se o nome não existir
                        if name_clean.lower() not in current_names:
                            active_status = row["Ativo"] if "Ativo" in row else True
                            v = Volunteer.from_dict({
                                "Nome": row["Nome"],
                                "Cargo": row["Cargo"],
                                "Gênero": row["Gênero"],
                                "Ativo": active_status
                            })
                            if v:
                                st.session_state.volunteers.append(v)
                                added_count += 1

                    save_data()
                    st.success(
                        f"{added_count} novos obreiros adicionados! Pressione 'R'.")
                    st.rerun()
            else:
                st.error(
                    f"O arquivo CSV precisa ter as colunas: {', '.join(required_cols)}. Verifique se está separado por ponto e vírgula (;).")

        except Exception as e:
            st.error(
                f"Erro ao ler o arquivo. Verifique se é um CSV válido. Detalhe: {e}")

    st.divider()

    # 2. Área de DOWNLOAD (Backup) - Com a correção de acentos para Excel
    st.markdown("### 💾 Backup dos Dados")
    col_bkp1, col_bkp2 = st.columns([2, 1])

    with col_bkp1:
        st.info("Baixe a planilha formatada para Excel (com acentos corrigidos).")

    with col_bkp2:
        if st.session_state.volunteers:
            # 1. Converter objetos para DataFrame
            current_data = [v.to_dict() for v in st.session_state.volunteers]
            df_export = pd.DataFrame(current_data)

            # 2. Gerar o texto CSV (separado por ponto e vírgula)
            # Nota: Não passamos encoding aqui, pois ele ignoraria ao gerar string
            csv_text = df_export.to_csv(index=False, sep=';')

            # 3. A MÁGICA: Converter Texto -> Bytes com a assinatura BOM (UTF-8-SIG)
            # Isso força o Excel a entender que é UTF-8
            csv_bytes = csv_text.encode('utf-8-sig')

            st.download_button(
                label="📥 Baixar Planilha (Excel)",
                data=csv_bytes,  # Entregamos os bytes assinados, não o texto
                file_name="obreiros_backup.csv",
                mime="text/csv"
            )

    st.divider()

    # 3. Formulário Manual (Mantido original)
    with st.expander("➕ Adicionar Novo Obreiro (Cadastro Rápido)", expanded=True):
        st.caption(
            "Preencha os dados e clique em 'Salvar Novo' para adicionar rapidamente.")
        with st.form("new_volunteer_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                new_name = st.text_input("Nome Completo")
            with c2:
                # Recupera os valores do Enum Role
                new_role = st.selectbox("Cargo", [r.value for r in Role])
            new_gender = st.radio("Gênero", ["M", "F"], horizontal=True)

            if st.form_submit_button("Salvar Novo"):
                if new_name:
                    # Lógica para converter string de volta para Enum e salvar
                    role_enum = next(r for r in Role if r.value == new_role)
                    gender_enum = Gender.MALE if new_gender == "M" else Gender.FEMALE
                    st.session_state.volunteers.append(
                        Volunteer(new_name, role_enum, gender_enum))
                    save_data()
                    st.success(f"{new_name} adicionado!")
                    st.rerun()

    st.divider()
    st.subheader("📋 Tabela Geral de Obreiros")
    st.warning("**Como Editar ou Excluir:**\n1. Clique duas vezes para **Editar**.\n2. Del para **Excluir**.\n3. Clique em **Salvar**.")

    if st.session_state.volunteers:
        data_rows = [v.to_dict() for v in st.session_state.volunteers]
        df_volunteers = pd.DataFrame(data_rows)
    else:
        df_volunteers = pd.DataFrame(
            columns=["Nome", "Cargo", "Gênero", "Ativo"])

    edited_volunteers = st.data_editor(
        df_volunteers,
        num_rows="dynamic",
        column_config={
            "Nome": st.column_config.TextColumn("Nome Completo", required=True),
            "Cargo": st.column_config.SelectboxColumn("Cargo", options=[r.value for r in Role], required=True),
            "Gênero": st.column_config.SelectboxColumn("Gênero", options=["M", "F"], required=True),
            "Ativo": st.column_config.CheckboxColumn("Ativo?", default=True)
        },
        use_container_width=True,
        key="editor_volunteers"
    )

    if st.button("💾 Salvar Alterações da Tabela"):
        new_list = []
        records = edited_volunteers.to_dict('records')
        for row in records:
            if row.get("Nome") and str(row["Nome"]).strip() != "":
                obj = Volunteer.from_dict(row)
                if obj:
                    new_list.append(obj)
        st.session_state.volunteers = new_list
        save_data()
        st.success("Tabela atualizada com sucesso!")
        st.rerun()
