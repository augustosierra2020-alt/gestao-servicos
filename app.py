import streamlit as st
import pandas as pd
import random
from datetime import datetime, timedelta
import json
import hashlib
from streamlit_gsheets import GSheetsConnection
import extra_streamlit_components as stx

# --- CONFIGURAÇÃO INICIAL DA PÁGINA ---
st.set_page_config(page_title="Gestão de Designações e Partes", layout="wide", initial_sidebar_state="expanded")

# --- CONFIGURAÇÃO DE ADMINISTRADOR MASTER ---
EMAIL_ADMIN = "augustosierra2020@gmail.com"

# --- 🛡️ FUNÇÕES DE BLINDAGEM (PROGRAMAÇÃO DEFENSIVA) ---
def formatar_data_segura(serie_datas):
    """Amortecedor global para datas: Tenta converter e formatar de forma segura."""
    try:
        return pd.to_datetime(serie_datas, dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
    except Exception:
        return serie_datas.astype(str)

def normalizar_tabela(df):
    """Garante que os tipos de dados sejam puros e exatos, evitando falhas do Streamlit"""
    if df is None or df.empty:
        return pd.DataFrame({
            "Grupo": pd.Series(dtype="string"),
            "Tarefa": pd.Series(dtype="string"),
            "Data de Trabalho": pd.Series(dtype="datetime64[ns]"),
            "Principal": pd.Series(dtype="string"),
            "Ajudante": pd.Series(dtype="string"),
            "Data de Registro": pd.Series(dtype="string")
        })
    
    df = df.copy()
    
    try:
        df["Data de Trabalho"] = pd.to_datetime(df["Data de Trabalho"], dayfirst=True, errors="coerce")
        df["Data de Trabalho"] = df["Data de Trabalho"].fillna(pd.Timestamp("today").normalize())
    except Exception:
        pass 
    
    colunas_texto = ["Grupo", "Tarefa", "Principal", "Ajudante", "Data de Registro"]
    for col in colunas_texto:
        if col in df.columns:
            df[col] = df[col].astype(str).replace(["nan", "NaN", "NaT", "None", "<NA>"], "")
            
    return df

# --- 🛡️ INICIALIZAÇÃO DE MEMÓRIA ---
if "user_id" not in st.session_state:
    st.session_state.user_id = None
if "user_nome" not in st.session_state:
    st.session_state.user_nome = None
if "user_email" not in st.session_state:
    st.session_state.user_email = None
if "view_mode" not in st.session_state:
    st.session_state.view_mode = "app" 
if "escala_temporaria" not in st.session_state:
    st.session_state.escala_temporaria = None
if "deslogado" not in st.session_state:
    st.session_state.deslogado = False
if "historico_definitivo" not in st.session_state:
    st.session_state.historico_definitivo = normalizar_tabela(pd.DataFrame())
if "historico_carregado" not in st.session_state:
    st.session_state.historico_carregado = False
if "admin_verificado" not in st.session_state:
    st.session_state.admin_verificado = False

# --- FUNÇÃO DE HORÁRIO (BRASÍLIA) ---
def get_horario_brasilia():
    return (datetime.utcnow() - timedelta(hours=3)).strftime("%d/%m/%Y %H:%M:%S")

# --- CONEXÃO COM O GOOGLE SHEETS ---
conn_sheets = st.connection("gsheets", type=GSheetsConnection)

# 🚀 CACHE E LEITURA OTIMIZADA COM BLINDAGEM
@st.cache_data(ttl=300, show_spinner=False)
def carregar_aba(aba_nome):
    try:
        import requests
        URL_DO_SCRIPT = st.secrets["connections"]["gsheets"].get("script_url", "")
        if URL_DO_SCRIPT:
            payload = {"action": "read", "aba": aba_nome}
            response = requests.post(URL_DO_SCRIPT, json=payload, timeout=10)
            if response.status_code == 200:
                try:
                    res_json = response.json()
                    if res_json.get("status") == "success":
                        dados = res_json.get("dados", [])
                        if dados:
                            df = pd.DataFrame(dados)
                            df.columns = df.columns.str.strip().str.lower()
                            return df
                except ValueError:
                    pass
    except Exception:
        pass
        
    try:
        df = conn_sheets.read(worksheet=aba_nome, ttl=0)
        if df is not None and not df.empty:
            df.columns = df.columns.str.strip().str.lower()
            df = df.dropna(how='all').reset_index(drop=True)
            return df
    except Exception:
        pass
    return pd.DataFrame()

def salvar_aba(df, aba_nome):
    try:
        import requests
        URL_DO_SCRIPT = st.secrets["connections"]["gsheets"].get("script_url", "")
        if URL_DO_SCRIPT:
            if "id" in df.columns:
                df["id"] = pd.to_numeric(df["id"], errors='coerce').fillna(0).astype(int)
            payload = {"action": "write", "aba": aba_nome, "dados": df.to_json(orient="records", date_format="iso")}
            response = requests.post(URL_DO_SCRIPT, json=payload, timeout=10)
            if response.status_code == 200:
                carregar_aba.clear(aba_nome)
                return True
    except Exception as e:
        st.error(f"Erro ao salvar na nuvem: {e}")
    return False

# --- FUNÇÕES DE SEGURANÇA E GERENCIAMENTO ---
def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

def inicializar_admin_master():
    df_usuarios = carregar_aba("usuarios")
    colunas_ordem = ["id", "nome", "email", "senha", "grupos_json"]
    
    if df_usuarios.empty or "email" not in df_usuarios.columns:
        admin_fixo = pd.DataFrame([{
            "id": 1,
            "nome": "Sérgio Sierra",
            "email": EMAIL_ADMIN.strip().lower(),
            "senha": hash_senha("adm01"),
            "grupos_json": json.dumps({})
        }])
        salvar_aba(admin_fixo[colunas_ordem], "usuarios")
    else:
        df_usuarios["email"] = df_usuarios["email"].astype(str).str.strip().str.lower()
        if not (df_usuarios["email"] == EMAIL_ADMIN.lower()).any():
            try:
                ids_validos = pd.to_numeric(df_usuarios["id"], errors='coerce').dropna()
                novo_id = int(ids_validos.max()) + 1 if not ids_validos.empty else 1
            except:
                novo_id = len(df_usuarios) + 1
                
            admin_fixo = pd.DataFrame([{
                "id": int(novo_id),
                "nome": "Sérgio Sierra",
                "email": EMAIL_ADMIN.strip().lower(),
                "senha": hash_senha("adm01"),
                "grupos_json": json.dumps({})
            }])
            df_usuarios = pd.concat([df_usuarios, admin_fixo], ignore_index=True)
            salvar_aba(df_usuarios[colunas_ordem], "usuarios")

if not st.session_state.admin_verificado:
    inicializar_admin_master()
    st.session_state.admin_verificado = True

def cadastrar_usuario(nome, email, senha):
    df_usuarios = carregar_aba("usuarios")
    colunas_ordem = ["id", "nome", "email", "senha", "grupos_json"]
    
    if df_usuarios.empty or "email" not in df_usuarios.columns:
        df_usuarios = pd.DataFrame(columns=colunas_ordem)
        novo_id = 1
    else:
        df_usuarios["email"] = df_usuarios["email"].astype(str).str.strip().str.lower()
        df_usuarios["nome"] = df_usuarios["nome"].astype(str).str.strip().str.lower()
        
        if (df_usuarios["email"] == email.strip().lower()).any():
            return "email_duplicado"
        if (df_usuarios["nome"] == nome.strip().lower()).any():
            return "nome_duplicado"
            
        try:
            ids_validos = pd.to_numeric(df_usuarios["id"], errors='coerce').dropna()
            novo_id = int(ids_validos.max()) + 1 if not ids_validos.empty else len(df_usuarios) + 1
        except:
            novo_id = len(df_usuarios) + 1
    
    novo_registro = pd.DataFrame([{
        "id": int(novo_id),
        "nome": nome.strip(),
        "email": email.strip().lower(),
        "senha": hash_senha(senha.strip()),
        "grupos_json": json.dumps({})
    }])
    
    df_usuarios = carregar_aba("usuarios")
    if df_usuarios.empty:
        df_usuarios = novo_registro[colunas_ordem]
    else:
        df_usuarios = df_usuarios[colunas_ordem]
        novo_registro = novo_registro[colunas_ordem]
        df_usuarios = pd.concat([df_usuarios, novo_registro], ignore_index=True)
        
    if salvar_aba(df_usuarios, "usuarios"):
        return "sucesso"
    return "erro_salvar"

def verificar_login(email, senha):
    email_limpo = email.strip().lower()
    senha_limpa = senha.strip()
    
    df_usuarios = carregar_aba("usuarios")
    if not df_usuarios.empty and "email" in df_usuarios.columns and "senha" in df_usuarios.columns:
        df_usuarios["email"] = df_usuarios["email"].astype(str).str.strip().str.lower()
        df_usuarios["senha"] = df_usuarios["senha"].astype(str).str.strip()
        
        user_row = df_usuarios[df_usuarios["email"] == email_limpo]
        if not user_row.empty:
            if email_limpo == EMAIL_ADMIN.lower():
                if user_row.iloc[0]["senha"] == hash_senha(senha_limpa) or senha_limpa == "adm01":
                    return user_row.iloc[0].to_dict()
            else:
                if user_row.iloc[0]["senha"] == hash_senha(senha_limpa):
                    return user_row.iloc[0].to_dict()

    if email_limpo == EMAIL_ADMIN.lower() and senha_limpa == "adm01":
        return {"id": 1, "nome": "Sérgio Sierra", "email": EMAIL_ADMIN.lower(), "grupos_json": json.dumps({})}
    return None

def buscar_usuario_por_id(user_id):
    if int(user_id) == 1:
        df_usuarios = carregar_aba("usuarios")
        if not df_usuarios.empty and "id" in df_usuarios.columns:
            user = df_usuarios[pd.to_numeric(df_usuarios["id"], errors='coerce') == 1]
            if not user.empty:
                return user.iloc[0].to_dict()
        return {"id": 1, "nome": "Sérgio Sierra", "email": EMAIL_ADMIN.lower(), "grupos_json": json.dumps({})}

    df_usuarios = carregar_aba("usuarios")
    if df_usuarios.empty or "id" not in df_usuarios.columns:
        return None
    user = df_usuarios[pd.to_numeric(df_usuarios["id"], errors='coerce') == int(user_id)]
    if not user.empty:
        return user.iloc[0].to_dict()
    return None

# --- FUNÇÕES DO HISTÓRICO EM BLOCOS ---
def acrescentar_historico_db(user_id, df_novo_bloco):
    df_global = carregar_aba("historico")
    novos_registros = []
    proximo_id = 1 if df_global.empty or "id" not in df_global.columns else int(df_global["id"].max()) + 1
    
    for _, row in df_novo_bloco.iterrows():
        dt_val = row['Data de Trabalho']
        if isinstance(dt_val, pd.Timestamp) or isinstance(dt_val, datetime):
            data_str = dt_val.strftime("%Y-%m-%d")
        else:
            try: data_str = pd.to_datetime(dt_val, dayfirst=True).strftime("%Y-%m-%d")
            except: data_str = str(dt_val)

        novos_registros.append({
            "id": proximo_id,
            "user_id": int(user_id),
            "grupo": str(row.get('Grupo', '')),
            "tarefa": str(row.get('Tarefa', '')),
            "data_trabalho": data_str,
            "principal": str(row.get('Principal', '')),
            "ajudante": str(row.get('Ajudante', '')),
            "data_registro": str(row.get('Data de Registro', get_horario_brasilia()))
        })
        proximo_id += 1
        
    if novos_registros:
        df_novos = pd.DataFrame(novos_registros)
        if df_global.empty:
            df_global = df_novos
        else:
            df_global = pd.concat([df_global, df_novos], ignore_index=True)
        salvar_aba(df_global, "historico")

def atualizar_historico_completo_db(user_id, df_historico_completo):
    df_global = carregar_aba("historico")
    if not df_global.empty and "user_id" in df_global.columns:
        df_global = df_global[df_global["user_id"] != int(user_id)]
        
    novos_registros = []
    proximo_id = 1 if df_global.empty or "id" not in df_global.columns else int(df_global["id"].max()) + 1
    
    for _, row in df_historico_completo.iterrows():
        dt_val = row['Data de Trabalho']
        if isinstance(dt_val, pd.Timestamp) or isinstance(dt_val, datetime):
            data_str = dt_val.strftime("%Y-%m-%d")
        else:
            try: data_str = pd.to_datetime(dt_val, dayfirst=True).strftime("%Y-%m-%d")
            except: data_str = str(dt_val)

        novos_registros.append({
            "id": proximo_id,
            "user_id": int(user_id),
            "grupo": str(row.get('Grupo', '')),
            "tarefa": str(row.get('Tarefa', '')),
            "data_trabalho": data_str,
            "principal": str(row.get('Principal', '')),
            "ajudante": str(row.get('Ajudante', '')),
            "data_registro": str(row.get('Data de Registro', get_horario_brasilia()))
        })
        proximo_id += 1
        
    if novos_registros:
        df_novos = pd.DataFrame(novos_registros)
        if df_global.empty:
            df_global = df_novos
        else:
            df_global = pd.concat([df_global, df_novos], ignore_index=True)
            
    salvar_aba(df_global, "historico")

def carregar_historico_db(user_id):
    df_global = carregar_aba("historico")
    if df_global.empty or "user_id" not in df_global.columns:
        return normalizar_tabela(pd.DataFrame())
        
    df_user = df_global[df_global["user_id"] == int(user_id)].copy()
    if df_user.empty:
        return normalizar_tabela(pd.DataFrame())
        
    df_user = df_user.rename(columns={
        "grupo": "Grupo", "tarefa": "Tarefa", "data_trabalho": "Data de Trabalho",
        "principal": "Principal", "ajudante": "Ajudante", "data_registro": "Data de Registro"
    })
    
    return normalizar_tabela(df_user[["Grupo", "Tarefa", "Data de Trabalho", "Principal", "Ajudante", "Data de Registro"]])

# --- FUNÇÕES EXCLUSIVAS DO ADMIN ---
def listar_todos_usuarios():
    df_usuarios = carregar_aba("usuarios")
    if df_usuarios.empty:
        return pd.DataFrame(columns=["id", "Nome", "Email"])
    
    df_exibir = df_usuarios[["id", "nome", "email"]].copy()
    df_exibir["email"] = df_exibir["email"].astype(str).str.strip()
    
    is_admin = df_exibir["email"].str.lower() == EMAIL_ADMIN.lower()
    df_exibir.loc[is_admin, "nome"] = "👑 " + df_exibir.loc[is_admin, "nome"]
    
    return df_exibir.rename(columns={"nome": "Nome", "email": "Email"})

def listar_todo_historico_admin():
    df_historico = carregar_aba("historico")
    df_usuarios = carregar_aba("usuarios")
    
    if df_historico.empty or df_usuarios.empty:
        return pd.DataFrame()
        
    df_merge = df_historico.merge(df_usuarios, left_on="user_id", right_on="id", suffixes=('_hist', '_user'))
    df_merge = df_merge.sort_values(by="id_hist", ascending=False)
    
    return df_merge[["nome", "email", "grupo", "tarefa", "data_trabalho", "principal", "ajudante", "data_registro"]].rename(
        columns={
            "nome": "Usuário", "email": "Email", "grupo": "Grupo", "tarefa": "Tarefa",
            "data_trabalho": "Data do Evento", "principal": "Principal", "ajudante": "Ajudante", "data_registro": "Salvo em"
        }
    )

def deletar_usuario_admin(user_id):
    df_usuarios = carregar_aba("usuarios")
    df_historico = carregar_aba("historico")
    
    if not df_usuarios.empty:
        user_para_deletar = df_usuarios[df_usuarios["id"] == int(user_id)]
        if not user_para_deletar.empty:
            email_alvo = str(user_para_deletar.iloc[0]["email"]).strip().lower()
            if email_alvo == EMAIL_ADMIN.lower() or int(user_id) == 1:
                return False
            
        df_usuarios = df_usuarios[df_usuarios["id"] != int(user_id)]
        salvar_aba(df_usuarios, "usuarios")
    if not df_historico.empty:
        df_historico = df_historico[df_historico["user_id"] != int(user_id)]
        salvar_aba(df_historico, "historico")
    return True

def atualizar_nome_usuario_admin(user_id, novo_nome):
    df_usuarios = carregar_aba("usuarios")
    if not df_usuarios.empty:
        idx = df_usuarios[df_usuarios["id"] == int(user_id)].index
        if len(idx) > 0:
            df_usuarios.loc[idx, "nome"] = novo_nome.strip()
            salvar_aba(df_usuarios, "usuarios")

def salvar_grupos_db(user_id, grupos_dict):
    df_usuarios = carregar_aba("usuarios")
    if not df_usuarios.empty:
        idx = df_usuarios[df_usuarios["id"] == int(user_id)].index
        if len(idx) > 0:
            df_usuarios.loc[idx, "grupos_json"] = json.dumps(grupos_dict)
            salvar_aba(df_usuarios, "usuarios")

@st.dialog("🔒 Autenticação Restrita")
def pop_up_senha_adm():
    st.write("Por favor, insira a Senha do ADM para acessar o painel administrativo.")
    senha_inserida = st.text_input("Senha do ADM:", type="password")
    
    if st.button("Confirmar Acesso", type="primary", use_container_width=True):
        if senha_inserida == "adm01":
            st.session_state.view_mode = "admin"
            st.rerun()
        else:
            st.error("Senha incorreta! Acesso negado.")

def gerar_escala_sem_repeticao(membros):
    if len(membros) < 2:
        return None
        
    principais = membros.copy()
    random.shuffle(principais)
    ajudantes = membros.copy()
    
    tentativas = 0
    while tentativas < 1000:
        random.shuffle(ajudantes)
        valido = True
        for p, a in zip(principais, ajudantes):
            if p == a:
                valido = False
                break
        if valido:
            break
        tentativas += 1
        
    scale_data = []
    for p, a in zip(principais, ajudantes):
        scale_data.append({
            "Principal": p,
            "Ajudante": a,
            "Data de Trabalho": pd.Timestamp("today").normalize() 
        })
    return pd.DataFrame(scale_data)

# --- 🚀 GERENCIAMENTO DE COOKIES E LOGIN COM AGENDAMENTO ---
cookie_manager = stx.CookieManager(key="gerenciador_cookies")

# 1. PROCESSA SALVAMENTO DE COOKIE AGENDADO (Dribla o bug do st.rerun)
if "set_cookie_now" in st.session_state:
    uid = st.session_state.pop("set_cookie_now")
    validade = datetime.now() + timedelta(days=30)
    cookie_manager.set("user_id", str(uid), expires_at=validade)

# 2. PROCESSA EXCLUSÃO DE COOKIE AGENDADA
if "delete_cookie_now" in st.session_state:
    st.session_state.pop("delete_cookie_now")
    cookie_manager.delete("user_id")

# 3. VERIFICA SE O USUÁRIO JÁ TEM COOKIE SALVO PARA LOGIN AUTOMÁTICO
cookie_user_id = cookie_manager.get(cookie="user_id")
if cookie_user_id is not None and st.session_state.user_id is None and not st.session_state.deslogado:
    usuario = buscar_usuario_por_id(int(cookie_user_id))
    if usuario:
        st.session_state.user_id = int(usuario["id"])
        st.session_state.user_nome = usuario["nome"]
        st.session_state.user_email = usuario["email"]
        st.session_state.grupos = json.loads(usuario["grupos_json"]) if usuario["grupos_json"] else {}
        st.session_state.historico_definitivo = carregar_historico_db(st.session_state.user_id)
        st.session_state.historico_carregado = True
        st.rerun()

if st.session_state.user_id is not None and not st.session_state.historico_carregado:
    st.session_state.historico_definitivo = carregar_historico_db(st.session_state.user_id)
    st.session_state.historico_carregado = True

# --- INTERFACE: TELA DESLOGADO ---
if st.session_state.user_id is None:
    st.title("🔐 Bem-vindo ao Sistema de Designações")
    st.markdown("Faça login ou crie uma conta. Todos os seus dados serão armazenados de forma segura na nuvem.")
    
    aba_login, aba_cadastro = st.tabs(["Entrar", "Criar Conta"])
    
    with aba_login:
        st.subheader("Fazer Login")
        email_login = st.text_input("Email", key="login_email")
        senha_login = st.text_input("Senha", type="password", key="login_senha")
        manter_logado = st.checkbox("Lembrar meu acesso", value=True)
        
        if st.button("Entrar", type="primary", use_container_width=True):
            usuario = verificar_login(email_login, senha_login)
            if usuario:
                st.session_state.user_id = int(usuario["id"])
                st.session_state.user_nome = usuario["nome"]
                st.session_state.user_email = usuario["email"]
                st.session_state.grupos = json.loads(usuario["grupos_json"]) if usuario["grupos_json"] else {}
                st.session_state.historico_definitivo = carregar_historico_db(st.session_state.user_id)
                st.session_state.historico_carregado = True
                st.session_state.deslogado = False 
                
                if manter_logado:
                    # 💡 AQUI ESTÁ A MÁGICA: Agendamos o cookie para a próxima renderização da página!
                    st.session_state.set_cookie_now = usuario["id"]
                    
                st.rerun()
            else:
                st.error("Email ou senha incorretos.")
                
    with aba_cadastro:
        st.subheader("Novo Cadastro")
        novo_nome = st.text_input("Nome Completo", key="cad_nome")
        novo_email = st.text_input("Email", key="cad_email")
        nova_senha = st.text_input("Senha", type="password", key="cad_senha")
        
        if st.button("Cadastrar", type="secondary", use_container_width=True):
            if novo_nome and novo_email and nova_senha:
                resultado = cadastrar_usuario(novo_nome, novo_email, nova_senha)
                if resultado == "sucesso":
                    st.success("Conta criada com sucesso! Faça login na aba ao lado.")
                elif resultado == "email_duplicado":
                    st.error("⚠️ Este email já está cadastrado.")
                elif resultado == "nome_duplicado":
                    st.error("⚠️ Este nome já está em uso.")
                else:
                    st.error("Erro técnico de comunicação com o Google. Tente novamente.")
            else:
                st.warning("Preencha todos os campos.")

# --- INTERFACE: TELA LOGADO ---
else:
    with st.sidebar:
        st.title(f"Olá, {st.session_state.user_nome} 👋")
        
        email_atual_limpo = st.session_state.user_email.strip().lower() if st.session_state.user_email else ""
        if email_atual_limpo == EMAIL_ADMIN.strip().lower():
            st.write("---")
            if st.session_state.view_mode == "app":
                if st.button("🛠️ Sala do Adm", use_container_width=True, type="primary"):
                    pop_up_senha_adm() 
            else:
                if st.button("🏠 Voltar para o Sistema", use_container_width=True, type="primary"):
                    st.session_state.view_mode = "app"
                    st.rerun()
            st.write("---")
            
        if st.button("🚪 Sair (Logout)", use_container_width=True):
            # 💡 Agendando a exclusão do cookie com segurança
            st.session_state.delete_cookie_now = True
            st.session_state.deslogado = True
            st.session_state.user_id = None
            st.session_state.user_nome = None
            st.session_state.user_email = None
            st.session_state.historico_carregado = False
            st.session_state.historico_definitivo = normalizar_tabela(pd.DataFrame())
            st.rerun()

    # --- PAINEL ADMINISTRADOR ---
    if st.session_state.view_mode == "admin":
        st.title("🛠️ Painel Administrativo Cloud")
        aba_admin_usuarios, aba_admin_historico = st.tabs(["👥 Gerenciar Usuários", "🌎 Ver Histórico Global"])
        
        with aba_admin_usuarios:
            df_usuarios = listar_todos_usuarios()
            if not df_usuarios.empty:
                st.write(f"**Total de usuários cadastrados na nuvem:** {len(df_usuarios)}")
                st.dataframe(df_usuarios, use_container_width=True, hide_index=True)
                
                st.write("---")
                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("✏️ Editar Nome")
                    user_id_editar = st.selectbox("Selecione o Usuário:", df_usuarios["Nome"] + " (" + df_usuarios["Email"] + ")", key="sel_edit")
                    novo_nome_admin = st.text_input("Novo Nome:")
                    if st.button("Salvar Alteração"):
                        if novo_nome_admin.strip():
                            idx = df_usuarios[df_usuarios["Nome"] + " (" + df_usuarios["Email"] + ")" == user_id_editar].index[0]
                            atualizar_nome_usuario_admin(int(df_usuarios.loc[idx, "id"]), novo_nome_admin)
                            st.success("Nome updated!")
                            st.rerun()
                with col2:
                    st.subheader("🗑️ Remover Usuário")
                    user_id_remover = st.selectbox("Selecione o Usuário a ser apagado:", df_usuarios["Nome"] + " (" + df_usuarios["Email"] + ")", key="sel_del")
                    if st.button("Confirmar Exclusão", type="primary"):
                        idx = df_usuarios[df_usuarios["Nome"] + " (" + df_usuarios["Email"] + ")" == user_id_remover].index[0]
                        id_real = int(df_usuarios.loc[idx, "id"])
                        email_real = df_usuarios.loc[idx, "Email"].strip().lower()
                        
                        if email_real == EMAIL_ADMIN.strip().lower() or id_real == 1:
                            st.error("❌ Ação Bloqueada! Por segurança, você não pode excluir o Administrador Master.")
                        else:
                            sucesso = deletar_usuario_admin(id_real)
                            if sucesso:
                                st.success("Usuário removido!")
                                st.rerun()
                            else:
                                st.error("Erro ao tentar remover o usuário.")
            else:
                st.info("Nenhum usuário cadastrado.")
                
        with aba_admin_historico:
            df_hist_global = listar_todo_historico_admin()
            if not df_hist_global.empty:
                st.dataframe(df_hist_global, use_container_width=True, hide_index=True)
            else:
                st.info("Nenhum histórico salvo na nuvem ainda.")

    # --- VISÃO SISTEMA PADRÃO ---
    else:
        st.title("👥 Sistema Gestão de Designações")
        aba_gerador, aba_historico, aba_grupos = st.tabs(["🗓️ Gerar e Editar Escala", "📊 Histórico Consolidado", "🗂️ Gestão de Grupos"])

        with aba_gerador:
            st.header("1. Configurar Nova Atividade")
            if not st.session_state.grupos:
                st.warning("Você não tem nenhum grupo cadastrado. Vá até a aba 'Gestão de Grupos' para criar um.")
            else:
                col1, col2 = st.columns([1, 1])
                with col1: grupo_selecionado = st.selectbox("Selecione o Grupo:", list(st.session_state.grupos.keys()))
                with col2: tarefa_nome = st.text_input("Nome da Atividade:", value="Coordenação do Evento")
                    
                if st.button("🔄 Gerar Sugestão de Duplas", type="primary", use_container_width=True):
                    membros = st.session_state.grupos[grupo_selecionado]
                    if len(membros) < 2: st.error("O grupo precisa ter pelo menos 2 pessoas.")
                    else:
                        st.session_state.escala_temporaria = normalizar_tabela(gerar_escala_sem_repeticao(membros))
                        st.toast("Sugestão de duplas gerada com sucesso!", icon="💡")

                if st.session_state.escala_temporaria is not None:
                    st.write("---")
                    st.subheader("✍️ 2. Área Editável: Ajuste, Adicione ou Remova Duplas")
                    
                    escala_editada = st.data_editor(
                        st.session_state.escala_temporaria,
                        column_config={
                            "Principal": st.column_config.TextColumn("🧑‍✈️ Principal (Líder)", required=True),
                            "Ajudante": st.column_config.TextColumn("🧑‍🔧 Ajudante", required=True),
                            "Data de Trabalho": st.column_config.DateColumn("📅 Data de Trabalho", format="DD/MM/YYYY", required=True)
                        },
                        num_rows="dynamic", hide_index=True, use_container_width=True
                    )
                    
                    if st.button("💾 Confirmar e Registrar no Histórico", type="secondary", use_container_width=True):
                        df_registro = escala_editada.dropna(subset=["Principal", "Ajudante"]).copy()
                        if df_registro.empty: st.warning("A tabela está vazia.")
                        else:
                            df_registro["Grupo"] = grupo_selecionado
                            df_registro["Tarefa"] = tarefa_nome
                            df_registro["Data de Registro"] = get_horario_brasilia()
                            
                            df_registro = normalizar_tabela(df_registro[["Grupo", "Tarefa", "Data de Trabalho", "Principal", "Ajudante", "Data de Registro"]])
                            
                            acrescentar_historico_db(st.session_state.user_id, df_registro)
                            
                            st.session_state.historico_definitivo = normalizar_tabela(pd.concat([st.session_state.historico_definitivo, df_registro], ignore_index=True))
                            st.session_state.escala_temporaria = None
                            st.success("Escala salva com sucesso na nuvem por blocos!")
                            st.rerun()

        with aba_historico:
            st.header("📊 Seu Histórico de Atividades")
            if not st.session_state.historico_definitivo.empty:
                
                # --- MODO LEITURA (BLINDADO COM DEFENSIVE PROGRAMMING) ---
                df_exibir = normalizar_tabela(st.session_state.historico_definitivo).iloc[::-1].reset_index(drop=True)
                
                # Aplica a função de blindagem para evitar o AttributeError do .dt
                df_exibir["Data de Trabalho"] = formatar_data_segura(df_exibir["Data de Trabalho"])
                
                df_exibir_visual = df_exibir.rename(columns={
                    "Grupo": "🗂️ Grupo", "Tarefa": "📝 Tarefa", "Data de Trabalho": "📅 Data de Trabalho",
                    "Principal": "🧑‍✈️ Principal", "Ajudante": "🧑‍🔧 Ajudante", "Data de Registro": "🕰️ Salvo Em"
                })
                
                cores_pasteis = ['#E8F4F8', '#FFF3CD', '#D1E7DD', '#F8D7DA', '#E2E3E5', '#F3D8F4']
                unique_timestamps = df_exibir_visual['🕰️ Salvo Em'].unique()
                map_cores = {ts: cores_pasteis[i % len(cores_pasteis)] for i, ts in enumerate(unique_timestamps)}
                
                def colorir_fundo(row):
                    cor = map_cores.get(row['🕰️ Salvo Em'], '#FFFFFF')
                    return [f'background-color: {cor}; color: #000000; font-weight: 500;'] * len(row)
                
                try:
                    df_estilizado = df_exibir_visual.style.apply(colorir_fundo, axis=1).hide(axis="index")
                except AttributeError:
                    df_estilizado = df_exibir_visual.style.apply(colorir_fundo, axis=1).hide_index()

                st.dataframe(df_estilizado, use_container_width=True)
                
                # --- MODO EDIÇÃO AVANÇADA ---
                with st.expander("🛠️ Ativar Modo de Edição Manual (Painel Branco)"):
                    st.info("💡 Como a tabela acima foi blindada para aceitar as cores, use este painel se precisar alterar dados manualmente.")
                    
                    df_editated = st.data_editor(
                        df_exibir_visual,
                        column_config={
                            "🕰️ Salvo Em": st.column_config.Column(disabled=True)
                        },
                        use_container_width=True, hide_index=True, key="editor_historico_definitivo"
                    )
                    
                    df_editated_db = df_editated.rename(columns={
                        "🗂️ Grupo": "Grupo", "📝 Tarefa": "Tarefa", "📅 Data de Trabalho": "Data de Trabalho",
                        "🧑‍✈️ Principal": "Principal", "🧑‍🔧 Ajudante": "Ajudante", "🕰️ Salvo Em": "Data de Registro"
                    })
                    df_exibir_db = df_exibir.rename(columns={
                        "🗂️ Grupo": "Grupo", "📝 Tarefa": "Tarefa", "📅 Data de Trabalho": "Data de Trabalho",
                        "🧑‍✈️ Principal": "Principal", "🧑‍🔧 Ajudante": "Ajudante", "🕰️ Salvo Em": "Data de Registro"
                    })

                    df_editated_seguro = normalizar_tabela(df_editated_db)
                    df_exibir_seguro = normalizar_tabela(df_exibir_db)

                    if not df_editated_seguro.equals(df_exibir_seguro):
                        df_para_salvar = df_editated_seguro.iloc[::-1].reset_index(drop=True)
                        st.session_state.historico_definitivo = normalizar_tabela(df_para_salvar)
                        atualizar_historico_completo_db(st.session_state.user_id, df_para_salvar)
                        st.success("Alteração manual salva na nuvem!")
                        st.rerun()
                
                st.write("---")
                col_nome, col_baixar, col_limpar = st.columns([2, 1, 1])
                with col_nome: nome_historico = st.text_input("Nome do arquivo de backup:", value="Escala_Geral")
                with col_baixar:
                    st.write(""); st.write("")
                    csv = st.session_state.historico_definitivo.to_csv(index=False).encode('utf-8')
                    st.download_button("📥 Baixar Backup (.csv)", data=csv, file_name=f"{nome_historico}.csv", mime='text/csv', use_container_width=True)
                with col_limpar:
                    st.write(""); st.write("")
                    if st.button("🗑️ Limpar Todo o Histórico", use_container_width=True):
                        st.session_state.historico_definitivo = normalizar_tabela(pd.DataFrame(columns=["Grupo", "Tarefa", "Data de Trabalho", "Principal", "Ajudante", "Data de Registro"]))
                        atualizar_historico_completo_db(st.session_state.user_id, st.session_state.historico_definitivo)
                        st.rerun()
            else: st.info("Nenhum registro confirmado.")

        with aba_grupos:
            st.header("🗂️ Gestão de Grupos e Integrantes")
            st.subheader("1. Gerenciar Grupos")
            col_add_grp, col_ren_grp, col_del_grp = st.columns(3)
            
            with col_add_grp:
                novo_grupo_nome = st.text_input("Criar Novo Grupo:", key="input_add_grp")
                if st.button("➕ Adicionar Grupo", use_container_width=True):
                    if novo_grupo_nome.strip() and novo_grupo_nome.strip() not in st.session_state.grupos:
                        st.session_state.grupos[novo_grupo_nome.strip()] = []
                        salvar_grupos_db(st.session_state.user_id, st.session_state.grupos) 
                        st.rerun()
            with col_ren_grp:
                if st.session_state.grupos:
                    grupo_a_renomear = st.selectbox("Grupo a Renomear:", list(st.session_state.grupos.keys()), key="select_ren_grp")
                    novo_nome_para_grupo = st.text_input("Novo Nome:", key="input_ren_grp")
                    if st.button("✏️ Confirmar Nome", use_container_width=True):
                        if novo_nome_para_grupo.strip() and novo_nome_para_grupo.strip() not in st.session_state.grupos:
                            st.session_state.grupos[novo_nome_para_grupo.strip()] = st.session_state.grupos.pop(grupo_a_renomear)
                            salvar_grupos_db(st.session_state.user_id, st.session_state.grupos) 
                            st.rerun()
            with col_del_grp:
                if st.session_state.grupos:
                    grupo_a_remover = st.selectbox("Excluir Grupo:", list(st.session_state.grupos.keys()), key="select_del_grp")
                    if st.button("🗑️ Excluir Grupo", use_container_width=True):
                        del st.session_state.grupos[grupo_a_remover]
                        salvar_grupos_db(st.session_state.user_id, st.session_state.grupos) 
                        st.rerun()

            st.write("---")
            st.subheader("2. Visualizar e Editar Integrantes")
            if not st.session_state.grupos: st.info("Você não possui grupos.")
            else:
                grupo_visualizar = st.radio("Selecione o grupo:", list(st.session_state.grupos.keys()), horizontal=True)
                lista_atual = st.session_state.grupos[grupo_visualizar]
                st.write(f"**Integrantes do grupo ({len(lista_atual)} pessoas):**")
                st.dataframe(pd.DataFrame(lista_atual, columns=["Nome do Integrante"]), use_container_width=True, hide_index=True)
                
                st.write("---")
                col_adicionar, col_remover = st.columns([1, 1])
                with col_adicionar:
                    st.markdown("#### ➕ Adicionar integrante:")
                    grupo_destino = st.selectbox("Escolha o grupo:", list(st.session_state.grupos.keys()), key="select_add_destino")
                    novo_nome = st.text_input("Nome da pessoa:", key="input_add_membro")
                    if st.button("Confirmar Adição", type="secondary", use_container_width=True):
                        if novo_nome.strip() and novo_nome.strip() not in st.session_state.grupos[grupo_destino]:
                            st.session_state.grupos[grupo_destino].append(novo_nome.strip())
                            salvar_grupos_db(st.session_state.user_id, st.session_state.grupos) 
                            st.rerun()
                with col_remover:
                    st.markdown("#### 🗑️ Remover integrante:")
                    grupo_origem = st.selectbox("Escolha o grupo:", list(st.session_state.grupos.keys()), key="select_del_destino")
                    lista_remocao = st.session_state.grupos[grupo_origem]
                    if len(lista_remocao) > 0:
                        selecionar_todos = st.checkbox("Selecionar todos", key="chk_sel_all")
                        nomes_remover = st.multiselect("Pessoas:", options=lista_remocao, default=lista_remocao if selecionar_todos else [])
                        if st.button("Confirmar Remoção", type="primary", use_container_width=True) and nomes_remover:
                            for nome in nomes_remover: st.session_state.grupos[grupo_origem].remove(nome)
                            salvar_grupos_db(st.session_state.user_id, st.session_state.grupos) 
                            st.rerun()

        # --- RODAPÉ ---
        st.write("---")
        st.markdown("""<div style="text-align: center; color: #888888; font-size: 12px; padding: 10px 0px;">Criado e atualizado por: Sérgio Sierra</div>""", unsafe_allow_html=True)