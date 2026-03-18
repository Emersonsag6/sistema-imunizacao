# ============================================================================
# SISTEMA COMPLETO DE GESTÃO DE IMUNIZAÇÃO - VERSÃO FINAL REFATORADA
# COM PROTOCOLOS DINÂMICOS, RELATÓRIOS E SEMÁFORO DE STATUS
# ============================================================================

# INSTALAÇÕES
# Servir arquivos estáticos (PWA)
import os
from pathlib import Path

# Criar diretório public se não existir
os.makedirs('public', exist_ok=True)

# Copiar manifest.json
manifest_content = """{
  "name": "Sistema de Imunização",
  "short_name": "Imunização",
  "description": "Sistema de Gestão de Imunização",
  "start_url": "/",
  "display": "standalone",
  "background_color": "#ffffff",
  "theme_color": "#4CAF50",
  "icons": [
    {
      "src": "data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 192 192'><rect fill='%234CAF50' width='192' height='192'/><text x='50%' y='50%' font-size='120' fill='white' text-anchor='middle' dominant-baseline='middle'>💉</text></svg>",
      "sizes": "192x192",
      "type": "image/svg+xml"
    }
  ]
}"""

with open('public/manifest.json', 'w') as f:
    f.write(manifest_content)

print("✅ PWA configurado!")

# ============================================================================
# SEÇÃO 1: LIMPEZA DE DADOS
# ============================================================================

class LimpadorDadosSujos:
    """Limpa planilhas sujas - busca inteligente de idade"""
    
    @staticmethod
    def normalizar_texto(texto):
        if not texto:
            return ""
        return unidecode(str(texto)).lower().strip()
    
    @staticmethod
    def remover_colunas_vazias(df):
        df = df.dropna(axis=1, how='all')
        return df
    
    @staticmethod
    def encontrar_linha_cabecalho(df_raw):
        print("\n🔍 PROCURANDO LINHA DE CABEÇALHO...")
        
        for idx in range(min(10, len(df_raw))):
            linha = df_raw.iloc[idx].astype(str).values
            texto_linha = " ".join(linha).lower()
            
            palavras_chave = ['nome', 'paciente', 'cpf', 'data', 'idade', 'aldeia', 'sexo', 'telefone', 'indio', 'índio']
            match_count = sum(1 for palavra in palavras_chave if palavra in texto_linha)
            
            if match_count >= 2:
                print(f"  ✅ Cabeçalho encontrado na linha {idx}")
                return idx
        
        return 0
    
    @staticmethod
    def processar_excel_sujo(arquivo_path):
        print("\n" + "="*70)
        print("📥 PROCESSANDO EXCEL SUJO")
        print("="*70)
        
        try:
            df_raw = pd.read_excel(arquivo_path, header=None)
            print(f"\n📊 Excel bruto: {df_raw.shape[0]} linhas x {df_raw.shape[1]} colunas")
            
            idx_cabecalho = LimpadorDadosSujos.encontrar_linha_cabecalho(df_raw)
            
            print(f"\n📖 Relendo com cabeçalho na linha {idx_cabecalho}...")
            df = pd.read_excel(arquivo_path, header=idx_cabecalho)
            
            df = LimpadorDadosSujos.remover_colunas_vazias(df)
            
            print(f"✅ Após limpeza: {df.shape[0]} linhas x {df.shape[1]} colunas")
            
            return df
        
        except Exception as e:
            print(f"❌ Erro ao processar: {e}")
            return None
    
    @staticmethod
    def encontrar_coluna_por_palavra_chave(df, palavras_chave_lista):
        melhor_match = None
        melhor_score = 0
        
        for col in df.columns:
            col_norm = LimpadorDadosSujos.normalizar_texto(col)
            
            if col_norm.startswith('unnamed') or col_norm in ['unnamed: 0', 'índice', 'index']:
                continue
            
            for palavra in palavras_chave_lista:
                palavra_norm = LimpadorDadosSujos.normalizar_texto(palavra)
                score = fuzz.ratio(col_norm, palavra_norm)
                
                if score > melhor_score:
                    melhor_score = score
                    melhor_match = col
        
        if melhor_score > 60:
            return melhor_match
        return None
    
    @staticmethod
    def encontrar_coluna_nome(df):
        print("\n🔎 PROCURANDO COLUNA DE NOME...")
        
        palavras_nome = [
            'nome', 'paciente', 'person', 'full name', 'nome completo', 'name', 
            'nomedopaciente', 'nome indio', 'nome do índio', 'nome indígena', 'indio', 'índio'
        ]
        
        resultado = LimpadorDadosSujos.encontrar_coluna_por_palavra_chave(df, palavras_nome)
        
        if resultado:
            print(f"✅ Coluna NOME encontrada: '{resultado}'")
            return resultado
        
        return None
    
    @staticmethod
    def encontrar_colunas_relevantes(df):
        print("\n📋 MAPEANDO COLUNAS...")
        
        mapa = {}
        
        mapa['nome'] = LimpadorDadosSujos.encontrar_coluna_nome(df)
        
        cpf_palavras = ['cpf', 'documento', 'doc', 'rg', 'identificação', 'identific']
        mapa['cpf'] = LimpadorDadosSujos.encontrar_coluna_por_palavra_chave(df, cpf_palavras)
        
        data_palavras = ['data', 'nasc', 'nascimento', 'birth', 'dob', 'data_nasc', 'data nasc', 'datanasc']
        mapa['data_nascimento'] = LimpadorDadosSujos.encontrar_coluna_por_palavra_chave(df, data_palavras)
        
        idade_palavras = ['idade', 'age', 'anos', 'year', 'idadepaciente', 'idade do paciente']
        mapa['idade'] = LimpadorDadosSujos.encontrar_coluna_por_palavra_chave(df, idade_palavras)
        
        sexo_palavras = ['sexo', 'gender', 'gênero', 'sex', 'genero', 'masculino', 'feminino']
        mapa['sexo'] = LimpadorDadosSujos.encontrar_coluna_por_palavra_chave(df, sexo_palavras)
        
        tel_palavras = ['telefone', 'phone', 'celular', 'tel', 'contact', 'contato', 'telfone']
        mapa['telefone'] = LimpadorDadosSujos.encontrar_coluna_por_palavra_chave(df, tel_palavras)
        
        aldeia_palavras = ['aldeia', 'comunidade', 'village', 'cidade', 'localidade', 'terra', 'povo', 'município']
        mapa['aldeia'] = LimpadorDadosSujos.encontrar_coluna_por_palavra_chave(df, aldeia_palavras)
        
        return mapa
    
    @staticmethod
    def limpar_cpf(cpf):
        if not cpf or str(cpf).lower() == 'nan':
            return ''
        cpf_clean = re.sub(r'\D', '', str(cpf))
        if len(cpf_clean) == 11:
            return f"{cpf_clean[0:3]}.{cpf_clean[3:6]}.{cpf_clean[6:9]}-{cpf_clean[9:11]}"
        return cpf_clean
    
    @staticmethod
    def limpar_data(data):
        if not data or str(data).lower() == 'nan':
            return ''
        
        data_str = str(data).strip()
        formatos = ['%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y/%m/%d', '%m/%d/%Y', 
                   '%Y%m%d', '%d.%m.%Y', '%Y.%m.%d', '%d/%m/%y']
        
        for fmt in formatos:
            try:
                data_obj = datetime.strptime(data_str, fmt)
                return data_obj.strftime('%Y-%m-%d')
            except:
                continue
        
        return data_str
    
    @staticmethod
    def limpar_sexo(sexo):
        if not sexo or str(sexo).lower() == 'nan':
            return ''
        sexo_str = str(sexo).strip().upper()
        return 'M' if sexo_str.startswith('M') else 'F' if sexo_str.startswith('F') else ''
    
    @staticmethod
    def limpar_nome(nome):
        if not nome or str(nome).lower() == 'nan':
            return ''
        nome_str = str(nome).strip()
        nome_str = re.sub(r'\s+', ' ', nome_str)
        return nome_str.title()
    
    @staticmethod
    def extrair_idade(valor_idade_coluna, data_nascimento, mapa_colunas):
        if valor_idade_coluna and str(valor_idade_coluna).lower() != 'nan':
            idade_str = str(valor_idade_coluna).strip()
            idade_str = re.sub(r'[\sa,]*(anos|a|years|year)?[\s]*', '', idade_str, flags=re.IGNORECASE)
            numeros = re.findall(r'\d+', idade_str)
            
            if numeros:
                try:
                    idade = int(numeros[0])
                    if 0 < idade < 150:
                        return str(idade)
                except:
                    pass
        
        if data_nascimento:
            try:
                data_nasc = datetime.strptime(data_nascimento, '%Y-%m-%d').date()
                hoje = date.today()
                idade = hoje.year - data_nasc.year - ((hoje.month, hoje.day) < (data_nasc.month, data_nasc.day))
                if 0 < idade < 150:
                    return str(idade)
            except:
                pass
        
        return ''
    
    @staticmethod
    def processar_dataframe(df, db):
        print("\n" + "="*70)
        print("🔄 PROCESSANDO DADOS")
        print("="*70)
        
        df = df.dropna(how='all')
        
        mapa_colunas = LimpadorDadosSujos.encontrar_colunas_relevantes(df)
        
        if not mapa_colunas.get('nome'):
            print("\n❌ ERRO: Coluna de NOME não identificada!")
            return [], [f"❌ Coluna de NOME não encontrada"]
        
        pacientes_processados = []
        erros = []
        
        print(f"\n📥 IMPORTANDO PACIENTES...")
        
        for idx, row in df.iterrows():
            try:
                nome = LimpadorDadosSujos.limpar_nome(row.get(mapa_colunas['nome']))
                
                if not nome:
                    continue
                
                cpf = LimpadorDadosSujos.limpar_cpf(row.get(mapa_colunas['cpf'])) if mapa_colunas.get('cpf') else ''
                data_nasc = LimpadorDadosSujos.limpar_data(row.get(mapa_colunas['data_nascimento'])) if mapa_colunas.get('data_nascimento') else ''
                sexo = LimpadorDadosSujos.limpar_sexo(row.get(mapa_colunas['sexo'])) if mapa_colunas.get('sexo') else ''
                telefone = str(row.get(mapa_colunas['telefone'], '')).strip() if mapa_colunas.get('telefone') else ''
                aldeia = str(row.get(mapa_colunas['aldeia'], '')).strip() if mapa_colunas.get('aldeia') else ''
                
                valor_idade_col = row.get(mapa_colunas['idade']) if mapa_colunas.get('idade') else None
                idade_str = LimpadorDadosSujos.extrair_idade(valor_idade_col, data_nasc, mapa_colunas)
                
                dados_originais = {str(col): str(row.get(col, '')) for col in df.columns}
                
                paciente = db.criar_paciente(nome, cpf, data_nasc, sexo, telefone, aldeia, dados_originais, idade_str)
                
                if paciente:
                    pacientes_processados.append(paciente)
                    if len(pacientes_processados) <= 3:
                        print(f"  ✅ {nome} - D/N: {data_nasc} - Idade: {idade_str}")
                
            except Exception as e:
                print(f"Erro na linha {idx}: {str(e)}")
                erros.append(str(e))
        
        print(f"\n✅ RESULTADO: {len(pacientes_processados)} pacientes importados")
        
        return pacientes_processados, erros

limpador = LimpadorDadosSujos()

# ============================================================================
# SEÇÃO 2: GERENCIADOR DE PROTOCOLOS
# ============================================================================

class GerenciadorProtocolos:
    """Gerencia protocolos vacinais em JSON com CRUD completo"""
    
    ARQUIVO_PROTOCOLO = "protocolo_vacinal.json"
    
    PROTOCOLO_PADRAO = {
        "protocolos": [
            {"id": 1, "imunizante": "BCG", "idade_minima_meses": 0, "idade_maxima_meses": 59, "esquema_doses": "1 dose", "intervalo_dias": 0, "tipo": "Dose Única", "faixa_etaria": "Ao nascer (menores de 5 anos)", "via_administracao": "ID (0,05-0,1 ml, região deltóide)"},
            {"id": 2, "imunizante": "Hepatite B", "idade_minima_meses": 0, "idade_maxima_meses": 1, "esquema_doses": "1 dose", "intervalo_dias": 0, "tipo": "Dose Única", "faixa_etaria": "Ao nascer / 0-1 mês", "via_administracao": "IM (0,5-1 ml, dependendo idade)"},
            {"id": 3, "imunizante": "VIP (Poliomielite)", "idade_minima_meses": 2, "idade_maxima_meses": 6, "esquema_doses": "3 doses", "intervalo_dias": 60, "tipo": "Esquema Sequencial", "faixa_etaria": "2m, 4m, 6m", "via_administracao": "IM (0,5 ml, vasto lateral coxa)"},
            {"id": 4, "imunizante": "VOP (Poliomielite reforço)", "idade_minima_meses": 15, "idade_maxima_meses": 48, "esquema_doses": "Reforço", "intervalo_dias": 0, "tipo": "Reforço", "faixa_etaria": "15m, 4a", "via_administracao": "Oral (2 gotas)"},
            {"id": 5, "imunizante": "Pentavalente", "idade_minima_meses": 2, "idade_maxima_meses": 6, "esquema_doses": "3 doses + reforços", "intervalo_dias": 60, "tipo": "Esquema Sequencial", "faixa_etaria": "2m, 4m, 6m", "via_administracao": "IM (0,5 ml, vasto lateral coxa)"},
            {"id": 6, "imunizante": "Pneumocócica 10-valente", "idade_minima_meses": 2, "idade_maxima_meses": 12, "esquema_doses": "3 doses + reforço", "intervalo_dias": 60, "tipo": "Esquema Sequencial", "faixa_etaria": "2m, 4m, reforço 12m", "via_administracao": "IM (0,5 ml, vasto lateral coxa)"},
            {"id": 7, "imunizante": "Rotavírus (VORH)", "idade_minima_meses": 2, "idade_maxima_meses": 4, "esquema_doses": "2 doses", "intervalo_dias": 60, "tipo": "Esquema Sequencial", "faixa_etaria": "2m, 4m", "via_administracao": "Oral (1,5 ml)"},
            {"id": 8, "imunizante": "Influenza", "idade_minima_meses": 6, "idade_maxima_meses": 600, "esquema_doses": "1-2 doses", "intervalo_dias": 365, "tipo": "Anual", "faixa_etaria": "A partir de 6m (dose anual)", "via_administracao": "IM (0,25-0,5 ml)"},
            {"id": 9, "imunizante": "Febre Amarela", "idade_minima_meses": 9, "idade_maxima_meses": 600, "esquema_doses": "1 dose", "intervalo_dias": 0, "tipo": "Dose Única", "faixa_etaria": "A partir de 9m", "via_administracao": "SC (0,5 ml)"},
            {"id": 10, "imunizante": "Tríplice Viral (SCR)", "idade_minima_meses": 12, "idade_maxima_meses": 15, "esquema_doses": "2 doses", "intervalo_dias": 90, "tipo": "Esquema Sequencial", "faixa_etaria": "12m, 15m", "via_administracao": "SC (0,5 ml)"},
            {"id": 11, "imunizante": "DTP (reforço)", "idade_minima_meses": 15, "idade_maxima_meses": 72, "esquema_doses": "Reforços", "intervalo_dias": 365, "tipo": "Reforço", "faixa_etaria": "15m, 4-6a", "via_administracao": "IM (0,5 ml, vasto lateral coxa)"},
            {"id": 12, "imunizante": "Tetra Viral", "idade_minima_meses": 15, "idade_maxima_meses": 15, "esquema_doses": "1 dose", "intervalo_dias": 0, "tipo": "Dose Única", "faixa_etaria": "15m (após SCR 12m)", "via_administracao": "SC (0,5 ml)"},
            {"id": 13, "imunizante": "Hepatite A", "idade_minima_meses": 15, "idade_maxima_meses": 15, "esquema_doses": "1 dose", "intervalo_dias": 0, "tipo": "Dose Única", "faixa_etaria": "15m", "via_administracao": "IM (0,5 ml)"},
            {"id": 14, "imunizante": "Varicela (monovalente)", "idade_minima_meses": 48, "idade_maxima_meses": 72, "esquema_doses": "1 reforço", "intervalo_dias": 0, "tipo": "Reforço", "faixa_etaria": "4-6a", "via_administracao": "SC (0,5 ml)"},
            {"id": 15, "imunizante": "Meningocócica C", "idade_minima_meses": 3, "idade_maxima_meses": 12, "esquema_doses": "2 doses + reforço", "intervalo_dias": 60, "tipo": "Esquema Sequencial", "faixa_etaria": "3m, 5m, reforço 12m", "via_administracao": "IM (0,5 ml, vasto lateral coxa)"},
            {"id": 16, "imunizante": "Meningocócica ACWY", "idade_minima_meses": 132, "idade_maxima_meses": 168, "esquema_doses": "1 dose", "intervalo_dias": 0, "tipo": "Dose Única", "faixa_etaria": "11-14a", "via_administracao": "IM (0,5 ml, deltoide)"},
            {"id": 17, "imunizante": "HPV Quadrivalente", "idade_minima_meses": 108, "idade_maxima_meses": 168, "esquema_doses": "2 doses", "intervalo_dias": 180, "tipo": "Esquema Sequencial", "faixa_etaria": "9-14a", "via_administracao": "IM (0,5 ml)"},
            {"id": 18, "imunizante": "dT (dupla adulto)", "idade_minima_meses": 84, "idade_maxima_meses": 600, "esquema_doses": "3 doses + reforços", "intervalo_dias": 60, "tipo": "Esquema Sequencial", "faixa_etaria": "A partir de 7a", "via_administracao": "IM (0,5 ml)"},
            {"id": 19, "imunizante": "dTpa", "idade_minima_meses": 180, "idade_maxima_meses": 600, "esquema_doses": "1 dose por gestação", "intervalo_dias": 0, "tipo": "Dose Única", "faixa_etaria": "Gestantes, profissionais", "via_administracao": "IM (0,5 ml)"},
            {"id": 20, "imunizante": "Pneumocócica 23-valente", "idade_minima_meses": 60, "idade_maxima_meses": 600, "esquema_doses": "1 dose", "intervalo_dias": 0, "tipo": "Dose Única", "faixa_etaria": "A partir de 5a", "via_administracao": "IM (0,5 ml)"},
            {"id": 21, "imunizante": "COVID-19", "idade_minima_meses": 6, "idade_maxima_meses": 600, "esquema_doses": "3 doses + reforços", "intervalo_dias": 28, "tipo": "Esquema Sequencial", "faixa_etaria": "A partir de 6m", "via_administracao": "IM (0,3-0,5 ml)"}
        ]
    }
    
    def __init__(self):
        self.carregar_protocolo()
    
    def carregar_protocolo(self):
        if os.path.exists(self.ARQUIVO_PROTOCOLO):
            try:
                with open(self.ARQUIVO_PROTOCOLO, 'r', encoding='utf-8') as f:
                    self.dados = json.load(f)
            except:
                self.dados = self.PROTOCOLO_PADRAO.copy()
                self.salvar_protocolo()
        else:
            self.dados = self.PROTOCOLO_PADRAO.copy()
            self.salvar_protocolo()
    
    def salvar_protocolo(self):
        try:
            with open(self.ARQUIVO_PROTOCOLO, 'w', encoding='utf-8') as f:
                json.dump(self.dados, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"❌ Erro ao salvar protocolo: {e}")
    
    def obter_protocolos(self):
        return self.dados.get("protocolos", [])
    
    def obter_protocolo_por_id(self, protocolo_id):
        for p in self.obter_protocolos():
            if p['id'] == int(protocolo_id):
                return p
        return None
    
    def obter_protocolo_por_vacina(self, vacina_nome):
        for p in self.obter_protocolos():
            if p['imunizante'].lower() == vacina_nome.lower():
                return p
        return None
    
    def adicionar_protocolo(self, imunizante, idade_min, idade_max, esquema, intervalo, tipo, faixa, via_admin):
        novo_id = max([p['id'] for p in self.obter_protocolos()], default=0) + 1
        
        novo_protocolo = {
            "id": novo_id,
            "imunizante": imunizante,
            "idade_minima_meses": int(idade_min),
            "idade_maxima_meses": int(idade_max),
            "esquema_doses": esquema,
            "intervalo_dias": int(intervalo),
            "tipo": tipo,
            "faixa_etaria": faixa,
            "via_administracao": via_admin
        }
        
        self.dados["protocolos"].append(novo_protocolo)
        self.salvar_protocolo()
        return novo_protocolo
    
    def atualizar_protocolo(self, protocolo_id, **kwargs):
        for p in self.dados["protocolos"]:
            if p['id'] == int(protocolo_id):
                p.update(kwargs)
                self.salvar_protocolo()
                return p
        return None
    
    def deletar_protocolo(self, protocolo_id):
        self.dados["protocolos"] = [p for p in self.dados["protocolos"] if p['id'] != int(protocolo_id)]
        self.salvar_protocolo()
    
    def obter_vacinas_por_idade(self, idade_meses):
        recomendadas = []
        
        for protocolo in self.obter_protocolos():
            if protocolo['idade_minima_meses'] <= idade_meses <= protocolo['idade_maxima_meses']:
                recomendadas.append(protocolo['imunizante'])
        
        return recomendadas
    
    def obter_intervalo_vacina(self, imunizante):
        for p in self.obter_protocolos():
            if p['imunizante'].lower() == imunizante.lower():
                return p['intervalo_dias']
        return 30

gerenciador_protocolos = GerenciadorProtocolos()

# ============================================================================
# SEÇÃO 3: CALENDÁRIO VACINAL
# ============================================================================

class CalendarioVacinal:
    
    @staticmethod
    def calcular_idade_em_meses(data_nascimento):
        if not data_nascimento:
            return 0
        
        try:
            data_nasc = datetime.strptime(data_nascimento, '%Y-%m-%d').date()
            hoje = date.today()
            
            meses = (hoje.year - data_nasc.year) * 12 + (hoje.month - data_nasc.month)
            return max(0, meses)
        except:
            return 0
    
    @staticmethod
    def calcular_proxima_dose(vacina, dose_atual, data_aplicacao, idade_meses=0):
        
        try:
            data_app = datetime.strptime(data_aplicacao, '%Y-%m-%d').date()
            intervalo = gerenciador_protocolos.obter_intervalo_vacina(vacina)
            
            protocolo = gerenciador_protocolos.obter_protocolo_por_vacina(vacina)
            if protocolo and protocolo['tipo'] == 'Dose Única':
                return "ESQUEMA COMPLETO", date.today()
            
            if dose_atual == "D.U (Dose Única)":
                proxima_dose = "ESQUEMA COMPLETO"
                data_proxima = data_app
            elif dose_atual == "D1 (1ª Dose)":
                proxima_dose = "D2 (2ª Dose)"
                data_proxima = data_app + timedelta(days=intervalo)
            elif dose_atual == "D2 (2ª Dose)":
                proxima_dose = "D3 (3ª Dose)"
                data_proxima = data_app + timedelta(days=intervalo)
            elif dose_atual == "D3 (3ª Dose)":
                proxima_dose = "Reforço"
                data_proxima = data_app + timedelta(days=180)
            elif dose_atual == "Reforço":
                proxima_dose = "Reforço 2"
                data_proxima = data_app + timedelta(days=365)
            else:
                proxima_dose = "Consultar protocolo"
                data_proxima = data_app + timedelta(days=30)
            
            return proxima_dose, data_proxima
        
        except Exception as e:
            return f"Erro: {str(e)}", date.today()
    
    @staticmethod
    def obter_status_vacina(vacina, idade_meses, vacinacoes_paciente):
        
        protocolo = gerenciador_protocolos.obter_protocolo_por_vacina(vacina)
        
        if not protocolo:
            return "⚪ SEM PROTOCOLO", ""
        
        idade_min = protocolo['idade_minima_meses']
        idade_max = protocolo['idade_maxima_meses']
        
        doses_registradas = [v for v in vacinacoes_paciente if v.get('vacina', '').lower() == vacina.lower()]
        
        if doses_registradas:
            ultima_dose = doses_registradas[-1]
            proxima_dose, data_proxima = CalendarioVacinal.calcular_proxima_dose(
                vacina, 
                ultima_dose.get('dose', ''), 
                ultima_dose.get('data', datetime.now().strftime('%Y-%m-%d')),
                idade_meses
            )
            
            if proxima_dose == "ESQUEMA COMPLETO":
                return "🟢 VACINADO", f"Esquema completo"
            else:
                return "🟢 VACINADO", f"Próxima: {proxima_dose} ({data_proxima.strftime('%d/%m/%Y')})"
        
        if idade_meses < idade_min:
            return "⚪ NÃO INDICADO", f"Aguarde {idade_min} meses"
        elif idade_meses <= idade_max:
            if idade_meses >= (idade_max * 0.8):
                return "🟡 ATENÇÃO", f"Próxima: D1 (1ª Dose) - LIMITE DE IDADE"
            else:
                return "🟢 NO PRAZO", f"Próxima: D1 (1ª Dose)"
        else:
            return "🔴 ATRASADA", f"Próxima: D1 (1ª Dose)"
    
    @staticmethod
    def validar_vacina_para_idade(vacina, idade_meses):
        
        protocolo = gerenciador_protocolos.obter_protocolo_por_vacina(vacina)
        
        if not protocolo:
            return True, "Protocolo não encontrado"
        
        if idade_meses < protocolo['idade_minima_meses']:
            return False, f"⚠️ Paciente muito jovem. Idade mínima: {protocolo['idade_minima_meses']} meses"
        
        if idade_meses > protocolo['idade_maxima_meses']:
            return False, f"⚠️ Paciente fora da faixa etária. Idade máxima: {protocolo['idade_maxima_meses']} meses"
        
        return True, "✅ Válido para esta idade"

# ============================================================================
# SEÇÃO 4: BANCO DE DADOS COMPLETO
# ============================================================================

class BancoDadosCompleto:
    def __init__(self):
        self.pacientes = []
        self.vacinacoes = []
        self.campanhas = []
        self.lotes = []
        self.insumos_estoque = []
        self.dados_preenchimento = {}
        
        self.proximo_id_paciente = 1
        self.proximo_id_vacinacao = 1
        self.proximo_id_campanha = 1
        self.proximo_id_lote = 1
        self.proximo_id_insumo = 1
        
        self.campanha_ativa = None
        
        self.carregar_historico()
        self.criar_campanhas_padrao()
    
    def criar_campanhas_padrao(self):
        if len(self.campanhas) == 0:
            print("🚀 Criando campanhas padrão do protocolo...")
            for protocolo in gerenciador_protocolos.obter_protocolos():
                self.criar_campanha(
                    f"Campanha {protocolo['imunizante']}",
                    protocolo['imunizante'],
                    f"LOTE-{protocolo['id']:03d}",
                    "Fornecedor Padrão",
                    "2025-12-31",
                    "Admin",
                    "1"
                )
            print(f"✅ {len(self.campanhas)} campanhas criadas!")
    
    def criar_paciente(self, nome, cpf='', data_nasc='', sexo='', telefone='', aldeia='', dados_originais=None, idade_importada=''):
        if not nome:
            return None
        
        paciente = {
            'id': self.proximo_id_paciente,
            'nome': nome,
            'cpf': cpf,
            'data_nascimento': data_nasc,
            'sexo': sexo,
            'telefone': telefone,
            'aldeia': aldeia,
            'idade_importada': idade_importada,
            'dados_completos': dados_originais or {},
            'data_cadastro': datetime.now().isoformat()
        }
        self.pacientes.append(paciente)
        self.proximo_id_paciente += 1
        return paciente
    
    def obter_paciente(self, paciente_id):
        try:
            for p in self.pacientes:
                if p['id'] == int(paciente_id):
                    return p
        except:
            pass
        return None
    
    def buscar_paciente_por_nome(self, nome_query):
        nome_norm = LimpadorDadosSujos.normalizar_texto(nome_query)
        resultados = []
        
        for p in self.pacientes:
            nome_pac_norm = LimpadorDadosSujos.normalizar_texto(p['nome'])
            if nome_norm in nome_pac_norm or nome_pac_norm in nome_norm:
                resultados.append(p)
        
        return resultados
    
    def calcular_idade(self, data_nascimento):
        if not data_nascimento:
            return "N/A"
        try:
            data_nasc = datetime.strptime(data_nascimento, '%Y-%m-%d').date()
            hoje = date.today()
            idade = hoje.year - data_nasc.year - ((hoje.month, hoje.day) < (data_nasc.month, data_nasc.day))
            return str(idade)
        except:
            return "N/A"
    
    def listar_pacientes(self):
        return self.pacientes
    
    def get_total_pacientes(self):
        return len(self.pacientes)
    
    def limpar_pacientes(self):
        self.pacientes = []
        self.proximo_id_paciente = 1
    
    def registrar_vacinacao(self, paciente_id, vacina, dose, data, lote, fabricante='', profissional='', 
                           local_injecao='', via_admin='', reacao_adversa='', obs='', prox_dose='', tipo_campanha=''):
        vacinacao = {
            'id': self.proximo_id_vacinacao,
            'paciente_id': int(paciente_id),
            'paciente_nome': '',
            'paciente_cpf': '',
            'vacina': vacina,
            'dose': dose,
            'data': data,
            'lote': lote,
            'fabricante': fabricante,
            'profissional': profissional,
            'local_injecao': local_injecao,
            'via_administracao': via_admin,
            'reacao_adversa': reacao_adversa,
            'observacoes': obs,
            'proxima_dose_prevista': prox_dose,
            'tipo_campanha': tipo_campanha,
            'data_criacao': datetime.now().isoformat()
        }
        
        paciente = self.obter_paciente(paciente_id)
        if paciente:
            vacinacao['paciente_nome'] = paciente['nome']
            vacinacao['paciente_cpf'] = paciente['cpf']
        
        self.vacinacoes.append(vacinacao)
        self.proximo_id_vacinacao += 1
        
        for l in self.lotes:
            if l['numero_lote'] == lote:
                l['quantidade_utilizada'] += 1
                l['saldo_disponivel'] = l['quantidade_recebida'] - l['quantidade_utilizada']
        
        for i in self.insumos_estoque:
            i['quantidade_aplicada'] += 1
            i['saldo_real'] = i['quantidade_recebida'] - i['quantidade_aplicada']
        
        self.salvar_historico()
        
        return vacinacao
    
    def listar_vacinacoes_paciente(self, paciente_id):
        return [v for v in self.vacinacoes if v['paciente_id'] == int(paciente_id)]
    
    def get_total_vacinacoes(self):
        return len(self.vacinacoes)
    
    def criar_campanha(self, nome, vacina, lote, fabricante, validade, profissional, dose='1'):
        campanha = {
            'id': self.proximo_id_campanha,
            'nome': nome,
            'vacina': vacina,
            'lote': lote,
            'fabricante': fabricante,
            'validade': validade,
            'profissional': profissional,
            'dose': dose,
            'ativa': True,
            'data_criacao': datetime.now().isoformat()
        }
        
        self.campanhas.append(campanha)
        self.campanha_ativa = campanha
        self.proximo_id_campanha += 1
        return campanha
    
    def ativar_campanha(self, campanha_id):
        for c in self.campanhas:
            if c['id'] == campanha_id:
                c['ativa'] = True
                self.campanha_ativa = c
                return c
        return None
    
    def obter_campanha_ativa(self):
        return self.campanha_ativa
    
    def listar_campanhas(self):
        return self.campanhas
    
    def deletar_campanha(self, campanha_id):
        self.campanhas = [c for c in self.campanhas if c['id'] != campanha_id]
    
    def atualizar_campanha(self, campanha_id, **kwargs):
        for c in self.campanhas:
            if c['id'] == int(campanha_id):
                c.update(kwargs)
                return c
        return None
    
    def atualizar_paciente(self, paciente_id, **kwargs):
        """Atualiza dados do paciente"""
        for p in self.pacientes:
            if p['id'] == int(paciente_id):
                p.update(kwargs)
                return p
        return None
    
    def adicionar_lote(self, nome_vacina, numero_lote, fabricante='', apresentacao='', 
                      qtd=0, validade='', recebimento='', obs=''):
        lote = {
            'id': self.proximo_id_lote,
            'nome_vacina': nome_vacina,
            'numero_lote': numero_lote,
            'fabricante': fabricante,
            'apresentacao': apresentacao,
            'quantidade_recebida': int(qtd) if qtd else 0,
            'quantidade_utilizada': 0,
            'quantidade_descartada': 0,
            'saldo_disponivel': int(qtd) if qtd else 0,
            'data_validade': validade,
            'data_recebimento': recebimento,
            'observacoes': obs,
            'status': 'Ativo',
            'data_criacao': datetime.now().isoformat()
        }
        self.lotes.append(lote)
        self.proximo_id_lote += 1
        return lote
    
    def listar_lotes(self):
        return self.lotes
    
    def adicionar_insumo(self, tipo, descricao, qtd_solicitada, qtd_recebida, 
                        ml_seringa='', calibre_agulha='', lote='', validade=''):
        insumo = {
            'id': self.proximo_id_insumo,
            'tipo': tipo,
            'descricao': descricao,
            'quantidade_solicitada': int(qtd_solicitada) if qtd_solicitada else 0,
            'quantidade_recebida': int(qtd_recebida) if qtd_recebida else 0,
            'divergencia': (int(qtd_recebida) if qtd_recebida else 0) - (int(qtd_solicitada) if qtd_solicitada else 0),
            'ml_seringa': ml_seringa,
            'calibre_agulha': calibre_agulha,
            'lote': lote,
            'validade': validade,
            'quantidade_aplicada': 0,
            'saldo_real': int(qtd_recebida) if qtd_recebida else 0,
            'data_criacao': datetime.now().isoformat()
        }
        self.insumos_estoque.append(insumo)
        self.proximo_id_insumo += 1
        return insumo
    
    def listar_insumos(self):
        return self.insumos_estoque
    
    def salvar_historico(self):
        try:
            df = pd.DataFrame(self.vacinacoes)
            df.to_csv('historico_vacinacoes.csv', index=False, encoding='utf-8')
        except Exception as e:
            print(f"❌ Erro ao salvar histórico: {e}")
    
    def carregar_historico(self):
        try:
            if os.path.exists('historico_vacinacoes.csv'):
                df = pd.read_csv('historico_vacinacoes.csv')
                self.vacinacoes = df.to_dict('records')
                if self.vacinacoes:
                    self.proximo_id_vacinacao = max([v.get('id', 0) for v in self.vacinacoes], default=0) + 1
        except Exception as e:
            print(f"⚠️ Aviso ao carregar histórico: {e}")

db = BancoDadosCompleto()

# ============================================================================
# SEÇÃO 5: FUNÇÕES GRADIO
# ============================================================================

def importar_excel(arquivo):
    try:
        if arquivo is None:
            return "❌ Nenhum arquivo selecionado!", None
        
        df = limpador.processar_excel_sujo(arquivo.name)
        
        if df is None:
            return "❌ Erro ao ler arquivo", None
        
        db.limpar_pacientes()
        pacientes, erros = limpador.processar_dataframe(df, db)
        
        msg = f"""✅ IMPORTAÇÃO CONCLUÍDA!

📊 Resultado:
  • Pacientes importados: {len(pacientes)}
  • Total no banco: {db.get_total_pacientes()}
  • Erros: {len(erros)}
        """
        
        if db.pacientes:
            dados = []
            for p in db.pacientes:
                idade = p['idade_importada'] or db.calcular_idade(p['data_nascimento'])
                data_nasc = p['data_nascimento'] or '-'
                dados.append([p['nome'], p['cpf'] or '-', data_nasc, idade, p['aldeia'] or '-'])
            return msg, gr.Dataframe(value=dados)
        
        return msg, None
    
    except Exception as e:
        print(f"ERRO IMPORTAR: {str(e)}")
        return f"❌ Erro: {str(e)}", None

def buscar_paciente_realtime(nome_query):
    """Busca em tempo real"""
    if not nome_query or len(nome_query) < 1:
        return pd.DataFrame(columns=["Nome", "CPF", "Data Nascimento", "Idade", "Aldeia"])
    
    try:
        resultados = db.buscar_paciente_por_nome(nome_query)
        
        dados = []
        for p in resultados:
            idade = p['idade_importada'] or db.calcular_idade(p['data_nascimento'])
            data_nasc = p['data_nascimento'] or '-'
            dados.append([p['nome'], p['cpf'] or '-', data_nasc, idade, p['aldeia'] or '-'])
        
        return pd.DataFrame(dados, columns=["Nome", "CPF", "Data Nascimento", "Idade", "Aldeia"]) if dados else pd.DataFrame(columns=["Nome", "CPF", "Data Nascimento", "Idade", "Aldeia"])
    
    except:
        return pd.DataFrame(columns=["Nome", "CPF", "Data Nascimento", "Idade", "Aldeia"])

def converter_data_para_iso(data_str):
    """Converte DD-MM-AAAA para YYYY-MM-DD"""
    try:
        data_obj = datetime.strptime(data_str, '%d-%m-%Y')
        return data_obj.strftime("%Y-%m-%d")
    except:
        return ""

def converter_data_para_br(data_iso):
    """Converte YYYY-MM-DD para DD-MM-AAAA"""
    try:
        data_obj = datetime.strptime(data_iso, '%Y-%m-%d')
        return data_obj.strftime("%d-%m-%Y")
    except:
        return data_iso

def preencher_vacinacao_ao_clicar(evt: gr.SelectData, tabela_dados):
    """Preenche dados ao clicar na tabela"""
    try:
        if tabela_dados is None or len(tabela_dados) == 0:
            return ("❌ Tabela vazia", "", "", "", "", "", "", "", "", "", 4)
        
        index = evt.index[0]
        df = pd.DataFrame(tabela_dados, columns=["Nome", "CPF", "Data Nascimento", "Idade", "Aldeia"])
        linha = df.iloc[index]
        
        nome_p = str(linha['Nome'])
        cpf_p = str(linha['CPF'])
        data_nasc_p = str(linha['Data Nascimento'])
        idade_p = str(linha['Idade'])
        
        camp = db.obter_campanha_ativa()
        v_nome, v_lote, v_val, v_fab, v_prof = ("", "", "", "", "")
        
        if camp:
            v_nome = f"{camp['vacina']} (Lote: {camp['lote']})"
            v_lote = camp['lote']
            v_val = camp['validade']
            v_fab = camp['fabricante']
            v_prof = camp['profissional']
        
        data_agora = datetime.now().strftime("%d-%m-%Y")
        
        db.dados_preenchimento = {
            "nome": nome_p,
            "cpf": cpf_p,
            "data_nascimento": data_nasc_p,
            "idade": idade_p,
            "vacina": v_nome,
            "lote": v_lote,
            "validade": v_val,
            "fabricante": v_fab,
            "profissional": v_prof,
            "data": data_agora
        }
        
        return (
            "✅ Dados carregados! Pulando para aba VACINAÇÃO...",
            nome_p,
            cpf_p,
            data_nasc_p,
            idade_p,
            v_nome,
            v_lote,
            v_val,
            v_fab,
            v_prof,
            4
        )
        
    except Exception as e:
        print(f"❌ Erro: {e}")
        return (f"❌ Erro ao carregar dados: {str(e)}", "", "", "", "", "", "", "", "", "", 3)

def adicionar_paciente_manual(nome, cpf, data_nasc, idade, aldeia):
    """Adiciona paciente manualmente"""
    try:
        if not nome:
            return "❌ Nome é obrigatório!"
        
        # Normaliza data
        if data_nasc:
            data_nasc = converter_data_para_iso(data_nasc)
        
        paciente = db.criar_paciente(nome, cpf, data_nasc, '', '', aldeia, None, idade)
        
        if paciente:
            return f"✅ Paciente '{nome}' adicionado com sucesso!"
        else:
            return "❌ Erro ao adicionar paciente!"
    
    except Exception as e:
        return f"❌ Erro: {str(e)}"

def editar_dados_paciente_atual(nome_atual, novo_nome, novo_cpf, nova_idade):
    """Edita dados do paciente que está em preenchimento"""
    try:
        # Encontra o paciente
        paciente = None
        for p in db.pacientes:
            if p['nome'].lower() == nome_atual.lower():
                paciente = p
                break
        
        if not paciente:
            return "❌ Paciente não encontrado!"
        
        # Atualiza dados
        updates = {}
        if novo_nome and novo_nome != nome_atual:
            updates['nome'] = novo_nome
        if novo_cpf:
            updates['cpf'] = novo_cpf
        if nova_idade:
            updates['idade_importada'] = str(nova_idade)
        
        if updates:
            db.atualizar_paciente(paciente['id'], **updates)
            return f"✅ Dados do paciente atualizados!"
        
        return "⚠️ Nenhuma alteração realizada"
    
    except Exception as e:
        return f"❌ Erro: {str(e)}"

def obter_vacinas_disponiveis():
    """Retorna lista de vacinas"""
    vacinas = []
    
    if db.campanhas:
        for c in db.campanhas:
            vacinas.append(f"{c['vacina']} (Lote: {c['lote']})")
        
        vacinas = list(dict.fromkeys(vacinas))
        return vacinas
    else:
        return ["➕ Nenhuma campanha cadastrada"]

def atualizar_dados_vacinacao(vacina_selecionada):
    """Preenche lote, validade, fabricante e VIA"""
    if not vacina_selecionada or "Nenhuma" in vacina_selecionada:
        return "", "", "", ""
    
    vacina_nome = vacina_selecionada.split(" (Lote:")[0]
    
    for c in db.campanhas:
        if c['vacina'] == vacina_nome and c['ativa']:
            protocolo = gerenciador_protocolos.obter_protocolo_por_vacina(vacina_nome)
            via = protocolo.get('via_administracao', '') if protocolo else ''
            return c['lote'], converter_data_para_br(c['validade']), c['fabricante'], via
    
    for protocolo in gerenciador_protocolos.obter_protocolos():
        if protocolo['imunizante'] == vacina_nome:
            return "", "", "", protocolo.get('via_administracao', '')
    
    return "", "", "", ""

def calcular_proxima_dose_automatica(vacina, dose_selecionada, idade):
    """Calcula próxima dose automaticamente"""
    if not vacina or not dose_selecionada or not idade:
        return ""
    
    try:
        idade_int = int(str(idade).replace(" anos", "").strip())
        proxima_dose, data_proxima = CalendarioVacinal.calcular_proxima_dose(
            vacina, dose_selecionada, datetime.now().strftime("%Y-%m-%d"), idade_int
        )
        return f"{proxima_dose} - {data_proxima.strftime('%d/%m/%Y')}"
    except:
        return ""

def obter_status_vacina_com_proxima(vacina, idade_anos, paciente_id):
    """Retorna status da vacina com próxima dose conforme protocolo"""
    if not vacina or not idade_anos or not paciente_id:
        return "⚪ SEM DADOS", ""
    
    try:
        vacina_clean = vacina.split(" (Lote:")[0] if " (Lote:" in vacina else vacina
        idade_meses = int(float(idade_anos)) * 12
        
        vacinacoes_paciente = db.listar_vacinacoes_paciente(int(paciente_id))
        
        status, proxima = CalendarioVacinal.obter_status_vacina(vacina_clean, idade_meses, vacinacoes_paciente)
        return status, proxima
    except:
        return "⚪ ERRO NO CÁLCULO", ""

def validar_vacina_idade(vacina, idade_anos):
    """Valida se vacina é apropriada para idade"""
    if not vacina or not idade_anos:
        return "⚪ SEM DADOS"
    
    try:
        vacina_clean = vacina.split(" (Lote:")[0] if " (Lote:" in vacina else vacina
        idade_meses = int(float(idade_anos)) * 12
        valido, mensagem = CalendarioVacinal.validar_vacina_para_idade(vacina_clean, idade_meses)
        return mensagem
    except:
        return "❌ Erro na validação"

def registrar_vacinacao(paciente_nome, vacina, dose, data_br, lote, fabricante, profissional, 
                       local_injecao, via_admin, reacao, obs, prox_dose, tipo_campanha):
    try:
        if not paciente_nome or not vacina or not data_br or not lote or not tipo_campanha:
            return "❌ Preencha TODOS os campos obrigatórios!"
        
        if not local_injecao:
            return "❌ Selecione Local de Injeção!"
        
        if not via_admin:
            return "❌ Selecione Via de Administração!"
        
        try:
            data_iso = converter_data_para_iso(data_br)
        except:
            return "❌ Data inválida! Use o formato DD-MM-AAAA"
        
        vacina_nome = vacina.split(" (Lote:")[0] if " (Lote:" in vacina else vacina
        
        for p in db.pacientes:
            if p['nome'].lower() == paciente_nome.lower():
                db.registrar_vacinacao(
                    p['id'], vacina_nome, dose,
                    data_iso, lote, fabricante, profissional, local_injecao, via_admin, reacao, obs, prox_dose, tipo_campanha
                )
                return f"✅ Vacinação registrada para {paciente_nome}! ({tipo_campanha})"
        
        return "❌ Paciente não encontrado!"
    
    except Exception as e:
        return f"❌ Erro: {str(e)}"

def criar_campanha_config(nome, vacina, lote, fabricante, validade, profissional, dose):
    """Cria campanha e atualiza lista"""
    try:
        if not nome or not vacina or not lote or not profissional:
            return "❌ Preencha: Nome, Vacina, Lote e Profissional!", pd.DataFrame(columns=["ID", "Nome", "Vacina", "Lote", "Fabricante", "Profissional", "Validade", "Dose", "Status"]), "", "", "", "", "", ""
        
        db.criar_campanha(nome, vacina, lote, fabricante, validade, profissional, dose)
        
        dados = []
        for c in db.campanhas:
            status = "🟢 ATIVA" if c['ativa'] else "⚪ Inativa"
            dados.append([c['id'], c['nome'], c['vacina'], c['lote'], c['fabricante'], c['profissional'], converter_data_para_br(c['validade']), c['dose'], status])
        
        return (
            f"✅ Campanha '{nome}' criada e ATIVADA!",
            pd.DataFrame(dados, columns=["ID", "Nome", "Vacina", "Lote", "Fabricante", "Profissional", "Validade", "Dose", "Status"]),
            "",
            "",
            "",
            "",
            "",
            ""
        )
    
    except Exception as e:
        return f"❌ Erro: {str(e)}", pd.DataFrame(columns=["ID", "Nome", "Vacina", "Lote", "Fabricante", "Profissional", "Validade", "Dose", "Status"]), "", "", "", "", "", ""

def ativar_campanha_config(evt: gr.SelectData, tabela_dados):
    try:
        if tabela_dados is None or len(tabela_dados) == 0:
            return "❌ Selecione uma campanha!"
        
        index = evt.index[0]
        df = pd.DataFrame(tabela_dados, columns=["ID", "Nome", "Vacina", "Lote", "Fabricante", "Profissional", "Validade", "Dose", "Status"])
        campanha_id = int(df.iloc[index]['ID'])
        
        db.ativar_campanha(campanha_id)
        
        return f"✅ CAMPANHA ATIVADA!"
        
    except Exception as e:
        return f"❌ Erro: {str(e)}"

def deletar_campanha(campanha_id):
    try:
        if not campanha_id:
            return "❌ Digite um ID válido!", pd.DataFrame(columns=["ID", "Nome", "Vacina", "Lote", "Fabricante", "Profissional", "Validade", "Dose", "Status"])
        
        db.deletar_campanha(int(campanha_id))
        
        dados = []
        for c in db.campanhas:
            status = "🟢 ATIVA" if c['ativa'] else "⚪ Inativa"
            dados.append([c['id'], c['nome'], c['vacina'], c['lote'], c['fabricante'], c['profissional'], converter_data_para_br(c['validade']), c['dose'], status])
        
        return f"✅ Campanha deletada!", pd.DataFrame(dados, columns=["ID", "Nome", "Vacina", "Lote", "Fabricante", "Profissional", "Validade", "Dose", "Status"])
    
    except Exception as e:
        return f"❌ Erro: {str(e)}", pd.DataFrame(columns=["ID", "Nome", "Vacina", "Lote", "Fabricante", "Profissional", "Validade", "Dose", "Status"])

def listar_campanhas_config():
    """Lista campanhas para edição"""
    dados = []
    for c in db.campanhas:
        status = "🟢 ATIVA" if c['ativa'] else "⚪ Inativa"
        dados.append([c['id'], c['nome'], c['vacina'], c['lote'], c['fabricante'], c['profissional'], converter_data_para_br(c['validade']), c['dose'], status])
    
    return pd.DataFrame(dados, columns=["ID", "Nome", "Vacina", "Lote", "Fabricante", "Profissional", "Validade", "Dose", "Status"]) if dados else pd.DataFrame(columns=["ID", "Nome", "Vacina", "Lote", "Fabricante", "Profissional", "Validade", "Dose", "Status"])

def atualizar_campanha_editar(campanha_id, campo, novo_valor):
    """Atualiza campo específico da campanha"""
    try:
        if not campanha_id or not campo:
            return "❌ Campos inválidos!"
        
        campos_validos = ['nome', 'vacina', 'lote', 'fabricante', 'profissional', 'validade']
        
        if campo not in campos_validos:
            return f"❌ Campo '{campo}' inválido!"
        
        # Converte validade se necessário
        if campo == 'validade' and novo_valor:
            novo_valor = converter_data_para_iso(novo_valor)
        
        db.atualizar_campanha(int(campanha_id), **{campo: novo_valor})
        return f"✅ Campo '{campo}' atualizado com sucesso!"
    
    except Exception as e:
        return f"❌ Erro: {str(e)}"

def buscar_pacientes_faltosos():
    vacinados_nomes = set(v['paciente_id'] for v in db.vacinacoes)
    faltosos = [p for p in db.pacientes if p['id'] not in vacinados_nomes]
    
    dados = []
    for p in faltosos:
        idade = p['idade_importada'] or db.calcular_idade(p['data_nascimento'])
        data_nasc = p['data_nascimento'] or '-'
        dados.append([p['nome'], p['cpf'] or '-', data_nasc, idade, p['aldeia'] or '-'])
    
    return pd.DataFrame(dados, columns=["Nome", "CPF", "Data Nascimento", "Idade", "Aldeia"]) if dados else pd.DataFrame(columns=["Nome", "CPF", "Data Nascimento", "Idade", "Aldeia"])

def adicionar_lote_estoque(vacina, lote, fabricante, qtd, validade):
    try:
        if not vacina or not lote or not qtd:
            return "❌ Preencha: Vacina, Lote e Quantidade!"
        
        db.adicionar_lote(vacina, lote, fabricante, "", int(qtd), validade, "", "")
        return f"✅ Lote {lote} adicionado!"
    
    except Exception as e:
        return f"❌ Erro: {str(e)}"

def listar_estoque():
    if not db.lotes:
        return None
    
    dados = []
    for l in db.lotes:
        dias_venc = ""
        if l['data_validade']:
            try:
                data_venc = datetime.strptime(l['data_validade'], '%Y-%m-%d').date()
                dias_restantes = (data_venc - date.today()).days
                if dias_restantes < 0:
                    dias_venc = f"🔴 VENCIDO"
                elif dias_restantes < 30:
                    dias_venc = f"🟡 {dias_restantes}d"
                else:
                    dias_venc = f"🟢 OK"
            except:
                dias_venc = converter_data_para_br(l['data_validade'])
        
        dados.append([
            l['nome_vacina'],
            l['numero_lote'],
            l['fabricante'],
            l['quantidade_recebida'],
            l['quantidade_utilizada'],
            l['saldo_disponivel'],
            dias_venc
        ])
    
    return pd.DataFrame(dados, columns=["Vacina", "Lote", "Fabricante", "Recebida", "Utilizada", "Saldo", "Validade"])

def listar_insumos_estoque():
    if not db.insumos_estoque:
        return None
    
    dados = []
    for i in db.insumos_estoque:
        status = "⚠️" if i['divergencia'] != 0 else "✅"
        dados.append([
            i['tipo'],
            i['descricao'],
            i['quantidade_solicitada'],
            i['quantidade_recebida'],
            f"{status} {i['divergencia']:+d}",
            i['ml_seringa'] or '-',
            i['calibre_agulha'] or '-',
            i['lote'],
            i['validade']
        ])
    
    return pd.DataFrame(dados, columns=["Tipo", "Descrição", "Solic.", "Receb.", "Diferença", "Seringa", "Calibre", "Lote", "Validade"])

def adicionar_insumo_estoque(tipo, descricao, qtd_solic, qtd_receb, ml_seringa, calibre, lote, validade):
    try:
        if not tipo or not descricao:
            return "❌ Preencha tipo e descrição!"
        
        db.adicionar_insumo(tipo, descricao, qtd_solic or 0, qtd_receb or 0, 
                           ml_seringa, calibre, lote, validade)
        return f"✅ Insumo '{descricao}' adicionado!"
    
    except Exception as e:
        return f"❌ Erro: {str(e)}"

def gerar_consolidado():
    total_vac = db.get_total_vacinacoes()
    saldo_doses = sum(l['saldo_disponivel'] for l in db.lotes)
    saldo_ins = sum(i['saldo_real'] for i in db.insumos_estoque)
    
    dados = []
    
    for l in db.lotes:
        dados.append([
            f"Vacina: {l['nome_vacina']}",
            "-",
            l['quantidade_recebida'],
            "-",
            l['quantidade_utilizada'],
            l['saldo_disponivel']
        ])
    
    for i in db.insumos_estoque:
        dif = i['divergencia']
        dif_str = f"⚠️ {dif}" if dif != 0 else str(dif)
        dados.append([
            i['descricao'],
            i['quantidade_solicitada'],
            i['quantidade_recebida'],
            dif_str,
            i['quantidade_aplicada'],
            i['saldo_real']
        ])
    
    return (
        int(total_vac),
        int(saldo_doses),
        int(saldo_ins),
        pd.DataFrame(dados, columns=["Insumo", "Solicitado", "Recebido", "Diferença", "Aplicado", "Saldo Real"]) if dados else pd.DataFrame(columns=["Insumo", "Solicitado", "Recebido", "Diferença", "Aplicado", "Saldo Real"])
    )

# ============================================================================
# CORREÇÃO: EXPORTAÇÃO DE ARQUIVOS EXCEL
# ============================================================================

def exportar_consolidado_completo_excel():
    """Exporta TUDO em um único arquivo Excel com múltiplas abas"""
    try:
        arquivo_saida = BytesIO()
        
        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            
            # ABA 1: HISTÓRICO DE VACINADOS
            if db.vacinacoes:
                df_vacinacoes = pd.DataFrame([
                    [v.get('paciente_nome', ''), v.get('paciente_cpf', ''), v.get('vacina', ''), 
                     v.get('dose', ''), converter_data_para_br(v.get('data', '')), v.get('lote', ''),
                     v.get('fabricante', ''), v.get('profissional', ''), v.get('local_injecao', ''),
                     v.get('via_administracao', ''), v.get('tipo_campanha', '')]
                    for v in db.vacinacoes
                ], columns=['Paciente', 'CPF', 'Vacina', 'Dose', 'Data', 'Lote', 'Fabricante', 
                           'Profissional', 'Local', 'Via', 'Campanha'])
            else:
                df_vacinacoes = pd.DataFrame(columns=['Paciente', 'CPF', 'Vacina', 'Dose', 'Data', 'Lote', 'Fabricante', 
                                                     'Profissional', 'Local', 'Via', 'Campanha'])
            df_vacinacoes.to_excel(writer, sheet_name='Historico Vacinados', index=False)
            
            # ABA 2: ESTOQUE
            df_estoque = listar_estoque()
            if df_estoque is not None and not df_estoque.empty:
                df_estoque.to_excel(writer, sheet_name='Estoque', index=False)
            else:
                pd.DataFrame(columns=["Vacina", "Lote", "Fabricante", "Recebida", "Utilizada", "Saldo", "Validade"]).to_excel(writer, sheet_name='Estoque', index=False)
            
            # ABA 3: PENDENTES/FALTOSOS
            df_faltosos = buscar_pacientes_faltosos()
            if not df_faltosos.empty:
                df_faltosos.to_excel(writer, sheet_name='Pendentes', index=False)
            else:
                pd.DataFrame(columns=["Nome", "CPF", "Data Nascimento", "Idade", "Aldeia"]).to_excel(writer, sheet_name='Pendentes', index=False)
            
            # ABA 4: RELATÓRIO
            df_relatorio = gerar_relatorio_historico()
            df_relatorio.to_excel(writer, sheet_name='Relatorio', index=False)
            
            # ABA 5: CONSOLIDADO
            total_vac, saldo_doses, saldo_ins, df_consolidado = gerar_consolidado()
            df_consolidado.to_excel(writer, sheet_name='Consolidado', index=False)
        
        arquivo_saida.seek(0)
        return arquivo_saida
    
    except Exception as e:
        print(f"ERRO CONSOLIDADO: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def exportar_historico_excel():
    """Exporta histórico de vacinados em Excel"""
    try:
        arquivo_saida = BytesIO()
        
        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            df_historico = listar_historico_vacinacoes()
            if not df_historico.empty:
                df_historico.to_excel(writer, sheet_name='Historico Vacinados', index=False)
            else:
                df_historico.to_excel(writer, sheet_name='Historico Vacinados', index=False)
        
        arquivo_saida.seek(0)
        return arquivo_saida
    
    except Exception as e:
        print(f"ERRO HISTÓRICO: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def exportar_estoque_excel():
    """Exporta estoque em Excel"""
    try:
        arquivo_saida = BytesIO()
        
        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            df_estoque = listar_estoque()
            if df_estoque is not None and not df_estoque.empty:
                df_estoque.to_excel(writer, sheet_name='Estoque', index=False)
            else:
                pd.DataFrame(columns=["Vacina", "Lote", "Fabricante", "Recebida", "Utilizada", "Saldo", "Validade"]).to_excel(writer, sheet_name='Estoque', index=False)
        
        arquivo_saida.seek(0)
        return arquivo_saida
    
    except Exception as e:
        print(f"ERRO ESTOQUE: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def exportar_faltosos_excel():
    """Exporta faltosos em Excel"""
    try:
        arquivo_saida = BytesIO()
        
        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            df_faltosos = buscar_pacientes_faltosos()
            df_faltosos.to_excel(writer, sheet_name='Pendentes', index=False)
        
        arquivo_saida.seek(0)
        return arquivo_saida
    
    except Exception as e:
        print(f"ERRO FALTOSOS: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def exportar_relatorio_excel(filtro_cpf="", filtro_vacina="", filtro_data_inicio="", filtro_data_fim=""):
    """Exporta relatório em Excel"""
    try:
        arquivo_saida = BytesIO()
        
        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            df_relatorio = gerar_relatorio_historico(filtro_cpf, filtro_vacina, filtro_data_inicio, filtro_data_fim)
            df_relatorio.to_excel(writer, sheet_name='Relatorio', index=False)
        
        arquivo_saida.seek(0)
        return arquivo_saida
    
    except Exception as e:
        print(f"ERRO RELATÓRIO: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def obter_dashboard():
    total_pac = db.get_total_pacientes()
    total_vac = db.get_total_vacinacoes()
    total_lot = len(db.lotes)
    
    return f"""
📊 DASHBOARD

👥 Pacientes: {total_pac}
💉 Vacinações: {total_vac}
💊 Lotes: {total_lot}

✅ Sistema Online"""

def listar_historico_vacinacoes():
    """Retorna histórico de vacinações em formato DataFrame"""
    if not db.vacinacoes:
        return pd.DataFrame(columns=[
            "Paciente", "CPF", "Vacina", "Dose", "Data", "Lote", 
            "Fabricante", "Profissional", "Local", "Via", "Campanha"
        ])
    
    dados = []
    for v in db.vacinacoes:
        dados.append([
            v.get('paciente_nome', ''),
            v.get('paciente_cpf', ''),
            v.get('vacina', ''),
            v.get('dose', ''),
            converter_data_para_br(v.get('data', '')),
            v.get('lote', ''),
            v.get('fabricante', ''),
            v.get('profissional', ''),
            v.get('local_injecao', ''),
            v.get('via_administracao', ''),
            v.get('tipo_campanha', '')
        ])
    
    return pd.DataFrame(dados, columns=[
        "Paciente", "CPF", "Vacina", "Dose", "Data", "Lote", 
        "Fabricante", "Profissional", "Local", "Via", "Campanha"
    ])

# FUNÇÕES DE PROTOCOLOS

def listar_protocolos_tabela():
    """Lista protocolos em formato DataFrame"""
    dados = []
    for p in gerenciador_protocolos.obter_protocolos():
        dados.append([
            p['id'],
            p['imunizante'],
            p['idade_minima_meses'],
            p['idade_maxima_meses'],
            p['esquema_doses'],
            p['intervalo_dias'],
            p['tipo'],
            p['faixa_etaria'],
            p.get('via_administracao', '')
        ])
    
    return pd.DataFrame(dados, columns=[
        "ID", "Imunizante", "Idade Mín.", "Idade Máx.", 
        "Esquema", "Intervalo (dias)", "Tipo", "Faixa Etária", "Via Admin"
    ]) if dados else pd.DataFrame(columns=[
        "ID", "Imunizante", "Idade Mín.", "Idade Máx.", 
        "Esquema", "Intervalo (dias)", "Tipo", "Faixa Etária", "Via Admin"
    ])

def atualizar_protocolo_campo(protocolo_id, campo, novo_valor):
    """Atualiza campo específico do protocolo"""
    try:
        if not protocolo_id or not campo:
            return "❌ Campos inválidos!"
        
        if campo in ['idade_minima_meses', 'idade_maxima_meses', 'intervalo_dias']:
            novo_valor = int(novo_valor)
        
        db_update = {campo: novo_valor}
        gerenciador_protocolos.atualizar_protocolo(int(protocolo_id), **db_update)
        
        return f"✅ Protocolo {protocolo_id} atualizado!"
    
    except Exception as e:
        return f"❌ Erro: {str(e)}"

def adicionar_protocolo(imunizante, idade_min, idade_max, esquema, intervalo, tipo, faixa, via_admin):
    """Adiciona novo protocolo"""
    try:
        if not imunizante or not tipo:
            return "❌ Preencha Imunizante e Tipo!"
        
        gerenciador_protocolos.adicionar_protocolo(
            imunizante, idade_min, idade_max, esquema, intervalo, tipo, faixa, via_admin
        )
        
        return f"✅ Protocolo '{imunizante}' adicionado!"
    
    except Exception as e:
        return f"❌ Erro: {str(e)}"

def deletar_protocolo(protocolo_id):
    """Deleta protocolo"""
    try:
        if not protocolo_id:
            return "❌ Digite um ID válido!"
        
        gerenciador_protocolos.deletar_protocolo(int(protocolo_id))
        return f"✅ Protocolo {protocolo_id} deletado!"
    
    except Exception as e:
        return f"❌ Erro: {str(e)}"

# FUNÇÕES DE RELATÓRIOS

def gerar_relatorio_historico(filtro_cpf="", filtro_vacina="", filtro_data_inicio="", filtro_data_fim=""):
    """Gera relatório com histórico de vacinações"""
    
    vacinacoes_filtradas = db.vacinacoes.copy()
    
    if filtro_cpf:
        vacinacoes_filtradas = [v for v in vacinacoes_filtradas if filtro_cpf.lower() in str(v.get('paciente_cpf', '')).lower()]
    
    if filtro_vacina:
        vacinacoes_filtradas = [v for v in vacinacoes_filtradas if filtro_vacina.lower() in v.get('vacina', '').lower()]
    
    if filtro_data_inicio:
        try:
            data_inicio = datetime.strptime(filtro_data_inicio, '%Y-%m-%d').date()
            vacinacoes_filtradas = [v for v in vacinacoes_filtradas if datetime.strptime(v.get('data', ''), '%Y-%m-%d').date() >= data_inicio]
        except:
            pass
    
    if filtro_data_fim:
        try:
            data_fim = datetime.strptime(filtro_data_fim, '%Y-%m-%d').date()
            vacinacoes_filtradas = [v for v in vacinacoes_filtradas if datetime.strptime(v.get('data', ''), '%Y-%m-%d').date() <= data_fim]
        except:
            pass
    
    dados = []
    for v in vacinacoes_filtradas:
        dados.append([
            v.get('paciente_nome', ''),
            v.get('paciente_cpf', ''),
            v.get('vacina', ''),
            v.get('lote', ''),
            v.get('fabricante', ''),
            v.get('dose', ''),
            converter_data_para_br(v.get('data', '')),
            v.get('profissional', ''),
            v.get('local_injecao', '')
        ])
    
    return pd.DataFrame(dados, columns=[
        "Nome Paciente", "CPF", "Vacina", "Lote", "Fabricante",
        "Dose", "Data Aplicacao", "Profissional", "Local Injecao"
    ]) if dados else pd.DataFrame(columns=[
        "Nome Paciente", "CPF", "Vacina", "Lote", "Fabricante",
        "Dose", "Data Aplicacao", "Profissional", "Local Injecao"
    ])

# ============================================================================
# SEÇÃO 6: INTERFACE GRADIO
# ============================================================================

print("\n🚀 INICIANDO INTERFACE FINAL\n")

with gr.Blocks(title="Sistema de Gestão de Imunização") as interface:
    gr.Markdown("# 🏥 SISTEMA DE GESTÃO DE IMUNIZAÇÃO\n### ✅ Versão Final")
    
    with gr.Tabs() as main_tabs:
        
        # TAB 0: DASHBOARD
        with gr.Tab("📊 Dashboard"):
            dash = gr.Textbox(lines=15, interactive=False)
            btn_dash = gr.Button("🔄 Atualizar")
            btn_dash.click(obter_dashboard, outputs=[dash])
            interface.load(obter_dashboard, outputs=[dash])
        
        # TAB 1: IMPORTAR + ADICIONAR
        with gr.Tab("📥 Importar Dados / Adicionar Paciente"):
            gr.Markdown("### 📥 Importe arquivo Excel com dados de pacientes\n⭐ O sistema detecta automaticamente as colunas!")
            
            arq = gr.File(label="Excel", file_types=[".xlsx", ".xls"])
            btn_imp = gr.Button("🚀 Importar", variant="primary")
            msg_imp = gr.Textbox(interactive=False, lines=4)
            tab_imp = gr.Dataframe(interactive=False, headers=["Nome", "CPF", "Data Nascimento", "Idade", "Aldeia"])
            
            btn_imp.click(importar_excel, inputs=[arq], outputs=[msg_imp, tab_imp])
            
            gr.Markdown("---")
            gr.Markdown("### ➕ Adicionar Paciente Manualmente")
            
            with gr.Row():
                add_nome = gr.Textbox(label="Nome *")
                add_cpf = gr.Textbox(label="CPF")
            
            with gr.Row():
                add_data_nasc = gr.Textbox(label="Data de Nascimento (DD-MM-AAAA)")
                add_idade = gr.Textbox(label="Idade (anos)")
                add_aldeia = gr.Textbox(label="Aldeia")
            
            btn_adicionar_pac = gr.Button("➕ Adicionar Paciente", variant="primary")
            msg_adicionar_pac = gr.Textbox(interactive=False)
            
            btn_adicionar_pac.click(
                adicionar_paciente_manual,
                inputs=[add_nome, add_cpf, add_data_nasc, add_idade, add_aldeia],
                outputs=[msg_adicionar_pac]
            )
        
        # TAB 2: CAMPANHAS
        with gr.Tab("⚙️ Campanhas"):
            gr.Markdown("### ➕ Criar Campanha")
            
            with gr.Row():
                camp_nome = gr.Textbox(label="Nome da Campanha *")
                camp_vacina = gr.Textbox(label="Vacina *")
            
            with gr.Row():
                camp_lote = gr.Textbox(label="Lote *")
                camp_fab = gr.Textbox(label="Fabricante")
                camp_valid = gr.Textbox(label="Validade (DD-MM-AAAA)")
            
            with gr.Row():
                camp_prof = gr.Textbox(label="Profissional *")
                camp_dose = gr.Textbox(label="Dose", value="1")
            
            btn_criar_camp = gr.Button("➕ Criar Campanha (Ativa)", variant="primary")
            msg_camp = gr.Textbox(interactive=False)
            
            btn_criar_camp.click(
                criar_campanha_config, 
                inputs=[camp_nome, camp_vacina, camp_lote, camp_fab, camp_valid, camp_prof, camp_dose], 
                outputs=[msg_camp, gr.State(), camp_nome, camp_vacina, camp_lote, camp_fab, camp_valid, camp_prof]
            )
            
            gr.Markdown("---")
            gr.Markdown("### 🎯 Campanhas Pré-cadastradas (EDITÁVEIS)")
            
            tab_camp = gr.Dataframe(
                interactive=False,
                headers=["ID", "Nome", "Vacina", "Lote", "Fabricante", "Profissional", "Validade", "Dose", "Status"]
            )
            
            with gr.Row():
                camp_id_edit = gr.Number(label="ID para editar", precision=0)
                camp_campo_edit = gr.Dropdown(
                    choices=["nome", "vacina", "lote", "fabricante", "profissional", "validade"],
                    label="Campo"
                )
                camp_novo_valor = gr.Textbox(label="Novo Valor")
                btn_edit_camp = gr.Button("✏️ Editar", variant="secondary")
            
            msg_edit_camp = gr.Textbox(interactive=False)
            
            with gr.Row():
                id_del = gr.Number(label="ID para deletar", precision=0)
                btn_del = gr.Button("🗑️ Deletar", variant="stop")
                btn_atualizar_camp = gr.Button("🔄 Atualizar")
            
            msg_ativar = gr.Textbox(interactive=False)
            msg_delete = gr.Textbox(interactive=False)
            
            btn_edit_camp.click(atualizar_campanha_editar, inputs=[camp_id_edit, camp_campo_edit, camp_novo_valor], outputs=[msg_edit_camp])
            btn_atualizar_camp.click(listar_campanhas_config, outputs=[tab_camp])
            interface.load(listar_campanhas_config, outputs=[tab_camp])
            
            tab_camp.select(ativar_campanha_config, inputs=[tab_camp], outputs=[msg_ativar])
            btn_del.click(deletar_campanha, inputs=[id_del], outputs=[msg_delete, tab_camp])
        
        # TAB 3: BUSCAR
        with gr.Tab("🔍 Buscar Paciente") as tab_buscar:
            gr.Markdown("### ⭐ Busque e clique em um paciente - PULARÁ para ABA VACINAÇÃO!")
            
            nome_busca = gr.Textbox(label="Digite o nome", placeholder="Conforme digita aparece...")
            msg_busca = gr.Textbox(interactive=False, label="Status")
            tab_busca = gr.Dataframe(
                interactive=False,
                headers=["Nome", "CPF", "Data Nascimento", "Idade", "Aldeia"],
                label="⭐ CLIQUE AQUI"
            )
            
            input_nome_vac_ref = gr.Textbox(visible=False)
            input_cpf_vac_ref = gr.Textbox(visible=False)
            input_data_nasc_vac_ref = gr.Textbox(visible=False)
            input_idade_vac_ref = gr.Textbox(visible=False)
            input_vacina_nome_ref = gr.Textbox(visible=False)
            input_lote_vac_ref = gr.Textbox(visible=False)
            input_validade_vac_ref = gr.Textbox(visible=False)
            input_fab_vac_ref = gr.Textbox(visible=False)
            input_prof_vac_ref = gr.Textbox(visible=False)
            aba_destino = gr.Number(visible=False, value=0)
            
            nome_busca.change(buscar_paciente_realtime, inputs=[nome_busca], outputs=[tab_busca])
            
            tab_busca.select(
                preencher_vacinacao_ao_clicar, 
                inputs=[tab_busca], 
                outputs=[
                    msg_busca,
                    input_nome_vac_ref,
                    input_cpf_vac_ref,
                    input_data_nasc_vac_ref,
                    input_idade_vac_ref,
                    input_vacina_nome_ref,
                    input_lote_vac_ref,
                    input_validade_vac_ref,
                    input_fab_vac_ref,
                    input_prof_vac_ref,
                    aba_destino
                ]
            )
        
        # TAB 4: VACINAÇÃO
        with gr.Tab("💉 Vacinação") as tab_vacinacao:
            
            with gr.Row():
                tipo_campanha = gr.Radio(
                    choices=["Campanha", "Rotina"],
                    label="Tipo de Atendimento *",
                    value="Campanha"
                )
            
            gr.Markdown("### 👤 Dados do Paciente (EDITÁVEIS)")
            
            with gr.Row():
                input_nome_vac = gr.Textbox(label="Nome *", interactive=True, placeholder="Aparecerá ao clicar em Buscar")
                input_cpf_vac = gr.Textbox(label="CPF", interactive=True, placeholder="Aparecerá ao clicar em Buscar")
            
            with gr.Row():
                input_data_nasc_vac = gr.Textbox(label="Data de Nascimento", interactive=True, placeholder="DD-MM-AAAA")
                input_idade_vac = gr.Textbox(label="Idade (anos)", interactive=True, placeholder="Aparecerá ao clicar em Buscar")
                paciente_id_hidden = gr.Number(visible=False, value=0)
            
            with gr.Row():
                btn_editar_dados = gr.Button("✏️ Editar Dados do Paciente", variant="secondary")
            
            msg_editar_dados = gr.Textbox(interactive=False)
            
            gr.Markdown("---")
            gr.Markdown("### 💉 Dados da Vacinação")
            
            with gr.Row():
                input_vacina_nome = gr.Dropdown(
                    choices=obter_vacinas_disponiveis(),
                    label="Vacina *",
                    interactive=True,
                    allow_custom_value=True
                )
                input_lote_vac = gr.Textbox(label="Lote *", interactive=False, placeholder="Pré-preenchido")
                input_validade_vac = gr.Textbox(label="Validade (DD-MM-AAAA)", interactive=False, placeholder="Pré-preenchida")
                input_fab_vac = gr.Textbox(label="Fabricante", interactive=False, placeholder="Pré-preenchido")
            
            with gr.Row():
                input_dose_vac = gr.Dropdown(
                    choices=["D.U (Dose Única)", "D1 (1ª Dose)", "D2 (2ª Dose)", "D3 (3ª Dose)", "Reforço", "Reforço 2"],
                    label="Dose *",
                    interactive=True
                )
                input_data_vac = gr.Textbox(label="Data *", placeholder="DD-MM-AAAA", value=datetime.now().strftime("%d-%m-%Y"))
            
            with gr.Row():
                input_prof_vac = gr.Textbox(label="Profissional", interactive=False, placeholder="Pré-preenchido")
                input_local_vac = gr.Dropdown(
                    choices=[
                        "Braço Direito",
                        "Braço Esquerdo",
                        "Coxa Direita",
                        "Coxa Esquerda",
                        "Glúteo Direito",
                        "Glúteo Esquerdo",
                        "Perna Direita",
                        "Perna Esquerda",
                        "Outro"
                    ],
                    label="Local de Injeção *"
                )
                input_via_vac = gr.Textbox(label="Via de Administração *", interactive=False, placeholder="Preenchida automaticamente")
            
            with gr.Row():
                input_reacao_vac = gr.Textbox(label="Reação Adversa")
                input_obs_vac = gr.Textbox(label="Observações")
            
            gr.Markdown("---")
            gr.Markdown("### 📊 Status de Vacinação (conforme Protocolo)")
            
            with gr.Row():
                status_semaforo = gr.Textbox(label="Status", interactive=False, value="⚪ Selecione vacina e idade")
                proxima_dose_protocolo = gr.Textbox(label="Próxima Dose (Protocolo)", interactive=False, value="⚪ SEM DADOS")
            
            msg_validacao = gr.Textbox(label="Validação", interactive=False, value="⚪ SEM DADOS")
            
            input_prox_vac = gr.Textbox(label="Próxima Dose (Manual)", interactive=False, placeholder="Calculada automaticamente")
            
            btn_registrar_vac = gr.Button("✅ Registrar Vacinação", variant="primary")
            msg_registrar_vac = gr.Textbox(interactive=False)
            
            def atualizar_status_com_protocolo(vacina, dose, idade, pac_id):
                if not vacina or not idade:
                    return "⚪ SEM DADOS", "⚪ SEM DADOS", "⚪ SEM DADOS", ""
                
                vacina_clean = vacina.split(" (Lote:")[0] if " (Lote:" in vacina else vacina
                idade_int = int(float(idade)) if idade else 0
                
                status, proxima = obter_status_vacina_com_proxima(vacina_clean, str(idade_int), pac_id if pac_id else 0)
                
                validacao = validar_vacina_idade(vacina_clean, str(idade_int))
                
                prox_manual = calcular_proxima_dose_automatica(vacina_clean, dose, str(idade_int))
                
                return status, proxima, validacao, prox_manual
            
            input_vacina_nome.change(
                atualizar_dados_vacinacao,
                inputs=[input_vacina_nome],
                outputs=[input_lote_vac, input_validade_vac, input_fab_vac, input_via_vac]
            )
            
            input_vacina_nome.change(
                atualizar_status_com_protocolo,
                inputs=[input_vacina_nome, input_dose_vac, input_idade_vac, paciente_id_hidden],
                outputs=[status_semaforo, proxima_dose_protocolo, msg_validacao, input_prox_vac]
            )
            
            input_dose_vac.change(
                atualizar_status_com_protocolo,
                inputs=[input_vacina_nome, input_dose_vac, input_idade_vac, paciente_id_hidden],
                outputs=[status_semaforo, proxima_dose_protocolo, msg_validacao, input_prox_vac]
            )
            
            def atualizar_dropdown_vacinas():
                return gr.update(choices=obter_vacinas_disponiveis())
            
            btn_atualizar_vacinas = gr.Button("🔄 Atualizar Lista de Vacinas", variant="secondary")
            
            btn_atualizar_vacinas.click(
                atualizar_dropdown_vacinas,
                outputs=[input_vacina_nome]
            )
            
            tab_vacinacao.select(
                atualizar_dropdown_vacinas,
                outputs=[input_vacina_nome]
            )
            
            def sincronizar_dados_busca(nome, cpf, data_nasc, idade, vacina, lote, validade, fab, prof, pac_id):
                # Obtém ID do paciente
                pac_id_real = 0
                for p in db.pacientes:
                    if p['nome'].lower() == nome.lower():
                        pac_id_real = p['id']
                        break
                
                return (
                    gr.update(value=nome),
                    gr.update(value=cpf),
                    gr.update(value=data_nasc if data_nasc != '-' else ''),
                    gr.update(value=idade),
                    gr.update(value=vacina) if vacina else gr.update(value=""),
                    gr.update(value=lote),
                    gr.update(value=validade),
                    gr.update(value=fab),
                    gr.update(value=prof),
                    pac_id_real
                )
            
            input_nome_vac_ref.change(
                sincronizar_dados_busca,
                inputs=[input_nome_vac_ref, input_cpf_vac_ref, input_data_nasc_vac_ref, input_idade_vac_ref, input_vacina_nome_ref, input_lote_vac_ref, input_validade_vac_ref, input_fab_vac_ref, input_prof_vac_ref, aba_destino],
                outputs=[input_nome_vac, input_cpf_vac, input_data_nasc_vac, input_idade_vac, input_vacina_nome, input_lote_vac, input_validade_vac, input_fab_vac, input_prof_vac, paciente_id_hidden]
            )
            
            def registrar_e_atualizar(paciente_nome, vacina, dose, data_br, lote, fabricante, profissional, 
                                      local_injecao, via_admin, reacao, obs, prox_dose, tipo_campanha):
                resultado = registrar_vacinacao(paciente_nome, vacina, dose, data_br, lote, fabricante, profissional, 
                                               local_injecao, via_admin, reacao, obs, prox_dose, tipo_campanha)
                
                historico = listar_historico_vacinacoes()
                
                return resultado, historico
            
            # Callback para editar dados do paciente
            btn_editar_dados.click(
                editar_dados_paciente_atual,
                inputs=[input_nome_vac, input_nome_vac, input_cpf_vac, input_idade_vac],
                outputs=[msg_editar_dados]
            )
            
            btn_registrar_vac.click(
                registrar_e_atualizar,
                inputs=[input_nome_vac, input_vacina_nome, input_dose_vac, input_data_vac, 
                       input_lote_vac, input_fab_vac, input_prof_vac, input_local_vac, 
                       input_via_vac, input_reacao_vac, input_obs_vac, input_prox_vac, tipo_campanha],
                outputs=[msg_registrar_vac, gr.State()]
            )
            
            gr.Markdown("---")
            gr.Markdown("### 📋 Histórico de Vacinados (EXPORTÁVEL)")
            
            tab_historico = gr.Dataframe(
                interactive=False,
                headers=[
                    "Paciente", "CPF", "Vacina", "Dose", "Data", "Lote", 
                    "Fabricante", "Profissional", "Local", "Via", "Campanha"
                ]
            )
            
            with gr.Row():
                btn_exportar_historico = gr.Button("📥 Exportar Histórico Excel", variant="primary")
            
            arquivo_historico = gr.File(label="Download Histórico", visible=True)
            
            tab_vacinacao.select(
                listar_historico_vacinacoes,
                outputs=[tab_historico]
            )
            
            btn_registrar_vac.click(
                listar_historico_vacinacoes,
                outputs=[tab_historico]
            )
            
            btn_exportar_historico.click(exportar_historico_excel, outputs=[arquivo_historico])
        
        # TAB 5: CONFIGURAR PROTOCOLOS
        with gr.Tab("⚙️ Configurar Protocolos"):
            gr.Markdown("### 📋 Tabela de Vacinas - PNV Indígena 2024 (EDITÁVEL)")
            
            gr.Markdown("#### ➕ Adicionar Novo Protocolo")
            
            with gr.Row():
                prot_imunizante = gr.Textbox(label="Imunizante *")
                prot_tipo = gr.Dropdown(
                    choices=["Dose Única", "Esquema Sequencial", "Anual", "Reforço"],
                    label="Tipo *"
                )
            
            with gr.Row():
                prot_idade_min = gr.Number(label="Idade Mínima (meses)", value=0, precision=0)
                prot_idade_max = gr.Number(label="Idade Máxima (meses)", value=120, precision=0)
            
            with gr.Row():
                prot_esquema = gr.Textbox(label="Esquema de Doses")
                prot_intervalo = gr.Number(label="Intervalo entre Doses (dias)", value=30, precision=0)
            
            with gr.Row():
                prot_faixa = gr.Textbox(label="Faixa Etária Principal")
                prot_via = gr.Textbox(label="Via de Administração")
            
            btn_adicionar_prot = gr.Button("➕ Adicionar Protocolo", variant="primary")
            msg_protocolo = gr.Textbox(interactive=False)
            
            btn_adicionar_prot.click(
                adicionar_protocolo,
                inputs=[prot_imunizante, prot_idade_min, prot_idade_max, prot_esquema, prot_intervalo, prot_tipo, prot_faixa, prot_via],
                outputs=[msg_protocolo]
            )
            
            gr.Markdown("---")
            gr.Markdown("#### 📊 Tabela de Protocolos de Vacinação (EDITÁVEIS)")
            
            tab_protocolos = gr.Dataframe(
                interactive=False,
                headers=["ID", "Imunizante", "Idade Mín.", "Idade Máx.", "Esquema", "Intervalo (dias)", "Tipo", "Faixa Etária", "Via Admin"]
            )
            
            with gr.Row():
                prot_id_edit = gr.Number(label="ID para editar", precision=0)
                prot_campo_edit = gr.Textbox(label="Campo a editar (ex: idade_minima_meses, intervalo_dias)")
                prot_novo_valor = gr.Textbox(label="Novo Valor")
                btn_edit_prot = gr.Button("✏️ Editar", variant="secondary")
            
            msg_edit_prot = gr.Textbox(interactive=False)
            
            with gr.Row():
                prot_id_del = gr.Number(label="ID para deletar", precision=0)
                btn_del_prot = gr.Button("🗑️ Deletar Protocolo", variant="stop")
                btn_atualizar_prot = gr.Button("🔄 Atualizar Lista")
            
            msg_del_prot = gr.Textbox(interactive=False)
            
            btn_edit_prot.click(atualizar_protocolo_campo, inputs=[prot_id_edit, prot_campo_edit, prot_novo_valor], outputs=[msg_edit_prot])
            btn_atualizar_prot.click(listar_protocolos_tabela, outputs=[tab_protocolos])
            btn_del_prot.click(deletar_protocolo, inputs=[prot_id_del], outputs=[msg_del_prot])
            interface.load(listar_protocolos_tabela, outputs=[tab_protocolos])
        
        # TAB 6: ESTOQUE
        with gr.Tab("📦 Estoque"):
            gr.Markdown("### 💉 Lotes de Vacinas")
            
            with gr.Row():
                est_vac = gr.Textbox(label="Vacina *")
                est_lote = gr.Textbox(label="Lote *")
                est_fab = gr.Textbox(label="Fabricante")
            
            with gr.Row():
                est_qtd = gr.Number(label="Quantidade *", precision=0)
                est_valid = gr.Textbox(label="Validade (DD-MM-AAAA)")
            
            btn_est = gr.Button("✅ Adicionar Lote", variant="primary")
            msg_est = gr.Textbox(interactive=False)
            
            btn_est.click(adicionar_lote_estoque, inputs=[est_vac, est_lote, est_fab, est_qtd, est_valid], outputs=[msg_est])
            
            gr.Markdown("---")
            gr.Markdown("### Insumos (Seringas, Agulhas, etc)")
            
            with gr.Row():
                ins_tipo = gr.Textbox(label="Tipo *")
                ins_desc = gr.Textbox(label="Descrição *")
            
            with gr.Row():
                ins_qtd_solic = gr.Number(label="Qtd. Solicitada", precision=0)
                ins_qtd_receb = gr.Number(label="Qtd. Recebida", precision=0)
            
            with gr.Row():
                ins_ml = gr.Textbox(label="Seringa (ml)")
                ins_cal = gr.Textbox(label="Calibre")
                ins_lote = gr.Textbox(label="Lote")
                ins_valid = gr.Textbox(label="Validade")
            
            btn_ins = gr.Button("✅ Adicionar Insumo", variant="primary")
            msg_ins = gr.Textbox(interactive=False)
            
            btn_ins.click(adicionar_insumo_estoque, inputs=[ins_tipo, ins_desc, ins_qtd_solic, ins_qtd_receb, ins_ml, ins_cal, ins_lote, ins_valid], outputs=[msg_ins])
            
            gr.Markdown("---")
            gr.Markdown("### Tabelas de Estoque (EXPORTÁVEL)")
            
            with gr.Row():
                btn_atualizar_est = gr.Button("🔄 Atualizar Estoque")
                btn_exportar_est = gr.Button("📥 Exportar Estoque Excel", variant="primary")
            
            tab_est = gr.Dataframe(interactive=False, headers=["Vacina", "Lote", "Fabricante", "Recebida", "Utilizada", "Saldo", "Validade"])
            tab_ins = gr.Dataframe(interactive=False, headers=["Tipo", "Descrição", "Solic.", "Receb.", "Diferença", "Seringa", "Calibre", "Lote", "Validade"])
            
            arquivo_estoque = gr.File(label="Download Estoque", visible=True)
            
            def atualizar_est():
                return listar_estoque(), listar_insumos_estoque()
            
            btn_atualizar_est.click(atualizar_est, outputs=[tab_est, tab_ins])
            btn_exportar_est.click(exportar_estoque_excel, outputs=[arquivo_estoque])
            interface.load(atualizar_est, outputs=[tab_est, tab_ins])
        
        # TAB 7: FALTOSOS
        with gr.Tab("📋 Pendências/Faltosos"):
            gr.Markdown("### ⚠️ Pacientes sem vacinação registrada (EXPORTÁVEL)")
            
            tab_faltosos = gr.Dataframe(interactive=False, headers=["Nome", "CPF", "Data Nascimento", "Idade", "Aldeia"])
            
            with gr.Row():
                btn_atualizar_faltosos = gr.Button("🔄 Atualizar Lista")
                btn_exportar_faltosos = gr.Button("📥 Exportar Pendentes Excel", variant="primary")
            
            arquivo_faltosos = gr.File(label="Download Pendentes", visible=True)
            
            btn_atualizar_faltosos.click(buscar_pacientes_faltosos, outputs=[tab_faltosos])
            btn_exportar_faltosos.click(exportar_faltosos_excel, outputs=[arquivo_faltosos])
            interface.load(buscar_pacientes_faltosos, outputs=[tab_faltosos])
        
        # TAB 8: RELATÓRIOS
        with gr.Tab("📊 Relatórios/Histórico"):
            gr.Markdown("### 📋 Histórico Consolidado de Vacinações (EXPORTÁVEL)")
            
            gr.Markdown("#### 🔍 Filtros")
            
            with gr.Row():
                rel_cpf = gr.Textbox(label="Filtrar por CPF")
                rel_vacina = gr.Textbox(label="Filtrar por Vacina")
            
            with gr.Row():
                rel_data_ini = gr.Textbox(label="Data Início (YYYY-MM-DD)")
                rel_data_fim = gr.Textbox(label="Data Fim (YYYY-MM-DD)")
            
            btn_gerar_rel = gr.Button("🔍 Gerar Relatório", variant="primary")
            
            tab_relatorio = gr.Dataframe(
                interactive=False,
                headers=[
                    "Nome Paciente", "CPF", "Vacina", "Lote", "Fabricante",
                    "Dose", "Data Aplicacao", "Profissional", "Local Injecao"
                ]
            )
            
            with gr.Row():
                btn_exportar_rel = gr.Button("📥 Exportar Relatório Excel", variant="primary")
            
            arquivo_relatorio = gr.File(label="Download Relatório", visible=True)
            
            btn_gerar_rel.click(
                gerar_relatorio_historico,
                inputs=[rel_cpf, rel_vacina, rel_data_ini, rel_data_fim],
                outputs=[tab_relatorio]
            )
            
            btn_exportar_rel.click(
                exportar_relatorio_excel,
                inputs=[rel_cpf, rel_vacina, rel_data_ini, rel_data_fim],
                outputs=[arquivo_relatorio]
            )
        
        # TAB 9: CONSOLIDADO
        with gr.Tab("📊 Consolidado"):
            gr.Markdown("### 📊 CONSOLIDADO TOTAL - UM ARQUIVO EXCEL COM TODAS AS ABAS")
            
            with gr.Row():
                btn_exportar_consolidado_super = gr.Button("📥 EXPORTAR CONSOLIDADO COMPLETO", variant="primary")
            
            msg_consolidado = gr.Textbox(interactive=False, label="Status")
            arquivo_consolidado_super = gr.File(label="📥 Download: CONSOLIDADO_COMPLETO.xlsx", visible=True)
            
            gr.Markdown("---")
            gr.Markdown("### 📊 Relatório Consolidado (Dados Resumidos)")
            
            with gr.Row():
                total_vac = gr.Number(label="Total Vacinados", interactive=False)
                saldo_doses = gr.Number(label="Saldo Doses", interactive=False)
                saldo_ins = gr.Number(label="Saldo Insumos", interactive=False)
            
            tab_consolidado = gr.Dataframe(interactive=False, headers=["Insumo", "Solicitado", "Recebido", "Diferença", "Aplicado", "Saldo Real"])
            
            with gr.Row():
                btn_atualizar_consolidado = gr.Button("🔄 Atualizar Consolidado", variant="primary")
            
            def gerar_consolidado_display():
                total_vac, saldo_doses, saldo_ins, df_consolidado = gerar_consolidado()
                return total_vac, saldo_doses, saldo_ins, df_consolidado, "✅ Consolidado gerado com sucesso!"
            
            def exportar_consolidado_com_msg():
                arquivo = exportar_consolidado_completo_excel()
                if arquivo is None:
                    return None, "❌ Erro ao gerar consolidado!"
                return arquivo, "✅ Excel gerado com sucesso! (5 abas: Historico Vacinados, Estoque, Pendentes, Relatorio, Consolidado)"
            
            btn_atualizar_consolidado.click(gerar_consolidado_display, outputs=[total_vac, saldo_doses, saldo_ins, tab_consolidado, msg_consolidado])
            btn_exportar_consolidado_super.click(exportar_consolidado_com_msg, outputs=[arquivo_consolidado_super, msg_consolidado])
            interface.load(gerar_consolidado_display, outputs=[total_vac, saldo_doses, saldo_ins, tab_consolidado, msg_consolidado])

print("\n✅ INTERFACE PRONTA - EXECUTANDO!\n")
print("=" * 80)
print("✅ SISTEMA COMPLETO!")
print("=" * 80 + "\n")

# ============================================================================
# SEÇÃO 7: LAUNCH
# ============================================================================

# ============================================================================
# INICIALIZAÇÃO
# ============================================================================

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 7860))
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
