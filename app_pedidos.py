import streamlit as st
import pandas as pd
import urllib.parse
import os
import unicodedata
import base64

# Configuração da página
st.set_page_config(page_title="Tigre Alimentos - Pedidos", page_icon="🐅", layout="wide")

# Estilos baseados no PDF (Identidade Visual Tigre Alimentos)
st.markdown("""
<style>
    /* Fundo da aplicação */
    .stApp {
        background-color: #f7fbf4; /* Fundo levemente esverdeado/off-white */
    }
    
    /* Título principal e subtítulo */
    h1, h2, h3 {
        color: #1a4d2e !important; /* Verde escuro (cor do logotipo e textos escuros) */
        font-family: 'Arial Black', sans-serif;
    }
    
    /* Linha divisória verde */
    hr {
        border-top: 2px solid #7ac142 !important; /* Verde claro vibrante do PDF */
    }
    
    /* Corrigir a cor do texto das categorias (Expanders) e aumentar a fonte */
    .streamlit-expanderHeader, summary, summary p, summary span {
        color: #1a4d2e !important;
        font-weight: 900 !important;
        font-size: 1.4rem !important;
    }
    
    /* Nome do produto (Preto, bold) */
    .produto-titulo {
        font-size: 1.6rem;
        font-weight: 800;
        color: #000000;
        text-transform: uppercase;
        margin-bottom: 0px;
        line-height: 1.2;
    }
    
    /* Preço do produto (Verde claro vibrante do PDF) */
    .produto-preco {
        font-size: 1.4rem;
        font-weight: 600;
        color: #7ac142;
        margin-top: 5px;
    }
    
    /* Estilizar o fundo de cada linha (como se fossem as faixas do PDF) */
    .produto-container {
        background-color: white;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0px 2px 5px rgba(0,0,0,0.05);
        margin-bottom: 10px;
        border-left: 5px solid #7ac142;
    }
    
    /* Uniformizar o tamanho e proporção de todas as fotos dos produtos */
    .produto-container [data-testid="stImage"] img, .produto-container .stImage img, .stImage img {
        height: 110px !important;
        width: 100% !important;
        object-fit: contain !important;
        margin: 0 auto;
        display: block;
    }
    
    /* Responsividade para Celulares */
    @media (max-width: 800px) {
        /* Imagens uniformes e adaptadas no celular */
        .produto-container [data-testid="stImage"] img, .produto-container .stImage img, .stImage img {
            height: 80px !important;
            max-width: 80px !important;
            object-fit: contain !important;
            margin: 0 auto;
            display: block;
        }
        
        /* Carrinho fixo no rodapé, max 20% da tela, com fontes reduzidas */
        .carrinho-flutuante {
            top: auto !important;
            bottom: 0px !important;
            right: 0px !important;
            left: 0px !important;
            width: 100% !important;
            border-radius: 15px 15px 0 0 !important;
            border: none !important;
            border-top: 2px solid #7ac142 !important;
            box-shadow: 0px -4px 10px rgba(0,0,0,0.15) !important;
            padding: 8px !important;
            max-height: 25vh !important;
            overflow-y: auto !important; /* Rolar internamente */
        }
        
        /* Reduzir TODAS as fontes dentro do carrinho no mobile */
        .carrinho-flutuante h2 { font-size: 1.1rem !important; margin: 0 0 2px 0 !important; }
        .carrinho-flutuante h3 { font-size: 1.0rem !important; margin: 0 !important; }
        .carrinho-flutuante p { font-size: 0.8rem !important; margin: 0 !important; }
        .carrinho-flutuante div { font-size: 0.85rem !important; }
        .carrinho-flutuante button { font-size: 0.9rem !important; padding: 8px !important; margin-top: 2px !important; }
        
        /* Tirar a caixa com scroll duplo dentro do carrinho */
        .carrinho-flutuante > div {
            max-height: none !important;
            margin: 5px 0 !important;
            padding: 8px !important;
        }
        
        /* Espaço no final da página para o carrinho não sobrepor os itens */
        .stApp {
            padding-bottom: 25vh !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# Título / Cabeçalho
if os.path.exists("logo_header.png"):
    with open("logo_header.png", "rb") as f:
        data_img = base64.b64encode(f.read()).decode("utf-8")
    st.markdown(f'''
        <div style="background-color: #3a7227; width: 100%; padding: 20px 0; text-align: center; border-radius: 12px; margin-bottom: 15px; box-shadow: 0px 4px 12px rgba(0,0,0,0.15);">
            <img src="data:image/png;base64,{data_img}" style="max-height: 200px; max-width: 95%; object-fit: contain;">
        </div>
    ''', unsafe_allow_html=True)
else:
    st.markdown("<h1 style='text-align: center; color: #1a4d2e;'>🐅 TIGRE ALIMENTOS</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center; color: #7ac142;'>Tabela de Produtos (Gerador de Pedidos)</h3>", unsafe_allow_html=True)
st.markdown("---")
@st.cache_data
def carregar_dados(mtime):
    # Carregando a planilha excel
    df = pd.read_excel('Produtos_tigre_pedido_exemplo.xlsx')
    
    # Limpando os dados
    df = df.dropna(subset=['PRODUTOS', 'PREÇO_UN', 'PREÇO_FD']) # Remover linhas sem produto ou preço
    
    # Filtrar itens indesejados (como 'Total Frutas', etc, caso ainda estejam na base)
    df = df[~df['PRODUTOS'].str.contains('Total|peso total|Frete|Seguro|ag. frete|Amarra', case=False, na=False)]
    
    # Transformar as colunas de preço em numérico
    df['PREÇO_UN'] = pd.to_numeric(df['PREÇO_UN'], errors='coerce')
    df['PREÇO_FD'] = pd.to_numeric(df['PREÇO_FD'], errors='coerce')
    df = df.dropna(subset=['PREÇO_UN', 'PREÇO_FD'])
    
    # Criar categorias baseadas na primeira palavra do produto
    def extrair_categoria(nome):
        p = nome.split()[0].upper()
        
        # Categorias principais
        if p == 'AMENDOIM': return 'AMENDOINS'
        if p in ['ACUCAR', 'AÇUCAR', 'AÇÚCAR', 'SAL', 'BICARBONATO']: return 'SAIS E AÇÚCARES'
        if p in ['FEIJAO', 'FEIJÃO']: return 'FEIJÕES'
        if p in ['FARINHA', 'FAROFA', 'FUBA', 'FÚBA', 'TAPIOCA', 'POLVILHO']: return 'FARINHAS'
        if p in ['DOCE', 'MELADO', 'RAPADURA', 'SAGU', 'PAÇOCA', 'PACOCA', 'PAÇOQUINHA']: return 'DOCES'
        
        # Agrupar grãos e sementes
        graos = ['ERVILHA', 'GRAO', 'GRÃO', 'FÁBA', 'FAVA', 'GIRASSOL', 'LENTILHA', 'MILHO', 'PIPOCA', 'SOJA', 'CANJICA', 'CANJIQUINHA', 'TRIGO']
        if p in graos:
            return 'OUTROS GRÃOS'
            
        return p
        
    df['CATEGORIA'] = df['PRODUTOS'].apply(extrair_categoria)
    
    return df

try:
    caminho_planilha = 'Produtos_tigre_pedido_exemplo.xlsx'
    mtime = os.path.getmtime(caminho_planilha) if os.path.exists(caminho_planilha) else 0
    df_produtos = carregar_dados(mtime)
    
    def normalizar_nome(texto):
        texto_norm = unicodedata.normalize('NFKD', str(texto)).encode('ASCII', 'ignore').decode('ASCII')
        return "".join([c for c in texto_norm if c.isalnum()]).lower()

    mapa_fotos = {}
    if os.path.exists("fotos_de_produtos"):
        for arq in os.listdir("fotos_de_produtos"):
            if arq.lower().endswith((".png", ".jpg", ".jpeg")):
                nome_base = os.path.splitext(arq)[0]
                mapa_fotos[normalizar_nome(nome_base)] = os.path.join("fotos_de_produtos", arq)
    
    st.write("Selecione a quantidade desejada de fardos/caixas de cada produto:")
    
    # Inicializar variáveis de controle do carrinho
    if 'carrinho' not in st.session_state:
        st.session_state.carrinho = {}

    total_pedido = 0.0
    
    # Criar lista de itens do pedido formatada
    itens_pedido_texto = ""

    # Criar colunas para o layout principal (produtos na esquerda, carrinho na direita)
    col_produtos, col_carrinho = st.columns([7, 3], gap="large")

    with col_produtos:
        # Ordem fixa solicitada pelo usuário
        ordem_desejada = [
            'AMENDOINS', 
            'FEIJÕES', 
            'OUTROS GRÃOS', 
            'FARINHAS', 
            'DOCES', 
            'SAIS E AÇÚCARES'
        ]
        
        categorias_existentes = df_produtos['CATEGORIA'].unique().tolist()
        categorias = [cat for cat in ordem_desejada if cat in categorias_existentes]
        
        # Adiciona no fim qualquer categoria não mapeada na ordem
        for cat in categorias_existentes:
            if cat not in ordem_desejada:
                categorias.append(cat)
        
        # Mapeamento de emojis por categoria
        emojis = {
            'AMENDOINS': '🥜',
            'FEIJÕES': '🫘',
            'SAIS E AÇÚCARES': '🧂',
            'FARINHAS': '🌾',
            'OUTROS GRÃOS': '🌱',
            'DOCES': '🍯',
            'DIVERSOS': '📦'
        }
        
        for cat in categorias:
            emoji = emojis.get(cat, '📦') # Usa a caixa genérica caso não ache a categoria
            # Expanders colapsáveis para cada categoria
            with st.expander(f"{emoji} {cat}", expanded=False):
                df_cat = df_produtos[df_produtos['CATEGORIA'] == cat].sort_values(
                    by='PRODUTOS',
                    key=lambda col: col.astype(str).apply(lambda s: unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('ASCII').upper())
                )
                
                for index, row in df_cat.iterrows():
                    produto = row['PRODUTOS'].strip()
                    preco_fardo = row['PREÇO_FD']
                    preco_unidade = row['PREÇO_UN']
                    
                    # Buscar imagem de forma inteligente (insensível a acentos, maiúsculas/minúsculas e símbolos)
                    chave_norm = normalizar_nome(produto)
                    if chave_norm in mapa_fotos:
                        url_imagem = mapa_fotos[chave_norm]
                    else:
                        url_imagem = row.get('IMAGEM', "https://dummyimage.com/100x100/7ac142/ffffff.png&text=Foto")
                        if pd.isna(url_imagem):
                            url_imagem = "https://dummyimage.com/100x100/7ac142/ffffff.png&text=Foto"
            
                    # Container visual para simular o card do PDF
                    st.markdown('<div class="produto-container">', unsafe_allow_html=True)
                    col_img, col_info, col_btn = st.columns([1.2, 3, 1.2], vertical_alignment="center")
                    
                    with col_img:
                        st.image(url_imagem, use_container_width=True)
            
                    with col_info:
                        st.markdown(f'<p class="produto-titulo">{produto}</p>', unsafe_allow_html=True)
                        st.markdown(f'<p class="produto-preco">R$ {preco_fardo:.2f} <span style="font-size: 1.2rem; color: #666; font-weight: normal;">(Ref: R$ {preco_unidade:.2f}/un)</span></p>'.replace(".", ","), unsafe_allow_html=True)
                    
                    with col_btn:
                        quantidade = st.number_input(f"Qtd ({produto})", min_value=0, value=0, step=1, key=f"qtd_{index}", label_visibility="collapsed")
                    st.markdown('</div>', unsafe_allow_html=True)
                        
                    if quantidade > 0:
                        subtotal = quantidade * preco_fardo
                        total_pedido += subtotal
                        st.session_state.carrinho[produto] = {'qtd': quantidade, 'subtotal': subtotal}
                        itens_pedido_texto += f"✅ {quantidade} fardos - {produto} (R$ {subtotal:.2f})\n"

    with col_carrinho:
        if total_pedido > 0:
            itens_html = itens_pedido_texto.replace('\n', '<br>')
            
            texto_wpp = f"🐅 *NOVO PEDIDO - TIGRE ALIMENTOS* 🐅\n\n*Itens:*\n{itens_pedido_texto}\n💰 *TOTAL DO PEDIDO:* R$ {total_pedido:.2f}\n\nAguardando confirmação e envio!"
            texto_codificado = urllib.parse.quote(texto_wpp)
            link_whatsapp = f"https://api.whatsapp.com/send?text={texto_codificado}"
            
            valor_formatado = f"R$ {total_pedido:.2f}".replace(".", ",")
            
            # Construir o HTML completo da janela flutuante
            cart_html = f"""<div class="carrinho-flutuante" style="position: fixed; top: 80px; right: 5%; width: 25%; background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0px 4px 10px rgba(0,0,0,0.1); border: 2px solid #7ac142; z-index: 9999;">
<h2 style="color: #1a4d2e; font-family: 'Arial Black', sans-serif; margin: 0 0 5px 0; font-size: 1.6rem;">🛒 Resumo do Pedido</h2>
<p style="font-size: 0.85rem; color: #666; text-align: center; font-style: italic; margin: 0 0 5px 0;">(Role para baixo para ver os detalhes 👇)</p>
<h3 style="color: #1a4d2e; font-family: 'Arial Black', sans-serif; font-size: 1.3rem; margin: 0;">Valor Total: <b>{valor_formatado}</b></h3>
<div style="background-color: #e8f5e3; padding: 10px; border-radius: 8px; border: 1px solid #7ac142; margin: 10px 0; max-height: 35vh; overflow-y: auto;">
<p style="color: #000000; font-weight: bold; margin: 0 0 5px 0; font-size: 1rem;">Itens selecionados:</p>
<div style="color: #000000; font-size: 0.95rem; line-height: 1.4;">{itens_html}</div>
</div>
<a href="{link_whatsapp}" target="_blank" style="text-decoration: none;">
<button style="background-color:#25D366;color:white;padding:12px 15px;border:none;border-radius:5px;font-size:15px;font-weight:bold;cursor:pointer;width:100%;margin-top:5px;box-shadow: 0px 4px 6px rgba(0,0,0,0.1);">Enviar Pedido via WhatsApp 📲</button>
</a>
</div>"""
            st.markdown(cart_html, unsafe_allow_html=True)
        else:
            # Janela flutuante vazia
            empty_html = """<div class="carrinho-flutuante" style="position: fixed; top: 80px; right: 5%; width: 25%; background-color: #ffffff; padding: 20px; border-radius: 10px; box-shadow: 0px 4px 10px rgba(0,0,0,0.1); border: 2px solid #7ac142; z-index: 9999;">
<h2 style="color: #1a4d2e; font-family: 'Arial Black', sans-serif; margin-top: 0; font-size: 1.8rem;">🛒 Resumo do Pedido</h2>
<p style="color: #666; font-size: 1.1rem;">Seu carrinho está vazio.</p>
</div>"""
            st.markdown(empty_html, unsafe_allow_html=True)

except Exception as e:
    st.error(f"Erro ao ler a planilha: {e}")
    st.write("Certifique-se de que o arquivo 'Produtos_tigre_pedido_exemplo.xlsx' está na mesma pasta.")
