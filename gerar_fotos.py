import os
import pdfplumber
import pandas as pd

def extract_images():
    print("Iniciando extração de imagens...")
    
    # 1. Obter nomes dos produtos
    df = pd.read_excel('Produtos_tigre_pedido_exemplo.xlsx')
    df = df.dropna(subset=['PRODUTOS', 'PREÇO_UN', 'PREÇO_FD'])
    df = df[~df['PRODUTOS'].str.contains('Total|peso total|Frete|Seguro|ag. frete|Amarra', case=False, na=False)]
    produtos = df['PRODUTOS'].str.strip().tolist()
    
    # 2. Criar pasta
    pasta = 'fotos_de_produtos'
    os.makedirs(pasta, exist_ok=True)
    
    # 3. Processar PDF
    prod_idx = 0
    
    with pdfplumber.open('Catálogo_com_preços.pdf') as pdf:
        for p_idx, page in enumerate(pdf.pages):
            if prod_idx >= len(produtos):
                break
                
            words = page.extract_words()
            y_coords = [w['top'] for w in words if 'R$' in w['text']]
            
            # Agrupar Y coords (cada produto tem 2 preços, então 2 coords Y próximas)
            clusters = []
            for y in sorted(y_coords):
                if not clusters:
                    clusters.append([y])
                elif y - clusters[-1][-1] < 40:
                    clusters[-1].append(y)
                else:
                    clusters.append([y])
            
            row_centers = [sum(c)/len(c) for c in clusters]
            
            for y_center in row_centers:
                if prod_idx >= len(produtos):
                    break
                    
                # A imagem do produto fica à esquerda. Reduzimos a margem direita (de 360 para 270)
                # para não pegar informações irrelevantes ou o começo dos preços/descrição.
                bbox = (30, max(0, y_center - 100), 270, min(1440, y_center + 100))
                cropped = page.crop(bbox)
                
                img = cropped.to_image(resolution=150)
                
                # Nome do arquivo seguro
                nome = produtos[prod_idx]
                nome_arquivo = "".join([c for c in nome if c.isalpha() or c.isdigit() or c==' ']).rstrip()
                
                caminho = os.path.join(pasta, f"{nome_arquivo}.png")
                img.save(caminho)
                
                prod_idx += 1
                print(f"[{prod_idx}/{len(produtos)}] Salvo: {caminho}")

if __name__ == "__main__":
    extract_images()
    print("Processo concluído!")
