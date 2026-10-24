import os
import time
import threading
import base64
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import pandas as pd
import requests
from selenium import webdriver

class AutomacaoDAI:
    def __init__(self, root):
        self.root = root
        self.root.title("Automação GRU - Download de DAI")
        self.root.geometry("650x550")
        
        self.file_path = tk.StringVar()
        
        # Interface Simplificada
        tk.Label(root, text="1. Selecione a Planilha com as DUIMPs:", font=("Arial", 10, "bold")).pack(pady=(20, 5))
        frame_file = tk.Frame(root)
        frame_file.pack(pady=5, fill="x", padx=20)
        tk.Entry(frame_file, textvariable=self.file_path, state="readonly", width=55).pack(side="left", padx=(0, 10))
        tk.Button(frame_file, text="Procurar...", command=self.browse_file).pack(side="left")
        
        tk.Label(root, text="2. Clique em Iniciar (O Chrome abrirá para você fazer o login)", font=("Arial", 10, "bold"), fg="blue").pack(pady=(20, 5))
        
        tk.Button(root, text="INICIAR AUTOMAÇÃO", command=self.start_thread, bg="green", fg="white", font=("Arial", 11, "bold"), height=2, width=25).pack(pady=15)
        
        tk.Label(root, text="Log de Execução:", font=("Arial", 10, "bold")).pack(pady=(5, 0))
        self.log_area = scrolledtext.ScrolledText(root, width=75, height=15, state="disabled")
        self.log_area.pack(pady=5, padx=20)
        
    def log(self, message):
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state="disabled")
        self.root.update()

    def browse_file(self):
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file:
            self.file_path.set(file)

    def start_thread(self):
        if not self.file_path.get():
            messagebox.showwarning("Aviso", "Por favor, selecione a planilha.")
            return
            
        threading.Thread(target=self.run_automation, args=(self.file_path.get(),), daemon=True).start()

    def capturar_cookie_automatico(self):
        self.log("🌐 Abrindo o Chrome... Por favor, faça o login na tela que aparecer.")
        try:
            # Abre o Chrome moderno (O Selenium 4.6+ já gerencia o driver automaticamente)
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            driver = webdriver.Chrome(options=options)
            driver.get("https://www.tarifagru.com.br")
            
            cookie_str = ""
            timeout = 300  # O usuário tem até 5 minutos para logar
            start_time = time.time()
            
            # Loop monitorando se o login foi concluído
            while time.time() - start_time < timeout:
                try:
                    cookies = driver.get_cookies()
                    has_jsession = any(c.get('name') == 'JSESSIONID' for c in cookies)
                    has_vcap = any(c.get('name') == '__VCAP_ID__' for c in cookies)
                    
                    # Se gerou os cookies de sessão e saiu da tela de login, captura tudo!
                    if has_jsession and has_vcap and "login" not in driver.current_url.lower():
                        cookie_str = "; ".join([f"{c['name']}={c['value']}" for c in cookies])
                        self.log("✅ Login detectado! Sessão capturada.")
                        time.sleep(2) # Aguarda 2 segundos para o servidor SAP estabilizar o token
                        break
                except Exception:
                    pass
                time.sleep(1)
                
            driver.quit()
            return cookie_str
            
        except Exception as e:
            self.log(f"[ERRO CHROME] {str(e)}")
            return ""

    def run_automation(self, file_path):
        self.log("=== INICIANDO AUTOMAÇÃO ===")
        
        # Passo 1: Captura Automática do Cookie
        cookie = self.capturar_cookie_automatico()
        
        if not cookie:
            self.log("❌ Automação cancelada. O login não foi concluído ou a janela foi fechada.")
            return
            
        self.log("🚀 Iniciando processamento em segundo plano...")
        
        headers = {
            'accept': 'application/json',
            'accept-language': 'pt-BR',
            'cookie': cookie,
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'x-requested-with': 'XMLHttpRequest'
        }
        
        pasta_destino = os.path.join(os.path.dirname(file_path), "Guias_DAI_Baixadas")
        os.makedirs(pasta_destino, exist_ok=True)
        
        try:
            df_temp = pd.read_excel(file_path, header=None)
            try:
                idx_cabecalho = df_temp[df_temp.eq('REF DSV').any(axis=1)].index[0]
            except IndexError:
                idx_cabecalho = 2 
            df = pd.read_excel(file_path, header=idx_cabecalho)
        except Exception as e:
            self.log(f"[ERRO CRÍTICO] Falha ao ler a planilha: {e}")
            return
            
        sucessos, falhas = 0, 0
        
        for index, row in df.iterrows():
            ref_dsv = str(row.get('REF DSV', '')).strip()
            duimp_original = str(row.get('NUMERO DUIMP', '')).strip()
            
            if pd.isna(row.get('NUMERO DUIMP')) or not duimp_original or duimp_original.lower() == 'nan':
                self.log("\n[!] Fim da lista de DUIMPs encontrado. Parando a busca.")
                break
                
            duimp_limpa = duimp_original.replace('-', '')
            nome_arquivo = f"{ref_dsv}_DAI.pdf"
            
            self.log(f"Processando: {ref_dsv} | DUIMP: {duimp_limpa}")
            
            # Busca
            url_busca = f"https://www.tarifagru.com.br/dbc7ec03-2cbb-407a-a91b-b7f60d379a78.gru-airport-pc-teca.brcomgrupctecacockpit/~03de271b-f3bb-411b-8f9c-1a5507233415~/v2/gru/VIEW_PC_TECA_LIST?$skip=0&$top=50&$filter=substringof('{duimp_limpa}',LIBERATORY_DOCTO)"
            
            try:
                resp_busca = requests.get(url_busca, headers=headers)
                if resp_busca.status_code != 200:
                    self.log(f"  -> [ERRO] Falha de conexão. O sistema derrubou a sessão.")
                    falhas += 1
                    continue
                    
                resultados = resp_busca.json().get('d', {}).get('results', [])
                
                if not resultados:
                    self.log(f"  -> [NÃO ENCONTRADO] Guia indisponível no sistema.")
                    falhas += 1
                    continue
                    
                cid = resultados[0].get('ID') 
                if not cid:
                    self.log(f"  -> [ERRO] Falha ao gerar código interno da guia.")
                    falhas += 1
                    continue
                    
                # Download
                url_down = f"https://www.tarifagru.com.br/dbc7ec03-2cbb-407a-a91b-b7f60d379a78.gru-airport-pc-teca.brcomgrupctecacockpit/~03de271b-f3bb-411b-8f9c-1a5507233415~/SAP_GATEWAY/sap/opu/odata/sap/ZGWNB_BOLETOS_SRV/HeaderSet?$filter=CID%20eq%20%27{cid}%27"
                
                resp_down = requests.get(url_down, headers=headers)
                itens_down = resp_down.json().get('d', {}).get('results', [])
                
                if not itens_down:
                    self.log(f"  -> [ERRO] SAP não liberou o arquivo PDF.")
                    falhas += 1
                    continue
                    
                pdf_base64 = itens_down[0].get('PDF') 
                
                if not pdf_base64:
                    self.log(f"  -> [ERRO] Conteúdo do PDF veio corrompido ou vazio.")
                    falhas += 1
                    continue
                    
                # Salvar Arquivo
                caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)
                with open(caminho_arquivo, "wb") as f:
                    f.write(base64.b64decode(pdf_base64))
                    
                self.log(f"  -> [SUCESSO] Salvo!")
                sucessos += 1
                
            except Exception as e:
                self.log(f"  -> [ERRO] {str(e)}")
                falhas += 1
                
        resumo = f"\n=== RESUMO FINAL ===\nDownloads bem-sucedidos: {sucessos}\nFalhas/Não encontrados: {falhas}"
        self.log(resumo)
        
        with open(os.path.join(pasta_destino, "Log_Automacao_DAI.txt"), "w", encoding="utf-8") as f_log:
            f_log.write(self.log_area.get("1.0", tk.END))
        
        messagebox.showinfo("Concluído", "Processamento finalizado!")

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomacaoDAI(root)
    root.mainloop()