from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from difflib import SequenceMatcher
import unicodedata
import re
import openpyxl
import difflib
import pandas as pd
import datetime
import time
import locale

def navegar_para_medicao(driver):
    """Executa a navegação até a página de Medição dentro do Prevision"""
    wait = WebDriverWait(driver, 20)
    
    try:

        print("🏗️ Acessando menu 'Obra'...")
        obra_card = wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//span[contains(text(), 'Obra')]/ancestor::div[contains(@class, 'v-card--link')]"
            ))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", obra_card)
        obra_card.click()

        print("🏡 Clicando no botão 'Selecionar' do projeto desejado...")
        print("🎯 Localizando botão 'Selecionar' da obra correta...")

        # Espera o DOM se estabilizar — garante que os cards já renderizaram
        for tentativa in range(10):  # até 10 tentativas com pequeno intervalo dinâmico
            try:
                obra_card = WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        "//span[@data-cy='project-card-name' and "
                        "contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'jardins cannes casas')]"
                    ))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", obra_card)
                print(f"✅ Card da obra encontrado (tentativa {tentativa + 1}).")

                # Rebusca o botão dentro do card — algumas renderizações atrasam esse elemento
                selecionar_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//span[@data-cy='project-card-name' and "
                        "contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'jardins cannes casas')]"
                        "/ancestor::div[contains(@class,'card-outter')]//button[.//span[contains(text(),'Selecionar')]]"
                    ))
                )

                driver.execute_script("arguments[0].scrollIntoView(true);", selecionar_btn)
                driver.execute_script("arguments[0].click();", selecionar_btn)
                print("✅ Botão 'Selecionar' clicado com sucesso!")
                break

            except Exception as e:
                print(f"⚠️ Tentativa {tentativa + 1}: botão ainda não disponível ({e}).")
                time.sleep(5)

        else:
            print("💥 Não foi possível encontrar o botão 'Selecionar' após múltiplas tentativas.")
        time.sleep(5)
        print("⏳ Aguardando carregamento da página...")
        WebDriverWait(driver, 20).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        time.sleep(10)
        print("✅ Página carregada, continuando...")

        """

        print("📑 Abrindo menu lateral (três barras)...")
        menu_btn = wait.until(
            EC.element_to_be_clickable((
                By.CSS_SELECTOR,
                "button[data-cy='features-list-scheduler-button']"
            ))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", menu_btn)
        driver.execute_script("arguments[0].click();", menu_btn)
        print("✅ Menu lateral aberto com sucesso!")

        print("🧩 Acessando seção 'Versões'...")

        try:
            # 🧠 Seletor primário – usa o texto visível 'Versões'
            versoes_option = wait.until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//div[contains(@class, 'v-list-item__content')]//span[normalize-space(text())='Versões']"
                ))
            )
        except:
            # 🪄 Fallback — busca o texto 'Versões' em qualquer local dentro do menu lateral
            print("⚠️ Fallback: tentando localizar 'Versões' por outro seletor...")
            versoes_option = wait.until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//span[normalize-space(text())='Versões']"
                ))
            )

        driver.execute_script("arguments[0].scrollIntoView(true);", versoes_option)
        driver.execute_script("arguments[0].click();", versoes_option)

        print("✅ Seção 'Versões' aberta com sucesso!")
        WebDriverWait(driver, 50).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        time.sleep(20)

        print("🧭 Expandindo lista 'Cenários'...")
        cenarios_btn = WebDriverWait(driver, 50).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//div[@role='button' and .//span[normalize-space(text())='Cenários']]"
            ))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", cenarios_btn)
        driver.execute_script("arguments[0].click();", cenarios_btn)
        print("✅ 'Cenários' expandido!")

        # Espera a lista carregar e encontra a opção 'Medição'
        print("🔍 Procurando item que contenha 'Medição' dentro de 'Cenários'...")
        medicao_item = WebDriverWait(driver, 50).until(
            EC.presence_of_element_located((
                By.XPATH,
                "//div[contains(@class,'v-list-item__title') and contains(translate(., 'MEDIÇÃO', 'medição'), 'medição')]"
            ))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", medicao_item)

        # Faz hover para revelar o botão oculto
        print("🖱️ Passando o mouse sobre o item de medição para revelar botão de restauração...")
        ActionChains(driver).move_to_element(medicao_item).perform()

        # Aguarda e clica no botão de restaurar (ícone mdi-restore)
        print("♻️ Clicando no botão de restaurar versão...")
        restore_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//button[.//i[contains(@class,'mdi-restore')]]"
            ))
        )
        driver.execute_script("arguments[0].click();", restore_button)
        print("✅ Versão de medição restaurada com sucesso!")
        
        print("🪄 Aguardando janela de confirmação aparecer...")
        confirm_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//span[normalize-space(text())='Restaurar']"
            ))
        )

        print("⚙️ Confirmando restauração...")
        driver.execute_script("arguments[0].scrollIntoView(true);", confirm_btn)
        driver.execute_script("arguments[0].click();", confirm_btn)
        print("✅ Restauração confirmada com sucesso!")
        time.sleep(5)
        print("⏳ Aguardando o fim do 'Atualizando projeto' (tempo variável)...")
        try:
            WebDriverWait(driver, 600).until(  # até 10 minutos se necessário
                EC.invisibility_of_element_located((
                    By.XPATH,
                    "//div[contains(@class, 'v-alert') and contains(., 'Atualizando projeto')]"
                ))
            )
            print("✅ Loading 'Atualizando projeto' desapareceu.")
        except:
            print("⚠️ Timeout: o alerta pode ter mudado de estrutura — prosseguindo mesmo assim...")

        # 🔁 Tenta localizar e clicar em "Selecionar" até conseguir
        print("🎯 Tentando localizar o botão 'Selecionar' da obra correta (espera dinâmica)...")
        for tentativa in range(30):  # até ~5 minutos, 10s entre tentativas
            try:
                obra_card = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        "//span[@data-cy='project-card-name' and "
                        "contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'jardins cannes casas')]"
                    ))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", obra_card)

                selecionar_btn = obra_card.find_element(
                    By.XPATH,
                    ".//ancestor::div[contains(@class,'card-outter')]"
                    "//button[.//span[contains(text(),'Selecionar')]]"
                )

                driver.execute_script("arguments[0].click();", selecionar_btn)
                print(f"✅ Botão 'Selecionar' clicado com sucesso! (tentativa {tentativa + 1})")
                break

            except Exception as e:
                print(f"⚠️ Tentativa {tentativa + 1}: obra ainda carregando ({type(e).__name__}). Aguardando 10s...")
                time.sleep(10)

        else:
            print("💥 Não foi possível reabrir a obra após várias tentativas. Verifique o carregamento no Prevision.")

        # ✅ Confirma o carregamento da página principal da obra
        WebDriverWait(driver, 120).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        """
        print("✅ Obra 'Jardins Cannes Casas' reaberta e página carregada com sucesso!")
        print("📄 Aguardando finalização do carregamento da página...")
        WebDriverWait(driver, 120).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        print("✅ Página principal carregada. Acessando 'Medições'...")

        # 🔍 Clicar na opção lateral 'Medições'
        try:
            medicoes_link = WebDriverWait(driver, 120).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//a[contains(@href, '/app/measurements')]"
                ))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", medicoes_link)
            driver.execute_script("arguments[0].click();", medicoes_link)
            print("✅ Opção lateral 'Medições' clicada com sucesso!")
        except Exception as e:
            print(f"💥 Erro ao clicar em 'Medições': {e}")

        # ⏳ Espera a página de medições carregar
        try:
            WebDriverWait(driver, 120).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//button[.//span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'criar medição')]]"
                ))
            )
            print("✅ Página de medições carregada com sucesso!")
        except:
            print("⚠️ A página de medições demorou para carregar. Tentando continuar mesmo assim...")

        # 🧩 Clicar no botão "Criar medição"
        try:
            print("🧱 Procurando botão 'Criar medição'...")
            criar_medicao_btn = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//button[.//span[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'criar medição')]]"
                ))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", criar_medicao_btn)
            driver.execute_script("arguments[0].click();", criar_medicao_btn)
            print("✅ Botão 'Criar medição' clicado com sucesso!")
        except Exception as e:
            print(f"💥 Erro ao clicar em 'Criar medição': {e}")
        try:
            locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
        except:
            locale.setlocale(locale.LC_TIME, 'pt_BR')

        print("📅 Aguardando exibição do calendário de medições...")
        calendar = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'v-calendar-monthly')]"))
        )
        print("✅ Calendário aberto com sucesso.")

        # 🧮 Calcular a última segunda-feira antes ou igual ao dia atual
        hoje = datetime.date.today()
        dias_para_voltar = (hoje.weekday() - 0) % 7  # segunda = 0
        ultima_segunda = hoje - datetime.timedelta(days=dias_para_voltar)
        dia_segunda = ultima_segunda.day
        mes_abrev = ultima_segunda.strftime('%b').lower().replace('.', '')[:3]  # exemplo: 'dez'
        print(f"📆 Última segunda-feira detectada: {ultima_segunda.strftime('%d/%m/%Y')}")

        # 🔍 Gerar possíveis formatos de texto do botão
        possiveis_textos = [
            f"{dia_segunda}",              # Ex: '2'
            f"{mes_abrev}. {dia_segunda}", # Ex: 'dez. 1'
            f"{mes_abrev} {dia_segunda}",  # Ex: 'dez 1' (fallback sem ponto)
        ]

        botao_dia = None
        for texto in possiveis_textos:
            try:
                print(f"🔍 Tentando localizar botão com texto '{texto}'...")
                xpath_botao = f"//button[.//span[contains(normalize-space(.), '{texto}')]]"
                botao_dia = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, xpath_botao))
                )
                break
            except:
                continue

        if botao_dia:
            driver.execute_script("arguments[0].scrollIntoView(true);", botao_dia)
            driver.execute_script("arguments[0].click();", botao_dia)
            print(f"✅ Clique efetuado na última segunda-feira ({ultima_segunda.strftime('%d/%m/%Y')}).")
        else:
            print(f"💥 Não foi possível localizar o botão da última segunda-feira ({ultima_segunda.strftime('%d/%m/%Y')}).")

        print("⏳ Aguardando carregamento das informações da medição...")

        try:
            # Espera o spinner aparecer (caso ainda não tenha carregado)
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "//circle[contains(@class, 'v-progress-circular__overlay')]"
                ))
            )
            print("🔁 Loading detectado, aguardando ele desaparecer...")

            # Espera até o spinner sumir completamente
            WebDriverWait(driver, 300).until(
                EC.invisibility_of_element_located((
                    By.XPATH,
                    "//circle[contains(@class, 'v-progress-circular__overlay')]"
                ))
            )
            print("✅ Loading finalizado, informações carregadas com sucesso!")
        except Exception as e:
            print(f"⚠️ Nenhum loading detectado ou timeout atingido ({e}). Continuando assim mesmo...")

        print("📋 Aguardando tabela de medições carregar...")

        container = WebDriverWait(driver, 600).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'v-data-table__wrapper')]"))
        )
        print("✅ Tabela de medições carregada com sucesso!")

        # Localizar o seletor "Linhas por página"
        print("🔽 Localizando seletor 'Linhas por página'...")
        seletor = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'v-select__slot')]"))
        )

        # Clicar no seletor para abrir o menu
        driver.execute_script("arguments[0].scrollIntoView(true);", seletor)
        driver.execute_script("arguments[0].click();", seletor)
        print("📂 Menu de linhas por página aberto.")

        # Esperar aparecer a opção "Todos"
        print("🔍 Aguardando opção 'Todos' aparecer...")
        opcao_todos = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'v-list-item__title') and contains(text(), 'Todos')]"))
        )

        # Clicar na opção "Todos"
        driver.execute_script("arguments[0].scrollIntoView(true);", opcao_todos)
        driver.execute_script("arguments[0].click();", opcao_todos)
        print("✅ Selecionada opção 'Todos' nas linhas por página.")

        def clicar_elemento(driver, elemento):
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
                elemento.click()
                return True
            except:
                pass
            try:
                driver.execute_script("arguments[0].click();", elemento)
                return True
            except:
                pass
            try:
                ActionChains(driver).move_to_element(elemento).pause(0.2).click().perform()
                return True
            except:
                return False
            
        df = pd.read_excel("data/cronograma.xlsx")
        df = df[['Pacote de trabalho/tarefas', 'Lote', 'Realizado', 'serviço']].dropna()
        print(f"📖 {len(df)} linhas carregadas da planilha com sucesso.")

        # === 1. Função para encontrar melhor correspondência no Excel ===

        def encontrar_valor_excel(nome_tela, df_lote):
            melhor_match = None
            melhor_similaridade = 0
            valor = None

            for _, linha in df_lote.iterrows():
                # 🧩 Pega o valor da planilha
                nome_excel = str(linha['Pacote de trabalho/tarefas']).strip()
                realizado = linha['Realizado']
                servico = linha['serviço']

                # 🧠 Ignora pacotes cujo valor realizado é 0 (ou equivalente)
                if realizado in [0, 0.0, "0", "0.0", "0,0"]:
                    continue

                # Calcula similaridade entre nomes
                similaridade = SequenceMatcher(None, nome_tela.lower(), nome_excel.lower()).ratio()
                if similaridade > melhor_similaridade:
                    melhor_similaridade = similaridade
                    melhor_match = nome_excel
                    valor = realizado

            # Retorna o melhor match se a similaridade for aceitável
            if melhor_similaridade >= 0.75:  # tolerância ajustável
                return valor, melhor_match, melhor_similaridade
            return None, None, 0
        
        #=== Função para normalizar valor de realizado ===

        def normalizar_realizado(valor):
            """Normaliza o valor de realizado para o formato 0–100 (inteiro)."""
            try:
                if isinstance(valor, str):
                    valor = valor.replace('%', '').strip()
                valor = float(valor)
                if 0 < valor <= 1:  # exemplo: 0.8 -> 80
                    return round(valor * 100, 2)
                elif valor > 100:  # evita erro caso alguém tenha 10000%
                    return 100.0
                return round(valor, 2)
            except Exception:
                return 0.0
            
        # === 2. Percorrer cada linha da planilha ===
        def preencher_input(inp, valor_realizado):
            try:
                # 🧩 Ignora inputs que não têm o símbolo de porcentagem próximo
                try:
                    suffix_text = inp.find_element(
                        By.XPATH,
                        "ancestor::div[contains(@class,'v-input')]//div[contains(@class,'v-text-field__suffix')]"
                    ).text
                    if "%" not in suffix_text:
                        print("⚪ Ignorado campo sem sufixo '%' (provável data).")
                        return True
                except:
                    print("⚪ Campo sem sufixo '%' — ignorado (provável data).")
                    return True

                # 🔒 Garante que o campo é editável
                if inp.get_attribute("readonly") or inp.get_attribute("disabled"):
                    print("⚪ Campo bloqueado (readonly/disabled) — ignorado.")
                    return True

                # ✏️ Preenche o valor
                inp.click()
                inp.clear()
                time.sleep(0.2)
                inp.send_keys(str(valor_realizado))
                time.sleep(0.2)
                print(f"✅ Campo aceitou valor → {valor_realizado}%")
                return True

            except Exception as e:
                print(f"⚠️ Erro ao preencher campo: {e}")
                return False
            
        def preencher_pacote(driver, pacote_span, valor_final):
            """Tenta preencher um pacote; se bloqueado, abre e preenche os subitens."""
            try:
                pacote_btn = pacote_span.find_element(By.XPATH, "./ancestor::button[contains(@class,'v-expansion-panel-header')]")

                # Tenta preencher diretamente se houver input no pacote principal
                inputs = pacote_btn.find_elements(By.XPATH, ".//input[@type='text']")
                preencheu = False
                if inputs:
                    for inp in inputs:
                        if preencher_input(inp, valor_realizado):
                            preencheu = True

                # Se não conseguiu preencher, abre o pacote e tenta os subitens
                if not preencheu:
                    driver.execute_script("arguments[0].scrollIntoView(true);", pacote_btn)
                    clicar_elemento(driver, pacote_btn)
                    time.sleep(0.7)

                    sub_inputs = driver.find_elements(
                        By.XPATH,
                        "//div[contains(@class,'v-expansion-panel--active')]//div[contains(@class,'job-row')]//input[@type='text']"
                    )

                    if sub_inputs:
                        print(f"↳ {len(sub_inputs)} subitens encontrados dentro de '{pacote_span.text.strip()}'.")
                        for inp in sub_inputs:
                            preencher_input(inp, valor_realizado)
                    else:
                        print(f"⚪ Nenhum subitem editável encontrado dentro de '{pacote_span.text.strip()}'.")
            except Exception as e:
                print(f"❌ Erro ao tentar preencher pacote '{pacote_span.text.strip()}': {e}")    



                
        for lote_excel in df['Lote'].unique():
            print(f"\n🏗️ Acessando Lote {lote_excel}...")

            # Abre o lote
            try:
                lote_xpath = f"//button[.//span[contains(normalize-space(.), '{lote_excel}')]]"
                lote_btn = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, lote_xpath)))
                clicar_elemento(driver, lote_btn)
                time.sleep(1.5)
                print(f"✅ Lote {lote_excel} expandido com sucesso.")
            except Exception as e:
                print(f"💥 Erro ao abrir lote {lote_excel}: {e}")
                continue

            # Coleta todos os pacotes visíveis no lote
            try:
                pacotes_visiveis = driver.find_elements(By.XPATH, "//span[contains(@class,'text-body-2') and contains(@class,'text-truncate')]")
                print(f"🔍 {len(pacotes_visiveis)} pacotes encontrados dentro do lote {lote_excel}.")
            except Exception as e:
                print(f"💥 Erro ao localizar pacotes visíveis: {e}")
                continue

            # Filtra planilha apenas para o lote atual
            subset_df = df[df['Lote'].astype(str).str.strip() == lote_excel.strip()]

            # Loop em cada pacote visível
            for pacote_span in pacotes_visiveis:
                nome_tela = pacote_span.text.strip()
                if not nome_tela:
                    continue

                # procura o melhor match dentro do lote
                valor_realizado, pacote_excel, similaridade = encontrar_valor_excel(nome_tela, subset_df)

                if valor_realizado is None:
                    print(f"⚪ Nenhum match encontrado para '{nome_tela}' (similaridade baixa)")
                    continue

                print(f"📦 '{nome_tela}' ≈ '{pacote_excel}' ({similaridade:.2f}) → {valor_realizado}%")

                try:
                    pacote_btn = pacote_span.find_element(By.XPATH, "./ancestor::button[contains(@class,'v-expansion-panel-header')]")
                    inputs = pacote_btn.find_elements(
                        By.XPATH,
                        ".//div[contains(@class,'v-input') and .//div[normalize-space(text())='%']]//input[@type='text']"
                    )
                    if inputs:
                        for inp in inputs:
                            driver.execute_script("arguments[0].scrollIntoView(true);", inp)
                            inp.clear()
                            valor_final = normalizar_realizado(valor_realizado)
                            preencher_pacote(driver, pacote_span, valor_final)
                        print(f"✅ Pacote '{nome_tela}' do lote '{lote_excel}' preenchido diretamente com valor '{valor_final}'.")
                        continue

                    # Caso não tenha input direto, tenta expandir e preencher subtarefas
                    clicar_elemento(driver, pacote_btn)
                    time.sleep(0.5)
                    sub_inputs = driver.find_elements(
                        By.XPATH,
                        "//div[contains(@class,'v-expansion-panel--active')]"
                        "//div[contains(@class,'job-row')]"
                        "//div[contains(@class,'v-input') and .//div[normalize-space(text())='%']]//input[@type='text']"
                    )
                    for inp in sub_inputs:
                        driver.execute_script("arguments[0].scrollIntoView(true);", inp)
                        inp.clear()
                        preencher_pacote(driver, pacote_span, valor_final)
                    print(f"✅ Pacote '{nome_tela}' expandido e {len(sub_inputs)} subitens preenchidos.")
                except Exception as e:
                    print(f"⚠️ Erro ao preencher '{nome_tela}': {e}")
                    continue

            print(f"🏁 Finalizado o lote {lote_excel}.")
            print("-" * 60)

    except Exception as e:
        print(f"💥 Erro durante navegação: {e}")