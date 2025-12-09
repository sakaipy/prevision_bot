from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time

def realizar_login(driver, user, password):
    """Realiza login no site Prevision (fluxo em duas etapas)"""
    wait = WebDriverWait(driver, 10)
    try:
        # ===== Etapa 1: Preencher e enviar o e-mail =====
        print("🪄 Inserindo e-mail...")
        print(password)
        email_input = wait.until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/main/div/div/div[1]/div/div/div[3]/div/form/div[1]/div/div/div[1]/div/div[1]/div[2]/input"))
        )
        email_input.clear()
        email_input.send_keys(user)

        botao_proximo = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/main/div/div/div[1]/div/div/div[3]/div/form/div[2]/div[2]/button"))
        )
        botao_proximo.click()
        print("📩 E-mail enviado, aguardando campo de senha...")

        # ===== Etapa 2: Aguardar campo de senha aparecer =====
        senha_input = wait.until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/main/div/div/div[1]/div/div/div[3]/div/form/div[1]/div/div[2]/div[2]/div/div[1]/div[2]/input"))
        )
        senha_input.clear()
        senha_input.send_keys(password)

        # ===== Etapa 3: Enviar senha e concluir login =====
        botao_login = wait.until(
            EC.element_to_be_clickable((By.CLASS_NAME, "v-btn__content"))
        )
        botao_login.click()
        print("🔓 Senha enviada, autenticando...")

        # ===== Etapa 4: Confirmar login bem-sucedido =====
        time.sleep(5)
        print("✅ Login concluído com sucesso (aguardando redirecionamento).")

    except TimeoutException:
        print("❌ Tempo excedido: algum elemento de login não foi encontrado.")
    except NoSuchElementException:
        print("⚠️ Elemento de login ausente. Verifique os seletores (ID ou CSS).")
    except Exception as e:
        print(f"💥 Erro inesperado no login: {e}")