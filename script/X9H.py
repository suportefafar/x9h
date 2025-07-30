import os
import sys
import platform
import subprocess
import datetime
import json
import requests
import psutil
import socket
import uuid
import cpuinfo  # pip install py-cpuinfo
import urllib3
from PyQt5 import QtWidgets, QtCore, QtGui  # Adicionado QtGui para QPixmap no futuro, se necessário
import re


# --- Determinar Caminho Base para Arquivos de Dados (Importante para Executável) ---
def get_base_path():
    """ Retorna o caminho base para arquivos de dados, seja rodando como script ou como bundle. """
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
    return application_path


BASE_APP_PATH = get_base_path()

# --- Constantes ---
DATA_REGISTRO_FILE = os.path.join(BASE_APP_PATH, "ultimo_envio.txt")
USER_DATA_FILE = os.path.join(BASE_APP_PATH, "user_data.json")
API_POST_URL = "https://intranet.farmacia.ufmg.br/wp-json/intranet/v1/submission"
API_GET_PLACES_URL = "https://intranet.farmacia.ufmg.br/wp-json/intranet/v1/submissions/object/place"
API_GET_CONFIG_PATRIMONIO_URL = "https://intranet.farmacia.ufmg.br/wp-json/intranet/v1/submissions/equipaments/?client=x9h&type=computador,netbook,notebook"
API_GET_USERS_URL = "https://intranet.farmacia.ufmg.br/wp-json/intranet/v1/users/"

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# --- Funções Auxiliares para GPU ---
def get_gpu_windows():
    try:
        process_creation_flags = subprocess.CREATE_NO_WINDOW
        result = subprocess.run(
            ["wmic", "path", "Win32_VideoController", "get", "Name"],
            capture_output=True, text=True, check=True, creationflags=process_creation_flags, encoding='cp850',
            errors='ignore'
        )
        gpus = [line.strip() for line in result.stdout.splitlines() if line.strip() and line.strip().lower() != "name"]
        return ", ".join(gpus) if gpus else "N/A"
    except Exception as e:
        print(f"DEBUG: Erro ao obter GPU (Windows): {e}")
        return "N/A"


def get_gpu_linux():
    gpu_name = "N/A"
    try:
        result_glx = subprocess.run(["glxinfo"], capture_output=True, text=True, check=False)
        if result_glx.returncode == 0:
            for line in result_glx.stdout.splitlines():
                if "OpenGL renderer string" in line:
                    gpu_name = line.split(":", 1)[1].strip()
                    return gpu_name
    except FileNotFoundError:
        print("DEBUG: glxinfo não encontrado.")
    except Exception as e_glx:
        print(f"DEBUG: Erro com glxinfo: {e_glx}.")

    if gpu_name == "N/A":
        try:
            result_lspci = subprocess.run(["lspci"], capture_output=True, text=True, check=True)
            gpus_lspci = []
            for line in result_lspci.stdout.splitlines():
                if "VGA compatible controller" in line or "3D controller" in line or "Display controller" in line:
                    parts = line.split(":", 2)
                    gpu_desc = parts[-1].strip() if len(parts) > 2 else line.strip()
                    gpu_desc = re.sub(r'^\[.*?\]\s*', '', gpu_desc)
                    gpu_desc = re.sub(r'\s*\(rev .*\)\s*$', '', gpu_desc)
                    if gpu_desc:
                        gpus_lspci.append(gpu_desc.strip())
            if gpus_lspci:
                return ", ".join(list(set(gpus_lspci)))
        except FileNotFoundError:
            print("DEBUG: lspci não encontrado.")
        except Exception as e_lspci:
            print(f"DEBUG: Erro ao obter GPU (Linux - lspci): {e_lspci}")
    return gpu_name


def get_gpu_macos():
    try:
        result = subprocess.run(["system_profiler", "SPDisplaysDataType"], capture_output=True, text=True, check=True)
        gpus = []
        for line in result.stdout.splitlines():
            stripped_line = line.strip()
            if stripped_line.startswith("Chipset Model:"):
                gpu_name = stripped_line.split(":", 1)[1].strip()
                if gpu_name not in gpus:
                    gpus.append(gpu_name)
        return ", ".join(gpus) if gpus else "N/A"
    except Exception as e:
        print(f"DEBUG: Erro ao obter GPU (macOS): {e}")
        return "N/A"


# --- Funções Auxiliares para Tipo de Disco (Raiz) ---
def get_disk_type_windows():
    try:
        process_creation_flags = subprocess.CREATE_NO_WINDOW
        ps_command = "(Get-PhysicalDisk (Get-Partition -DriveLetter C).DiskNumber).MediaType"
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_command],
            capture_output=True, text=True, check=True, creationflags=process_creation_flags
        )
        media_type_output = result.stdout.strip().upper()
        if media_type_output == "3" or media_type_output == "HDD":
            return "HDD"
        elif media_type_output == "4" or media_type_output == "SSD":
            return "SSD"
        elif media_type_output == "5":
            return "SSD (SCM)"
        elif media_type_output == "0" or media_type_output == "UNSPECIFIED":
            return "Desconhecido (Unspecified)"
        else:
            return f"Desconhecido ({media_type_output})"
    except Exception as e:
        print(f"DEBUG: Erro ao obter tipo de disco (Windows): {e}")
        return "Desconhecido"


def get_disk_type_linux():
    try:
        df_output_result = subprocess.run(["df", "/"], capture_output=True, text=True, check=True)
        df_output = df_output_result.stdout
        lines = df_output.strip().splitlines()
        if len(lines) > 1:
            device_path = lines[1].split()[0]
            if device_path.startswith("/dev/"):
                dev_name_full = device_path.split('/')[-1]
                base_dev_name = dev_name_full
                match_sd_hd_vd = re.match(r"([svh]d[a-z]+)[0-9]*$", dev_name_full)
                if match_sd_hd_vd:
                    base_dev_name = match_sd_hd_vd.group(1)
                else:
                    match_nvme = re.match(r"(nvme[0-9]+n[0-9]+)p[0-9]+$", dev_name_full)
                    if match_nvme:
                        base_dev_name = match_nvme.group(1)
                    elif 'p' in base_dev_name and "nvme" in base_dev_name:
                        base_dev_name = base_dev_name.split('p')[0]

                if base_dev_name.startswith("nvme"):
                    return "SSD (NVMe)"

                rotational_file = f"/sys/block/{base_dev_name}/queue/rotational"
                if os.path.exists(rotational_file):
                    with open(rotational_file, "r") as f:
                        val = f.read().strip()
                    if val == "0":
                        return "SSD"
                    elif val == "1":
                        return "HDD"
        return "Desconhecido"
    except Exception as e:
        print(f"DEBUG: Erro ao obter tipo de disco (Linux): {e}")
        return "Desconhecido"


def get_disk_type_macos():
    try:
        result = subprocess.run(["diskutil", "info", "/"], capture_output=True, text=True, check=True)
        for line in result.stdout.splitlines():
            sl = line.strip()
            parts = sl.split(":", 1)
            if len(parts) > 1:
                key, value = parts[0].strip(), parts[1].strip()
                if key == "Solid State" and value.upper() == "YES": return "SSD"
                if key == "Solid State" and value.upper() == "NO": return "HDD"
                if key == "Medium Type" and "SOLID STATE" in value.upper(): return "SSD"
                if key == "Medium Type" and "ROTATIONAL" in value.upper(): return "HDD"
        return "Desconhecido (macOS)"
    except Exception as e:
        print(f"DEBUG: Erro ao obter tipo de disco (macOS): {e}")
        return "Desconhecido"


# --- Funções Principais de Coleta e Envio ---
def get_mac_address():
    mac_num = uuid.getnode()
    mac_hex = "%012X" % mac_num
    mac = ':'.join(mac_hex[i:i + 2] for i in range(0, 12, 2))
    if mac == "00:00:00:00:00:00":
        try:
            for _iface, snics in psutil.net_if_addrs().items():
                for snic in snics:
                    if snic.family == psutil.AF_LINK and snic.address and len(snic.address) == 17:
                        return snic.address.upper()
        except Exception:
            return "Não disponível"
    return mac if mac != "00:00:00:00:00:00" else "Não disponível"


def get_hardware_info():
    current_os = platform.system()
    gpu_info = "N/A"
    disk_type_info = "Desconhecido"

    if current_os == "Windows":
        gpu_info = get_gpu_windows()
        disk_type_info = get_disk_type_windows()
    elif current_os == "Linux":
        gpu_info = get_gpu_linux()
        disk_type_info = get_disk_type_linux()
    elif current_os == "Darwin":
        gpu_info = get_gpu_macos()
        disk_type_info = get_disk_type_macos()

    processador, ip_address, disk_total_str, ram_total_gb = "N/A", "N/A", "N/A", "N/A"
    cpu_cores, cpu_threads = "N/A", "N/A"
    try:
        processador = cpuinfo.get_cpu_info().get('brand_raw', "N/A")
    except Exception as e:
        print(f"DEBUG: Erro cpuinfo: {e}")
    try:
        ip_address = socket.gethostbyname(socket.gethostname())
    except Exception as e:
        print(f"DEBUG: Erro IP: {e}")
    try:
        disk_total_str = f"{round(psutil.disk_usage('/').total / (1024 ** 3), 2)} GB"
    except Exception as e:
        print(f"DEBUG: Erro disco total: {e}")
    try:
        ram_total_gb = f"{round(psutil.virtual_memory().total / (1024 ** 3), 2)} GB"
    except Exception as e:
        print(f"DEBUG: Erro RAM: {e}")
    try:
        cpu_cores = psutil.cpu_count(logical=False)
        cpu_threads = psutil.cpu_count(logical=True)
    except Exception as e:
        print(f"DEBUG: Erro Cores/Threads: {e}")

    return {
        "sistema": f"{platform.system()} {platform.release()}",
        "arquitetura": platform.architecture()[0],
        "nome_pc": socket.gethostname(),
        "ip": ip_address,
        "mac": get_mac_address(),
        "processador": processador,
        "gpu": gpu_info,
        "nucleos": cpu_cores if cpu_cores is not None else "N/A",
        "threads": cpu_threads if cpu_threads is not None else "N/A",
        "ram": ram_total_gb,
        "disco_total": disk_total_str,
        "tipo_disco_principal": disk_type_info
    }


def verificar_envio():  # Esta função não é mais usada no __main__ para auto-envio, mas pode ser útil
    hoje = datetime.datetime.now()
    if not os.path.exists(DATA_REGISTRO_FILE):
        return True  # Indica que seria um "novo envio" se a lógica antiga fosse usada
    try:
        with open(DATA_REGISTRO_FILE, "r") as f:
            ultima_data_str = f.read().strip()
        if not ultima_data_str:
            return True
        ultima_data = datetime.datetime.strptime(ultima_data_str, "%Y-%m-%d")
        return hoje.year > ultima_data.year or \
            (hoje.year == ultima_data.year and hoje.month > ultima_data.month)
    except Exception as e:
        print(f"DEBUG: Erro em verificar_envio: {e}")
        return True
    # return False # Implícito se nenhuma condição True e sem erro


def enviar_dados_post(user_data_with_ids, hardware_info):
    payload = {
        "object_name": "equipament",
        "data": json.dumps({**user_data_with_ids, **hardware_info})
    }
    headers = {"Content-Type": "application/json"}
    try:
        print(f"Enviando payload para {API_POST_URL}: {json.dumps(payload, indent=2)}")
        response = requests.post(API_POST_URL, json=payload, headers=headers, verify=False, timeout=30)
        if response.status_code in [200, 201]:
            with open(DATA_REGISTRO_FILE, "w") as f:
                f.write(datetime.datetime.now().strftime("%Y-%m-%d"))
            print("Dados enviados com sucesso!")
            return True
        else:
            print(f"Erro no envio. Status: {response.status_code}. Resposta: {response.text[:500]}")
            return False
    except requests.exceptions.RequestException as e:
        print(f"Erro na requisição POST: {e}")
        return False
    except Exception as e:
        print(f"Erro inesperado durante o envio POST: {e}")
        return False


def obter_patrimonios_para_combobox():
    """Obtém lista de patrimônios da API de CONFIGURAÇÃO de equipamentos."""
    try:
        print(f"Buscando patrimônios (ComboBox) de: {API_GET_CONFIG_PATRIMONIO_URL}")
        response = requests.get(API_GET_CONFIG_PATRIMONIO_URL, verify=False, timeout=15)
        response.raise_for_status()
        data = response.json()
        patrimonios_list = []
        seen_patrimonios = set()

        for item_config in data.get("results", []):
            if isinstance(item_config, dict):
                pat_num = None
                sub_data = item_config.get("data", {}) if isinstance(item_config.get("data"), dict) else item_config

                if "asset" in sub_data:
                    pat_num = sub_data.get("asset")
                elif "patrimonio" in sub_data:
                    pat_num = sub_data.get("patrimonio")

                pat_str = str(pat_num).strip() if pat_num is not None else None
                if pat_str and pat_str not in seen_patrimonios:
                    patrimonios_list.append({"label": pat_str, "id": pat_str})
                    seen_patrimonios.add(pat_str)

        patrimonios_list.sort(key=lambda x: x["label"])
        print(f"Patrimônios (de config API para ComboBox) encontrados: {len(patrimonios_list)}")
        return patrimonios_list
    except Exception as e:
        print(f"Erro obter patrimônios (ComboBox): {e}")
        return []


def obter_salas():
    try:
        response = requests.get(API_GET_PLACES_URL, verify=False, timeout=15)
        response.raise_for_status()
        dados = response.json()
        salas_results = dados.get("results", [])
        lista_salas = []
        for sala_item in salas_results:
            if isinstance(sala_item, dict):
                sala_data = sala_item.get("data", {})
                numero = sala_data.get("number", "").strip()
                descricao = sala_data.get("desc", "").strip()
                nome_sala_parts = [p for p in [numero, descricao] if p]
                nome_sala = " - ".join(nome_sala_parts)
                sala_id = sala_item.get("id")
                if nome_sala and sala_id is not None:
                    lista_salas.append({"label": nome_sala, "id": str(sala_id)})
        lista_salas.sort(key=lambda x: x["label"])
        return lista_salas
    except Exception as e:
        print(f"Erro obter salas: {e}")
        return []


def obter_usuarios():
    try:
        response = requests.get(API_GET_USERS_URL, verify=False, timeout=15)
        response.raise_for_status()
        data = response.json()
        usuarios = []
        results = data.get("results", [])
        for user_item in results:
            if isinstance(user_item, dict):
                nome_exibicao = user_item.get("display_name", "").strip()
                identificador = user_item.get("ID")
                if nome_exibicao and identificador is not None:
                    usuarios.append({"label": nome_exibicao, "id": str(identificador)})
        usuarios.sort(key=lambda x: x["label"])
        return usuarios
    except Exception as e:
        print(f"Erro obter usuários: {e}")
        return []


def carregar_configuracoes_por_patrimonio(patrimonio_id):
    """Busca configuração específica de UM patrimônio na API de CONFIGURAÇÃO."""
    if not patrimonio_id:
        return {}
    try:
        print(
            f"API Call: Buscando config. específica para patrimônio {patrimonio_id} via {API_GET_CONFIG_PATRIMONIO_URL}")
        response = requests.get(API_GET_CONFIG_PATRIMONIO_URL, verify=False, timeout=15)
        response.raise_for_status()
        dados_resposta = response.json()
        results = dados_resposta.get("results", [])

        for item_config in results:
            if isinstance(item_config, dict):
                config_data_main = item_config.get("data", {})
                current_asset_number = config_data_main.get("asset", config_data_main.get("patrimonio"))

                if str(current_asset_number) == str(patrimonio_id):
                    print(f"DEBUG: Configuração correspondente ENCONTRADA para patrimônio {patrimonio_id}.")
                    # print(f"DEBUG: Dados completos do item_config encontrado: {json.dumps(item_config, indent=2)}")

                    responsavel_label = None
                    sala_id = None

                    relationships = item_config.get("relationships", {})
                    if isinstance(relationships, dict):
                        applicant_data = relationships.get("applicant")
                        if isinstance(applicant_data, dict):
                            responsavel_label = applicant_data.get("display_name")

                        place_data_wrapper = relationships.get("place")
                        if isinstance(place_data_wrapper, dict):
                            sala_id = place_data_wrapper.get("id")
                    else:
                        print(f"DEBUG: 'relationships' não é um dicionário ou não encontrado para {patrimonio_id}.")

                    print(
                        f"DEBUG: Extraído da API para {patrimonio_id}: Responsável Label='{responsavel_label}', Sala ID='{sala_id}'")

                    return_data = {}
                    if responsavel_label:
                        return_data["responsavel_label"] = responsavel_label
                    if sala_id:
                        return_data["sala_id"] = str(sala_id)
                    return return_data

        print(f"Nenhuma config específica encontrada para {patrimonio_id} após filtrar {len(results)} resultados.")
        return {}
    except Exception as e:
        print(f"Erro CRÍTICO ao carregar config. do patrimônio '{patrimonio_id}' pela API: {e}")
        return {}


# --- Classe Trabalhadora para Envio em Background ---
class SubmissionWorker(QtCore.QObject):
    submission_success = QtCore.pyqtSignal(str)
    submission_failure = QtCore.pyqtSignal(str)
    finished = QtCore.pyqtSignal()

    def __init__(self, user_data_to_save_dict, parent=None):
        super().__init__(parent)
        self.user_data_to_save = user_data_to_save_dict
        self._is_running = True

    @QtCore.pyqtSlot()
    def run_submission(self):
        if not self._is_running:
            self.finished.emit()
            return
        try:
            print("WORKER THREAD: Coletando informações de hardware...")
            hardware_info = get_hardware_info()
            if not self._is_running: self.finished.emit(); return
            print("WORKER THREAD: Informações de hardware coletadas. Enviando dados...")
            if enviar_dados_post(self.user_data_to_save, hardware_info):
                if self._is_running: self.submission_success.emit(
                    "Dados salvos localmente e enviados com sucesso para o servidor!")
            else:
                if self._is_running: self.submission_failure.emit(
                    "Os dados foram salvos localmente, mas houve um erro ao enviá-los para o servidor.\n"
                    "Verifique sua conexão e os logs no console para detalhes."
                )
        except Exception as e:
            print(f"WORKER THREAD: Exceção na thread de submissão: {e}")
            if self._is_running: self.submission_failure.emit(f"Ocorreu um erro inesperado durante o envio: {e}")
        finally:
            self.finished.emit()

    def stop(self):  # Chamado se a janela principal tentar fechar durante o envio
        self._is_running = False


# --- Interface Gráfica ---
class FormDialog(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.submission_thread = None  # Para manter referência à thread
        self.submission_worker = None  # Para manter referência ao worker

        self.setWindowTitle("Registro de Equipamento - STI FAFARM")
        self.setFixedSize(480, 400)
        self.setStyleSheet("""
            QWidget { font-size: 10pt; } 
            QLabel { margin-bottom: 2px; }
            QComboBox { padding: 4px; margin-bottom: 8px; border: 1px solid #ccc; border-radius: 3px;}
            QPushButton { 
                background-color: #3498db; color: white; 
                padding: 10px 15px; border-radius: 5px; margin-top: 10px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #2980b9; }
            QFormLayout QLabel { font-weight: normal; }
        """)

        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QtWidgets.QVBoxLayout(central_widget)
        form_layout = QtWidgets.QFormLayout()
        form_layout.setSpacing(10)
        form_layout.setContentsMargins(15, 15, 15, 15)

        # Patrimônio
        self.combo_patrimonio = QtWidgets.QComboBox(self)
        self.lista_patrimonios = obter_patrimonios_para_combobox()
        self.combo_patrimonio.addItem("-- Selecione um Patrimônio --", None)
        for p_item in self.lista_patrimonios:
            self.combo_patrimonio.addItem(p_item["label"], p_item["id"])
        self.combo_patrimonio.currentIndexChanged.connect(self.on_patrimonio_selection_changed)
        form_layout.addRow(QtWidgets.QLabel("<b>Patrimônio:</b>"), self.combo_patrimonio)

        # Responsável
        self.combo_responsavel = QtWidgets.QComboBox(self)
        self.usuarios = obter_usuarios()
        self.combo_responsavel.addItem("-- Selecione um Responsável --", None)
        for usuario_item in self.usuarios:
            self.combo_responsavel.addItem(usuario_item["label"], usuario_item["id"])
        form_layout.addRow(QtWidgets.QLabel("<b>Responsável:</b>"), self.combo_responsavel)

        # Sala
        self.combo_sala = QtWidgets.QComboBox(self)
        self.salas = obter_salas()
        self.combo_sala.addItem("-- Selecione um Local/Sala --", None)
        for sala_item in self.salas:
            self.combo_sala.addItem(sala_item["label"], sala_item["id"])
        form_layout.addRow(QtWidgets.QLabel("<b>Local/Sala:</b>"), self.combo_sala)

        main_layout.addLayout(form_layout)
        main_layout.addStretch(1)

        self.status_label = QtWidgets.QLabel("Pronto.")
        self.status_label.setStyleSheet("font-style: italic; color: gray; padding: 5px;")
        main_layout.addWidget(self.status_label)

        self.btn_salvar = QtWidgets.QPushButton("Salvar e Enviar Dados")
        self.btn_salvar.setFixedHeight(40)
        self.btn_salvar.setCursor(QtCore.Qt.PointingHandCursor)
        self.btn_salvar.clicked.connect(self.salvar_e_enviar)
        main_layout.addWidget(self.btn_salvar)

        QtCore.QTimer.singleShot(150, self.carregar_dados_locais_ui)

    def on_patrimonio_selection_changed(self, index):
        patrimonio_id_selecionado = self.combo_patrimonio.itemData(index)
        current_pat_text = self.combo_patrimonio.currentText()
        self.status_label.setText(f"Patrimônio '{current_pat_text}' selecionado.")

        if patrimonio_id_selecionado is not None:
            self.tentar_carregar_config_patrimonio_ui(patrimonio_id_selecionado)
        else:
            self.combo_responsavel.setCurrentIndex(0)
            self.combo_sala.setCurrentIndex(0)
            self.status_label.setText("Selecione um patrimônio.")

    def carregar_dados_locais_ui(self):
        self.status_label.setText("Carregando dados locais...")
        QtWidgets.QApplication.processEvents()

        if not os.path.exists(USER_DATA_FILE):
            self.status_label.setText("Nenhum dado local salvo encontrado.")
            return
        try:
            with open(USER_DATA_FILE, "r") as f:
                dados_locais = json.load(f)

            patrimonio_local_id = str(dados_locais.get("patrimonio", "")).strip()
            if patrimonio_local_id:
                index_pat = self.combo_patrimonio.findData(patrimonio_local_id)
                if index_pat != -1:
                    self.combo_patrimonio.setCurrentIndex(index_pat)
                else:
                    print(f"DEBUG: Patrimônio local '{patrimonio_local_id}' não encontrado na lista.")
                    self.status_label.setText(f"Patrimônio local '{patrimonio_local_id}' não listado.")

            if self.combo_responsavel.currentIndex() <= 0:
                id_responsavel_local = str(dados_locais.get("responsavel", "")).strip()
                if id_responsavel_local:
                    index_resp = self.combo_responsavel.findData(id_responsavel_local)
                    if index_resp != -1:
                        self.combo_responsavel.setCurrentIndex(index_resp)

            if self.combo_sala.currentIndex() <= 0:
                id_sala_local = str(dados_locais.get("sala", "")).strip()
                if id_sala_local:
                    index_sala = self.combo_sala.findData(id_sala_local)
                    if index_sala != -1:
                        self.combo_sala.setCurrentIndex(index_sala)

            if patrimonio_local_id or self.combo_patrimonio.currentIndex() > 0:
                self.status_label.setText("Dados locais e/ou de API para patrimônio carregados.")
            else:
                self.status_label.setText("Dados locais carregados (sem seleção de patrimônio).")

        except FileNotFoundError:
            self.status_label.setText(f"Arquivo de dados local ({USER_DATA_FILE}) não encontrado.")
        except json.JSONDecodeError:
            self.status_label.setText(f"Erro ao ler {USER_DATA_FILE}. Arquivo pode estar corrompido.")
            print(f"DEBUG: Erro de decodificação JSON ao ler {USER_DATA_FILE}.")
        except Exception as e:
            print(f"DEBUG: Erro ao carregar dados locais para UI: {e}")
            self.status_label.setText("Erro inesperado ao carregar dados locais.")

    def tentar_carregar_config_patrimonio_ui(self, patrimonio_id):
        if not patrimonio_id:
            return
        self.status_label.setText(f"Buscando config. API para patrimônio {patrimonio_id}...")
        QtWidgets.QApplication.processEvents()

        config_api = carregar_configuracoes_por_patrimonio(patrimonio_id)

        if not config_api:
            self.status_label.setText(f"Nenhuma config. da API para {patrimonio_id} (ou erro na busca).")
            self.combo_responsavel.setCurrentIndex(0)
            self.combo_sala.setCurrentIndex(0)
            return

        self.status_label.setText(f"Config. API para {patrimonio_id} recebida.")
        print(f"DEBUG: config_api recebida em tentar_carregar_config_patrimonio_ui: {config_api}")

        responsavel_label_da_api = config_api.get("responsavel_label")
        if responsavel_label_da_api:
            print(f"DEBUG: Tentando selecionar Responsável pelo LABEL: '{responsavel_label_da_api}'")
            index_resp = self.combo_responsavel.findText(responsavel_label_da_api, QtCore.Qt.MatchFixedString)
            if index_resp != -1:
                self.combo_responsavel.setCurrentIndex(index_resp)
                print(f"DEBUG: SUCESSO - Responsável '{responsavel_label_da_api}' selecionado no índice {index_resp}.")
            else:
                print(
                    f"DEBUG: FALHA - Responsável '{responsavel_label_da_api}' (da API) não encontrado na lista do ComboBox.")
                self.combo_responsavel.setCurrentIndex(0)
        else:
            print("DEBUG: 'responsavel_label' não encontrado ou vazio na config_api.")
            self.combo_responsavel.setCurrentIndex(0)

        sala_id_da_api = config_api.get("sala_id")
        if sala_id_da_api:
            print(f"DEBUG: Tentando selecionar Sala pelo ID: '{sala_id_da_api}'")
            index_sala = self.combo_sala.findData(str(sala_id_da_api))
            if index_sala != -1:
                self.combo_sala.setCurrentIndex(index_sala)
                print(
                    f"DEBUG: SUCESSO - Sala com ID '{sala_id_da_api}' (label: '{self.combo_sala.itemText(index_sala)}') selecionada no índice {index_sala}.")
            else:
                print(
                    f"DEBUG: FALHA - Sala com ID '{sala_id_da_api}' (da API) não encontrada na lista do ComboBox (verifique os 'userData').")
                self.combo_sala.setCurrentIndex(0)
        else:
            print("DEBUG: 'sala_id' não encontrado ou vazio na config_api.")
            self.combo_sala.setCurrentIndex(0)

        self.status_label.setText(f"Campos atualizados por API para {patrimonio_id}.")

    def salvar_e_enviar(self):
        patrimonio_id = self.combo_patrimonio.currentData()
        responsavel_id = self.combo_responsavel.currentData()
        sala_id = self.combo_sala.currentData()

        erros = []
        if patrimonio_id is None: erros.append("O campo <b>Patrimônio</b> é obrigatório.")
        if responsavel_id is None: erros.append("O campo <b>Responsável</b> é obrigatório.")
        if sala_id is None: erros.append("O campo <b>Local/Sala</b> é obrigatório.")

        if erros:
            QtWidgets.QMessageBox.warning(self, "Campos Obrigatórios", "<br>".join(erros))
            return

        user_data_to_save = {
            "patrimonio": patrimonio_id,
            "responsavel": responsavel_id,
            "sala": sala_id
        }

        try:
            with open(USER_DATA_FILE, "w") as f:
                json.dump(user_data_to_save, f, indent=4)
            print(f"Dados locais salvos com sucesso em {USER_DATA_FILE}")
        except IOError as e:
            QtWidgets.QMessageBox.critical(self, "Erro ao Salvar Localmente",
                                           f"Não foi possível salvar os dados localmente:\n{e}")
            return

        self.btn_salvar.setEnabled(False)
        self.status_label.setText("Enviando dados para o servidor... Por favor, aguarde.")
        QtWidgets.QApplication.processEvents()

        if self.submission_thread and self.submission_thread.isRunning():
            print("DEBUG: Submissão anterior ainda em progresso. Cancelando nova tentativa.")  # Evita múltiplas threads
            self.btn_salvar.setEnabled(True)  # Reabilita o botão se a lógica chegar aqui por engano
            return

        self.submission_thread = QtCore.QThread(self)
        self.submission_worker = SubmissionWorker(user_data_to_save)
        self.submission_worker.moveToThread(self.submission_thread)

        self.submission_thread.started.connect(self.submission_worker.run_submission)
        self.submission_worker.submission_success.connect(self.handle_submission_success)
        self.submission_worker.submission_failure.connect(self.handle_submission_failure)
        self.submission_worker.finished.connect(self.submission_thread.quit)
        self.submission_thread.finished.connect(self.submission_worker.deleteLater)
        self.submission_thread.finished.connect(self.submission_thread.deleteLater)
        self.submission_thread.finished.connect(lambda: self.btn_salvar.setEnabled(True))  # Reabilita o botão sempre

        self.submission_thread.start()

    @QtCore.pyqtSlot(str)
    def handle_submission_success(self, message):
        self.status_label.setText("Envio concluído com sucesso.")
        QtWidgets.QMessageBox.information(self, "Sucesso", message)
        self.close()

    @QtCore.pyqtSlot(str)
    def handle_submission_failure(self, error_message):
        self.status_label.setText("Falha no envio. Verifique o console.")
        QtWidgets.QMessageBox.critical(self, "Falha no Envio", error_message)

    def closeEvent(self, event):
        print("DEBUG: Close event chamado na FormDialog")
        if self.submission_thread and self.submission_thread.isRunning():
            print("DEBUG: Tentando parar a thread de submissão antes de fechar...")
            if hasattr(self.submission_worker, 'stop'):  # Verifica se o worker tem o método stop
                self.submission_worker.stop()
            self.submission_thread.quit()
            if not self.submission_thread.wait(1000):  # Espera até 1 segundo
                print("DEBUG: A thread de submissão não parou a tempo. Forçando o fechamento.")
        super().closeEvent(event)


# --- Funções de Autoexecução (Ajustadas para Executável) ---
def registrar_autoexec():
    if platform.system() == "Windows":
        task_name = "HardwareMonitorUFMGSTI"
        executable_to_run_in_task = ""

        if getattr(sys, 'frozen', False):
            executable_to_run_in_task = f'"{sys.executable}"'
        else:
            python_exe = sys.executable
            script_file = os.path.abspath(__file__)
            executable_to_run_in_task = f'"{python_exe}" "{script_file}"'

        try:
            process_creation_flags = subprocess.CREATE_NO_WINDOW
            command = f'schtasks /Create /SC ONLOGON /TN "{task_name}" /TR {executable_to_run_in_task} /RL HIGHEST /F'
            print(f"Tentando criar/atualizar tarefa agendada: {command}")
            result = subprocess.run(command, shell=True, capture_output=True, text=True, check=False,
                                    creationflags=process_creation_flags)

            if result.returncode == 0:
                print(f"Tarefa agendada ('{task_name}') criada/atualizada com sucesso.")
            else:
                print(f"Erro ao criar/atualizar tarefa agendada (código {result.returncode}):")
                if result.stdout: print(f"  stdout: {result.stdout.strip()}")
                if result.stderr: print(f"  stderr: {result.stderr.strip()}")
        except Exception as e:
            print(f"Exceção ao criar/atualizar tarefa agendada: {e}")


def remover_autoexec():
    if platform.system() == "Windows":
        task_name = "HardwareMonitorUFMGSTI"
        try:
            process_creation_flags = subprocess.CREATE_NO_WINDOW
            command = f'schtasks /Delete /TN "{task_name}" /F'
            print(f"Tentando remover tarefa agendada: {command}")
            result = subprocess.run(command, shell=True, capture_output=True, text=True, check=False,
                                    creationflags=process_creation_flags)
            if result.returncode == 0:
                print(f"Tarefa agendada ('{task_name}') removida com sucesso.")
            else:
                print(f"Erro ao remover tarefa agendada (pode não existir ou requerer permissão):")
                if result.stdout: print(f"  stdout: {result.stdout.strip()}")
                if result.stderr: print(f"  stderr: {result.stderr.strip()}")
        except Exception as e:
            print(f"Exceção ao remover tarefa agendada: {e}")


# --- Ponto de Entrada Principal ---
if __name__ == "__main__":
    # Descomente para gerenciar a tarefa agendada
    # registrar_autoexec()
    # remover_autoexec()

    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")

    # A lógica de reenvio automático foi removida daqui, conforme solicitado.
    # O programa sempre iniciará a interface gráfica.

    print("Iniciando a interface gráfica...")
    dialog = FormDialog()
    dialog.show()
    sys.exit(app.exec_())