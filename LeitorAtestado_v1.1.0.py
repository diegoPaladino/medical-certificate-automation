import base64
import csv
import json
import os
import queue
import re
import shutil
import socket
import ssl
import tempfile
import threading
import time
import unicodedata
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import requests
from openpyxl import load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


APP_DIR = Path(__file__).resolve().parent
APP_VERSION = "1.1.0"
GEMUL_LOGIN_URL = "https://gemul-aparecida.com.br/login.asp"
GEMUL_ATESTADO_URL = "https://gemul-aparecida.com.br/app/sca_justificativa_faltas.asp"
OPENAI_MODEL = (os.getenv("OPENAI_MODEL") or "gpt-5-mini").strip() or "gpt-5-mini"
OPENAI_TIMEOUT = int(os.getenv("OPENAI_TIMEOUT", "90"))
OPENAI_TENTATIVAS = int(os.getenv("OPENAI_TENTATIVAS", "2"))
OPENAI_IMAGE_MAX_SIDE = int(os.getenv("OPENAI_IMAGE_MAX_SIDE", "1600"))
OPENAI_IMAGE_JPEG_QUALITY = int(os.getenv("OPENAI_IMAGE_JPEG_QUALITY", "78"))
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff"}
DOCUMENT_EXTENSIONS = IMAGE_EXTENSIONS | {".txt", ".pdf"}
WIA_FORMAT_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
LOG_FILE = APP_DIR / "leitor_atestado.log"


@dataclass
class Aluno:
    turma: str
    matricula: str
    nome: str

    @property
    def nome_arquivo(self) -> str:
        return f"{self.matricula} - {self.turma} - {self.nome}"


def normalizar_texto(valor: Any) -> str:
    texto = str(valor or "").strip().upper()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")
    texto = re.sub(r"\s+", " ", texto)
    return texto


def limpar_nome_arquivo(valor: str) -> str:
    texto = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", str(valor or ""))
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto.rstrip(". ")


def formatar_data_iso_para_br(valor: str) -> str:
    try:
        return datetime.strptime(valor, "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return valor or ""


def parse_data_usuario(valor: str) -> Optional[datetime]:
    texto = str(valor or "").strip()
    for formato in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(texto, formato)
        except ValueError:
            pass
    return None


def somar_dias(data_iso: str, dias: int) -> str:
    inicio = datetime.strptime(data_iso, "%Y-%m-%d").date()
    fim = inicio + timedelta(days=max(1, int(dias or 1)) - 1)
    return fim.isoformat()


def carregar_alunos() -> List[Aluno]:
    candidatos = [
        APP_DIR / "lista_alunos.csv",
        APP_DIR / "lista_alunos_com_codigo_turma.xlsx",
        APP_DIR / "lista_alunos.xlsx",
    ]
    caminho = next((p for p in candidatos if p.exists()), None)
    if caminho is None:
        raise FileNotFoundError(
            "Nao encontrei lista_alunos.csv, lista_alunos_com_codigo_turma.xlsx ou lista_alunos.xlsx na pasta do programa."
        )

    if caminho.suffix.lower() == ".csv":
        return carregar_alunos_csv(caminho)
    return carregar_alunos_xlsx(caminho)


def carregar_alunos_csv(caminho: Path) -> List[Aluno]:
    with caminho.open("r", encoding="utf-8-sig", newline="") as arquivo:
        amostra = arquivo.read(4096)
        arquivo.seek(0)
        dialect = csv.Sniffer().sniff(amostra, delimiters=";,")
        leitor = csv.DictReader(arquivo, dialect=dialect)
        alunos = []
        for linha in leitor:
            turma = linha.get("Turma") or linha.get("turma") or ""
            matricula = linha.get("Matricula") or linha.get("Matrícula") or linha.get("matricula") or ""
            nome = linha.get("Nome do Aluno") or linha.get("Nome") or linha.get("nome") or ""
            if turma and matricula and nome:
                alunos.append(Aluno(str(turma).strip(), re.sub(r"\D", "", str(matricula)), str(nome).strip().upper()))
        return alunos


def carregar_alunos_xlsx(caminho: Path) -> List[Aluno]:
    wb = load_workbook(caminho, read_only=True, data_only=True)
    ws = wb.active
    cabecalho = [normalizar_texto(v) for v in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]

    def indice_possivel(*nomes: str) -> int:
        nomes_norm = {normalizar_texto(nome) for nome in nomes}
        for i, col in enumerate(cabecalho):
            if col in nomes_norm:
                return i
        raise ValueError(f"Coluna nao encontrada em {caminho.name}: {', '.join(nomes)}")

    idx_turma = indice_possivel("Turma")
    idx_matricula = indice_possivel("Matricula", "Matrícula")
    idx_nome = indice_possivel("Nome do Aluno", "Nome")

    alunos: List[Aluno] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        turma = row[idx_turma] if idx_turma < len(row) else ""
        matricula = row[idx_matricula] if idx_matricula < len(row) else ""
        nome = row[idx_nome] if idx_nome < len(row) else ""
        if turma and matricula and nome:
            alunos.append(Aluno(str(turma).strip(), re.sub(r"\D", "", str(matricula)), str(nome).strip().upper()))
    return alunos


def extrair_texto_resposta_openai(resposta: Dict[str, Any]) -> str:
    texto_direto = resposta.get("output_text")
    if isinstance(texto_direto, str) and texto_direto.strip():
        return texto_direto.strip()
    textos = []
    for item in resposta.get("output", []):
        for conteudo in item.get("content", []):
            texto = conteudo.get("text")
            if isinstance(texto, str) and texto.strip():
                textos.append(texto.strip())
    return "\n".join(textos).strip()


def parse_json_seguro(texto: str) -> Dict[str, Any]:
    bruto = (texto or "").strip()
    if bruto.startswith("```"):
        bruto = re.sub(r"^```(?:json)?\s*", "", bruto)
        bruto = re.sub(r"\s*```$", "", bruto)
    return json.loads(bruto)


def registrar_log_arquivo(mensagem: str) -> None:
    try:
        carimbo = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with LOG_FILE.open("a", encoding="utf-8") as arquivo:
            arquivo.write(f"[{carimbo}] {mensagem}\n")
    except Exception:
        pass


def enviar_requisicao_openai(payload: Dict[str, Any], api_key: str, log=None) -> Dict[str, Any]:
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    codigos_retry = {408, 409, 425, 429, 500, 502, 503, 504}
    ultimo_erro: Optional[BaseException] = None

    for tentativa in range(1, OPENAI_TENTATIVAS + 1):
        try:
            if log:
                log(f"OpenAI: tentativa {tentativa}/{OPENAI_TENTATIVAS}.")
            resposta = requests.post(
                "https://api.openai.com/v1/responses",
                headers=headers,
                json=payload,
                timeout=OPENAI_TIMEOUT,
            )
            if tentativa < OPENAI_TENTATIVAS and resposta.status_code in codigos_retry:
                if log:
                    log(f"OpenAI retornou HTTP {resposta.status_code}; tentando novamente.")
                time.sleep(min(8, 2 ** (tentativa - 1)))
                continue
            resposta.raise_for_status()
            return resposta.json()
        except requests.RequestException as exc:
            ultimo_erro = exc
            if log:
                log(f"OpenAI: falha na tentativa {tentativa}: {exc}")
            if tentativa < OPENAI_TENTATIVAS:
                time.sleep(min(8, 2 ** (tentativa - 1)))
                continue

    dados = json.dumps(payload).encode("utf-8")
    contexto_ssl = ssl.create_default_context()
    try:
        contexto_ssl.minimum_version = ssl.TLSVersion.TLSv1_2
    except Exception:
        pass
    request = urllib.request.Request(
        "https://api.openai.com/v1/responses",
        data=dados,
        headers=headers,
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=OPENAI_TIMEOUT, context=contexto_ssl) as resposta:
            return json.loads(resposta.read().decode("utf-8"))
    except (urllib.error.URLError, socket.timeout, TimeoutError) as exc:
        raise RuntimeError(f"Falha de conexao com a OpenAI. Detalhe: {exc!r}") from ultimo_erro or exc


def preparar_imagem_para_api(caminho: Path, log=None) -> bytes:
    original = caminho.read_bytes()
    try:
        from PIL import Image, ImageOps
    except Exception:
        if log:
            log("Pillow nao disponivel; enviando imagem original para a API.")
        return original

    try:
        with Image.open(caminho) as imagem:
            imagem = ImageOps.exif_transpose(imagem)
            imagem.thumbnail((OPENAI_IMAGE_MAX_SIDE, OPENAI_IMAGE_MAX_SIDE))
            if imagem.mode not in {"RGB", "L"}:
                imagem = imagem.convert("RGB")
            with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as temp:
                temp_path = Path(temp.name)
            try:
                imagem.save(temp_path, format="JPEG", quality=OPENAI_IMAGE_JPEG_QUALITY, optimize=True)
                otimizada = temp_path.read_bytes()
            finally:
                try:
                    temp_path.unlink()
                except OSError:
                    pass
    except Exception as exc:
        if log:
            log(f"Nao foi possivel otimizar a imagem; enviando original. Detalhe: {exc}")
        return original

    if len(otimizada) >= len(original):
        return original
    if log:
        log(f"Imagem otimizada para API: {len(original) // 1024} KB -> {len(otimizada) // 1024} KB.")
    return otimizada


def preparar_imagem_para_gemul(caminho: Path, log=None) -> Path:
    if caminho.suffix.lower() not in IMAGE_EXTENSIONS:
        return caminho
    if caminho.stat().st_size <= 2 * 1024 * 1024:
        return caminho

    try:
        from PIL import Image, ImageOps
    except Exception:
        if log:
            log("Pillow nao disponivel; imagem original pode exceder o limite de 2MB do GEMUL.")
        return caminho

    limite_bytes = 2 * 1024 * 1024
    destino = caminho.with_name(f"{caminho.stem}_gemul_2mb.jpg")
    with Image.open(caminho) as imagem:
        imagem = ImageOps.exif_transpose(imagem)
        if imagem.mode not in {"RGB", "L"}:
            imagem = imagem.convert("RGB")

        lado_maximo = 1800
        while lado_maximo >= 700:
            tentativa = imagem.copy()
            tentativa.thumbnail((lado_maximo, lado_maximo))
            qualidade = 82
            while qualidade >= 38:
                tentativa.save(destino, format="JPEG", quality=qualidade, optimize=True)
                if destino.stat().st_size <= limite_bytes:
                    if log:
                        log(
                            f"Imagem preparada para GEMUL: {caminho.stat().st_size // 1024} KB "
                            f"-> {destino.stat().st_size // 1024} KB."
                        )
                    return destino
                qualidade -= 8
            lado_maximo -= 200

    raise RuntimeError("Nao foi possivel reduzir a imagem para menos de 2MB.")


def montar_conteudo_documento(caminho: Path, log=None) -> List[Dict[str, str]]:
    extensao = caminho.suffix.lower()
    conteudo = [{"type": "input_text", "text": f"Arquivo analisado: {caminho.name}"}]
    if extensao in IMAGE_EXTENSIONS:
        b64 = base64.b64encode(preparar_imagem_para_api(caminho, log)).decode("ascii")
        conteudo.append({"type": "input_image", "image_url": f"data:image/jpeg;base64,{b64}"})
        return conteudo
    if extensao == ".txt":
        conteudo.append({"type": "input_text", "text": caminho.read_text(encoding="utf-8", errors="ignore")[:200000]})
        return conteudo
    if extensao == ".pdf":
        conteudo.append(
            {
                "type": "input_text",
                "text": (
                    "O arquivo selecionado e PDF. Se ele for um PDF escaneado, converta para imagem/JPG "
                    "ou escaneie direto pelo botao do programa para leitura visual pela IA."
                ),
            }
        )
        return conteudo
    raise ValueError("Formato nao suportado. Use imagem, TXT ou PDF.")


def extrair_atestado_com_ia(caminho: Path, alunos: List[Aluno], model: str, log=None) -> Dict[str, Any]:
    api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY nao configurada no ambiente.")

    data_hoje = datetime.now().strftime("%Y-%m-%d")
    schema = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "matricula": {"type": "string"},
            "turma": {"type": "string"},
            "nome_aluno": {"type": "string"},
            "titulo_documento": {"type": "string"},
            "tipo_justificativa": {"type": "string"},
            "data_inicio": {"type": "string", "description": "YYYY-MM-DD ou vazio"},
            "quantidade_dias": {"type": "integer"},
            "data_fim": {"type": "string", "description": "YYYY-MM-DD ou vazio"},
            "confianca": {"type": "number"},
            "texto_encontrado": {"type": "string"},
            "observacoes": {"type": "string"},
        },
        "required": [
            "matricula",
            "turma",
            "nome_aluno",
            "titulo_documento",
            "tipo_justificativa",
            "data_inicio",
            "quantidade_dias",
            "data_fim",
            "confianca",
            "texto_encontrado",
            "observacoes",
        ],
    }

    conteudo_usuario = [
        {
            "type": "input_text",
            "text": (
                "Leia o atestado/documento e identifique o aluno e o periodo de afastamento. "
                "Leia inclusive texto manuscrito. Retorne o nome e a matricula exatamente como conseguir ler no documento. "
                "Se nao encontrar nome, deixe nome_aluno vazio. Se nao encontrar matricula, deixe matricula vazia. "
                "Identifique o titulo do documento, por exemplo 'Atestado medico' ou 'Atestado de comparecimento'. "
                "Para tipo_justificativa, use 'Atestado medico' somente quando for atestado medico; para outros documentos use 'Outros'. "
                "Se o documento disser 'datado de 12/06/2026 e 3 dias', entenda como 12/06/2026 a 14/06/2026. "
                "Se houver apenas data e nao houver quantidade de dias, use 1 dia. "
                "Datas devem sair em YYYY-MM-DD. Hoje e "
                f"{data_hoje}."
            ),
        }
    ]
    conteudo_usuario.extend(montar_conteudo_documento(caminho, log))

    payload = {
        "model": model or OPENAI_MODEL,
        "reasoning": {"effort": "low"},
        "input": [
            {
                "role": "system",
                "content": [
                    {
                        "type": "input_text",
                        "text": (
                            "Voce extrai dados de atestados escolares. "
                            "Retorne apenas JSON valido no schema solicitado."
                        ),
                    }
                ],
            },
            {"role": "user", "content": conteudo_usuario},
        ],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "extracao_atestado_escolar",
                "strict": True,
                "schema": schema,
            }
        },
    }
    resposta = enviar_requisicao_openai(payload, api_key, log)
    resultado = parse_json_seguro(extrair_texto_resposta_openai(resposta))
    if resultado.get("data_inicio") and not resultado.get("data_fim"):
        resultado["data_fim"] = somar_dias(resultado["data_inicio"], int(resultado.get("quantidade_dias") or 1))
    return validar_aluno_extraido(resultado, alunos)


def validar_aluno_extraido(resultado: Dict[str, Any], alunos: List[Aluno]) -> Dict[str, Any]:
    matricula = re.sub(r"\D", "", str(resultado.get("matricula") or ""))
    aluno = encontrar_aluno_por_matricula(alunos, matricula)
    if aluno is None:
        alvo = normalizar_texto(resultado.get("nome_aluno"))
        if alvo:
            melhores = sorted(
                alunos,
                key=lambda item: similaridade(alvo, normalizar_texto(item.nome)),
                reverse=True,
            )
            if melhores and similaridade(alvo, normalizar_texto(melhores[0].nome)) >= 0.35:
                aluno = melhores[0]
    if aluno:
        aplicar_aluno_no_resultado(resultado, aluno)
    return resultado


def encontrar_aluno_por_matricula(alunos: List[Aluno], matricula: Any) -> Optional[Aluno]:
    matricula_limpa = re.sub(r"\D", "", str(matricula or ""))
    if not matricula_limpa:
        return None
    por_matricula = {aluno.matricula: aluno for aluno in alunos}
    return por_matricula.get(matricula_limpa)


def aplicar_aluno_no_resultado(resultado: Dict[str, Any], aluno: Aluno) -> None:
    resultado["matricula"] = aluno.matricula
    resultado["turma"] = aluno.turma
    resultado["nome_aluno"] = aluno.nome
    resultado["nome_arquivo"] = aluno.nome_arquivo


def similaridade(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    set_a = set(a.split())
    set_b = set(b.split())
    intersecao = len(set_a & set_b)
    uniao = len(set_a | set_b) or 1
    return intersecao / uniao


class GemulBrowser:
    def __init__(self, log):
        self.log = log
        self.play = None
        self.browser = None
        self.context = None
        self.page = None
        self.fila = queue.Queue()
        self.worker = None
        self.worker_id = None

    def _garantir_worker(self):
        if self.worker and self.worker.is_alive():
            return
        self.worker = threading.Thread(target=self._loop_worker, daemon=True)
        self.worker.start()

    def _loop_worker(self):
        self.worker_id = threading.get_ident()
        while True:
            func, args, kwargs, evento, resposta = self.fila.get()
            try:
                if func is None:
                    resposta["resultado"] = None
                    return
                resposta["resultado"] = func(*args, **kwargs)
            except BaseException as exc:
                resposta["erro"] = exc
            finally:
                evento.set()

    def _executar_no_worker(self, func, *args, **kwargs):
        if threading.get_ident() == self.worker_id:
            return func(*args, **kwargs)
        self._garantir_worker()
        evento = threading.Event()
        resposta = {}
        self.fila.put((func, args, kwargs, evento, resposta))
        evento.wait()
        if "erro" in resposta:
            raise resposta["erro"]
        return resposta.get("resultado")

    def maximizar_janela(self):
        if not self.page or self.page.is_closed() or not self.context:
            return
        try:
            sessao = self.context.new_cdp_session(self.page)
            janela = sessao.send("Browser.getWindowForTarget")
            sessao.send("Browser.setWindowBounds", {"windowId": janela["windowId"], "bounds": {"windowState": "maximized"}})
        except Exception as exc:
            self.log(f"Nao foi possivel maximizar automaticamente o navegador: {exc}")

    def _selecionar_por_texto(self, seletor: str, texto: str, obrigatorio: bool = True) -> bool:
        if not self.page:
            raise RuntimeError("Navegador nao aberto.")
        valor = self.page.locator(seletor).evaluate(
            """(select, textoBusca) => {
                const normalizar = (valor) => (valor || "")
                    .normalize("NFD").replace(/[\\u0300-\\u036f]/g, "")
                    .replace(/[^A-Za-z0-9]+/g, " ")
                    .replace(/\\s+/g, " ").trim().toUpperCase();
                const textoNorm = normalizar(textoBusca);
                const opcoes = Array.from(select.options || []);
                const exata = opcoes.find((opcao) => normalizar(opcao.textContent) === textoNorm);
                const parcial = opcoes.find((opcao) => normalizar(opcao.textContent).includes(textoNorm));
                const opcao = exata || parcial;
                if (!opcao) return null;
                select.value = opcao.value;
                select.dispatchEvent(new Event("change", {bubbles: true}));
                return opcao.value;
            }""",
            texto,
        )
        if valor:
            return True
        if obrigatorio:
            raise RuntimeError(f"Nao encontrei a opcao '{texto}' em {seletor}.")
        return False

    def _selecionar_bimestre(self, data_inicio: str):
        data = parse_data_usuario(data_inicio)
        if not data:
            self.log("Tabela bimestral nao preenchida: data inicial ausente ou invalida.")
            return
        indice = 1 if data.month <= 3 else 2 if data.month <= 6 else 3 if data.month <= 9 else 4
        self._selecionar_por_texto("#cmbBimestre", f"{indice} BIMESTRE", obrigatorio=False)

    def _preencher_data(self, seletor: str, valor: str, rotulo: str):
        campo = self.page.locator(seletor)
        campo.fill(valor)
        campo.evaluate(
            """(el) => {
                el.dispatchEvent(new Event("input", {bubbles: true}));
                el.dispatchEvent(new Event("change", {bubbles: true}));
                el.dispatchEvent(new Event("blur", {bubbles: true}));
            }"""
        )
        self.page.wait_for_timeout(900)
        try:
            self.page.wait_for_load_state("networkidle", timeout=3000)
        except PlaywrightTimeoutError:
            pass
        valor_atual = campo.input_value(timeout=5000).strip()
        if valor_atual != valor:
            campo.fill(valor)
            self.page.wait_for_timeout(400)
            valor_atual = campo.input_value(timeout=5000).strip()
        if valor_atual != valor:
            raise RuntimeError(f"{rotulo} nao permaneceu preenchida no GEMUL. Valor esperado: {valor}.")

    def preencher_justificativa(self, resultado: Dict[str, Any], arquivo: Optional[Path] = None):
        return self._executar_no_worker(self._preencher_justificativa, resultado, arquivo)

    def _preencher_justificativa(self, resultado: Dict[str, Any], arquivo: Optional[Path] = None):
        if not self.page or self.page.is_closed():
            raise RuntimeError("Abra o GEMUL antes de lancar a justificativa.")
        self.page.goto(GEMUL_ATESTADO_URL, wait_until="domcontentloaded", timeout=60000)
        self.maximizar_janela()
        self.page.locator("#frm").wait_for(state="visible", timeout=15000)

        data_inicio = formatar_data_iso_para_br(str(resultado.get("data_inicio") or ""))
        data_fim = formatar_data_iso_para_br(str(resultado.get("data_fim") or ""))
        turma = str(resultado.get("turma") or "").strip()
        nome_aluno = str(resultado.get("nome_aluno") or "").strip()
        matricula = re.sub(r"\D", "", str(resultado.get("matricula") or ""))
        titulo = str(resultado.get("titulo_documento") or "").strip() or "Atestado medico"
        tipo = str(resultado.get("tipo_justificativa") or "").strip()
        tipo_opcao = "ATESTADO MEDICO" if "ATESTADO" in normalizar_texto(tipo or titulo) and "MEDICO" in normalizar_texto(tipo or titulo) else "OUTROS"

        self._selecionar_bimestre(data_inicio)
        try:
            self.page.wait_for_load_state("domcontentloaded", timeout=10000)
        except PlaywrightTimeoutError:
            pass

        self._selecionar_por_texto("#cmbTipoJustificativa", tipo_opcao)
        try:
            self.page.wait_for_load_state("domcontentloaded", timeout=10000)
        except PlaywrightTimeoutError:
            pass

        if not turma:
            raise RuntimeError("Turma nao identificada; digite a matricula ou revise a leitura da IA antes de lancar.")
        self._selecionar_por_texto("#cmbTurma", turma)
        try:
            self.page.wait_for_load_state("domcontentloaded", timeout=15000)
        except PlaywrightTimeoutError:
            pass

        if nome_aluno:
            selecionou = self._selecionar_por_texto("#cmbAluno", nome_aluno, obrigatorio=False)
        else:
            selecionou = False
        if not selecionou and matricula:
            selecionou = self._selecionar_por_texto("#cmbAluno", matricula, obrigatorio=False)
        if not selecionou:
            raise RuntimeError("Aluno nao encontrado no combo do GEMUL apos selecionar a turma.")

        if data_inicio:
            self._preencher_data("#txtDataInicial", data_inicio, "Data Inicial")
        if data_fim:
            self._preencher_data("#txtDataFinal", data_fim, "Data Final")
        self.page.locator("#txtJustificativa").fill(titulo)

        if arquivo and arquivo.exists() and arquivo.suffix.lower() in IMAGE_EXTENSIONS:
            self.page.locator("#editor").evaluate("(el) => { el.innerHTML = ''; }")
            self.page.locator("#txtTexto").evaluate("(el) => { el.value = ''; }")
            arquivo_upload = preparar_imagem_para_gemul(arquivo, self.log)
            tamanho_upload = arquivo_upload.stat().st_size
            if tamanho_upload > 2 * 1024 * 1024:
                raise RuntimeError(f"Imagem ainda excede 2MB apos compactacao: {tamanho_upload // 1024} KB.")
            upload = self.page.locator('input[type="file"][data-edit="insertImage"]')
            try:
                upload.set_input_files(str(arquivo_upload))
                self.page.wait_for_timeout(2000)
                self.page.locator("#editor img").first.wait_for(state="attached", timeout=5000)
                self.log(f"Imagem anexada no editor do GEMUL: {arquivo_upload.name} ({tamanho_upload // 1024} KB).")
            except Exception as exc:
                self.page.locator("#editor").evaluate("(el) => { el.innerHTML = ''; }")
                self.page.locator("#txtTexto").evaluate("(el) => { el.value = ''; }")
                raise RuntimeError(f"Falha ao anexar imagem no GEMUL. O editor foi limpo para nao bloquear o Incluir. Detalhe: {exc}") from exc
        elif arquivo and arquivo.exists():
            self.log("Documento nao anexado no GEMUL: a pagina aceita JPG/PNG no campo de imagem.")

        self.log("Formulario do GEMUL preenchido. Revise e clique em Incluir.")

    def abrir(self, usuario: str, senha: str):
        return self._executar_no_worker(self._abrir, usuario, senha)

    def _abrir(self, usuario: str, senha: str):
        if self.page and not self.page.is_closed():
            self.page.goto(GEMUL_ATESTADO_URL, wait_until="domcontentloaded", timeout=60000)
            self.maximizar_janela()
            return
        self.play = sync_playwright().start()
        self.browser = self.play.chromium.launch(headless=False, args=["--start-maximized"])
        self.context = self.browser.new_context(accept_downloads=True, viewport=None)
        self.page = self.context.new_page()
        self.maximizar_janela()
        self.page.goto(GEMUL_LOGIN_URL, wait_until="domcontentloaded", timeout=60000)
        self.page.locator('xpath=//*[@id="txtLogin"]').wait_for(state="visible", timeout=15000)
        if usuario:
            self.page.locator('xpath=//*[@id="txtLogin"]').fill(usuario)
        if senha:
            self.page.locator('xpath=//*[@id="txtSenha"]').fill(senha)
        if usuario and senha:
            self.page.locator('xpath=//*[@id="frm"]/button').click()
            try:
                self.page.wait_for_load_state("networkidle", timeout=15000)
            except PlaywrightTimeoutError:
                pass
            self.page.goto(GEMUL_ATESTADO_URL, wait_until="domcontentloaded", timeout=60000)
        self.maximizar_janela()
        self.log("Navegador aberto na tela de justificativa de faltas.")

    def fechar(self):
        try:
            self._executar_no_worker(self._fechar)
        finally:
            if self.worker and self.worker.is_alive():
                evento = threading.Event()
                self.fila.put((None, (), {}, evento, {}))
                evento.wait(timeout=3)
            self.worker = None
            self.worker_id = None

    def _fechar(self):
        try:
            if self.context:
                self.context.close()
            if self.browser:
                self.browser.close()
            if self.play:
                self.play.stop()
        finally:
            self.play = None
            self.browser = None
            self.context = None
            self.page = None


class LeitorAtestadoApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"Leitor de Atestado v{APP_VERSION}")
        self.root.geometry("980x680")
        self.alunos = carregar_alunos()
        self.browser = GemulBrowser(self.log)
        self.arquivo_atual: Optional[Path] = None
        self.resultado_atual: Optional[Dict[str, Any]] = None
        self.escaneando = False

        self.pasta_saida_var = tk.StringVar(value=str(APP_DIR / "atestados_salvos"))
        self.usuario_var = tk.StringVar(value=os.getenv("GEMUL_USUARIO", ""))
        self.senha_var = tk.StringVar(value=os.getenv("GEMUL_SENHA", ""))
        self.modelo_var = tk.StringVar(value=OPENAI_MODEL)
        self.arquivo_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value=f"{len(self.alunos)} alunos carregados.")
        self._montar_interface()

    def _montar_interface(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(5, weight=1)

        frame_login = ttk.LabelFrame(self.root, text="GEMUL")
        frame_login.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        frame_login.columnconfigure(1, weight=1)
        frame_login.columnconfigure(3, weight=1)
        ttk.Label(frame_login, text="Usuario").grid(row=0, column=0, padx=8, pady=8, sticky="w")
        ttk.Entry(frame_login, textvariable=self.usuario_var).grid(row=0, column=1, padx=8, pady=8, sticky="ew")
        ttk.Label(frame_login, text="Senha").grid(row=0, column=2, padx=8, pady=8, sticky="w")
        ttk.Entry(frame_login, textvariable=self.senha_var, show="*").grid(row=0, column=3, padx=8, pady=8, sticky="ew")
        ttk.Button(frame_login, text="Abrir e login", command=self.abrir_gemul_thread).grid(row=0, column=4, padx=8, pady=8)

        frame_saida = ttk.LabelFrame(self.root, text="Salvar")
        frame_saida.grid(row=1, column=0, sticky="ew", padx=10, pady=6)
        frame_saida.columnconfigure(1, weight=1)
        ttk.Label(frame_saida, text="Pasta").grid(row=0, column=0, padx=8, pady=8, sticky="w")
        ttk.Entry(frame_saida, textvariable=self.pasta_saida_var).grid(row=0, column=1, padx=8, pady=8, sticky="ew")
        ttk.Button(frame_saida, text="Selecionar", command=self.selecionar_pasta).grid(row=0, column=2, padx=8, pady=8)

        frame_doc = ttk.LabelFrame(self.root, text="Documento")
        frame_doc.grid(row=2, column=0, sticky="ew", padx=10, pady=6)
        frame_doc.columnconfigure(1, weight=1)
        ttk.Button(frame_doc, text="Escanear Kyocera/WIA", command=self.escanear_thread).grid(row=0, column=0, padx=8, pady=8)
        ttk.Entry(frame_doc, textvariable=self.arquivo_var).grid(row=0, column=1, padx=8, pady=8, sticky="ew")
        ttk.Button(frame_doc, text="Selecionar arquivo", command=self.selecionar_arquivo).grid(row=0, column=2, padx=8, pady=8)

        frame_ia = ttk.Frame(self.root)
        frame_ia.grid(row=3, column=0, sticky="ew", padx=10, pady=6)
        ttk.Label(frame_ia, text="Modelo").pack(side="left")
        ttk.Entry(frame_ia, textvariable=self.modelo_var, width=18).pack(side="left", padx=(6, 12))
        ttk.Button(frame_ia, text="Ler com IA", command=self.ler_ia_thread).pack(side="left")
        ttk.Button(frame_ia, text="Salvar com nome padrao", command=self.salvar_documento).pack(side="left", padx=8)
        ttk.Button(frame_ia, text="Lancar no GEMUL", command=self.lancar_gemul_thread).pack(side="left")

        frame_resultado = ttk.LabelFrame(self.root, text="Resultado")
        frame_resultado.grid(row=4, column=0, sticky="ew", padx=10, pady=6)
        frame_resultado.columnconfigure(1, weight=1)
        self.campos = {}
        nomes = [
            ("matricula", "Matricula"),
            ("turma", "Turma"),
            ("nome_aluno", "Aluno"),
            ("titulo_documento", "Ocorrencia"),
            ("tipo_justificativa", "Tipo"),
            ("data_inicio", "Inicio"),
            ("data_fim", "Fim"),
            ("quantidade_dias", "Dias"),
            ("confianca", "Confianca"),
        ]
        for i, (chave, rotulo) in enumerate(nomes):
            ttk.Label(frame_resultado, text=rotulo).grid(row=i // 2, column=(i % 2) * 2, padx=8, pady=4, sticky="w")
            var = tk.StringVar(value="")
            self.campos[chave] = var
            entrada = ttk.Entry(frame_resultado, textvariable=var)
            entrada.grid(row=i // 2, column=(i % 2) * 2 + 1, padx=8, pady=4, sticky="ew")
            if chave == "matricula":
                entrada.bind("<Return>", self.aplicar_matricula_manual)
                entrada.bind("<FocusOut>", self.aplicar_matricula_manual)

        frame_log = ttk.LabelFrame(self.root, text="Log")
        frame_log.grid(row=5, column=0, sticky="nsew", padx=10, pady=(6, 10))
        frame_log.rowconfigure(0, weight=1)
        frame_log.columnconfigure(0, weight=1)
        self.log_text = tk.Text(frame_log, height=12, wrap="word")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        barra = ttk.Scrollbar(frame_log, command=self.log_text.yview)
        barra.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=barra.set)
        ttk.Label(self.root, textvariable=self.status_var).grid(row=6, column=0, sticky="ew", padx=10, pady=(0, 8))

    def log(self, mensagem: str):
        registrar_log_arquivo(mensagem)

        def atualizar():
            carimbo = datetime.now().strftime("%H:%M:%S")
            self.log_text.insert("end", f"[{carimbo}] {mensagem}\n")
            self.log_text.see("end")
            self.status_var.set(mensagem)

        self.root.after(0, atualizar)

    def executar_thread(self, func):
        threading.Thread(target=func, daemon=True).start()

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(initialdir=self.pasta_saida_var.get() or str(APP_DIR))
        if pasta:
            self.pasta_saida_var.set(pasta)

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            initialdir=str(APP_DIR),
            filetypes=[
                ("Documentos", "*.png *.jpg *.jpeg *.webp *.tif *.tiff *.pdf *.txt"),
                ("Todos", "*.*"),
            ],
        )
        if caminho:
            self.definir_arquivo(Path(caminho))

    def definir_arquivo(self, caminho: Path):
        self.arquivo_atual = caminho
        self.resultado_atual = None
        self.root.after(0, lambda: self.arquivo_var.set(str(caminho)))
        self.root.after(0, self.limpar_campos_resultado)
        self.log(f"Documento selecionado: {caminho}")

    def limpar_campos_resultado(self):
        for var in self.campos.values():
            var.set("")

    def obter_documento_atual(self) -> Path:
        if self.escaneando:
            raise RuntimeError("Aguarde a digitalizacao terminar antes de usar a IA.")
        texto_caminho = (self.arquivo_var.get() or "").strip()
        caminho = self.arquivo_atual or (Path(texto_caminho) if texto_caminho else None)
        if caminho is None:
            raise FileNotFoundError("Selecione ou escaneie um documento antes de usar a IA.")
        caminho = Path(caminho)
        if not caminho.exists():
            raise FileNotFoundError(f"Documento nao encontrado: {caminho}")
        if not caminho.is_file():
            raise ValueError(f"O caminho selecionado nao e um arquivo: {caminho}")
        if caminho.suffix.lower() not in DOCUMENT_EXTENSIONS:
            raise ValueError(
                f"Formato nao suportado para leitura por IA: '{caminho.suffix or 'sem extensao'}'. "
                "Use PNG, JPG, JPEG, WEBP, TIF, TIFF, PDF ou TXT."
            )
        self.arquivo_atual = caminho
        return caminho

    def abrir_gemul_thread(self):
        self.executar_thread(self.abrir_gemul)

    def abrir_gemul(self):
        try:
            self.log("Abrindo navegador...")
            self.browser.abrir(self.usuario_var.get().strip(), self.senha_var.get().strip())
        except Exception as exc:
            self.log(f"Erro ao abrir GEMUL: {exc}")
            messagebox.showerror("GEMUL", str(exc))

    def escanear_thread(self):
        self.executar_thread(self.escanear)

    def escanear(self):
        try:
            self.escaneando = True
            pasta = Path(self.pasta_saida_var.get() or APP_DIR)
            pasta.mkdir(parents=True, exist_ok=True)
            destino = pasta / f"scan_atestado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
            self.log("Abrindo interface do scanner WIA...")
            caminho = escanear_wia(destino)
            self.definir_arquivo(caminho)
            self.log("Digitalizacao concluida.")
        except Exception as exc:
            self.log(f"Erro ao escanear: {exc}")
            messagebox.showerror("Scanner", str(exc))
        finally:
            self.escaneando = False

    def ler_ia_thread(self):
        self.executar_thread(self.ler_ia)

    def lancar_gemul_thread(self):
        self.executar_thread(self.lancar_gemul)

    def ler_ia(self):
        try:
            caminho = self.obter_documento_atual()
            tamanho_kb = caminho.stat().st_size // 1024
            modelo = self.modelo_var.get().strip() or OPENAI_MODEL
            self.log(f"Enviando documento para a OpenAI: {caminho.name} ({tamanho_kb} KB), modelo {modelo}.")
            inicio = time.monotonic()
            resultado = extrair_atestado_com_ia(caminho, self.alunos, modelo, self.log)
            self.resultado_atual = resultado
            self.root.after(0, lambda: self.exibir_resultado(resultado))
            self.log(f"Leitura por IA concluida em {time.monotonic() - inicio:.1f}s.")
        except Exception as exc:
            self.log(f"Erro na leitura por IA: {exc}")
            messagebox.showerror("Leitura por IA", str(exc))

    def exibir_resultado(self, resultado: Dict[str, Any]):
        for chave, var in self.campos.items():
            valor = resultado.get(chave, "")
            if chave in {"data_inicio", "data_fim"}:
                valor = formatar_data_iso_para_br(str(valor))
            var.set(str(valor))
        self.log(f"Aluno identificado: {resultado.get('nome_aluno', '')}")

    def sincronizar_resultado_dos_campos(self) -> Dict[str, Any]:
        resultado = dict(self.resultado_atual or {})
        for chave, var in self.campos.items():
            valor = var.get().strip()
            if valor:
                resultado[chave] = valor
        aluno = encontrar_aluno_por_matricula(self.alunos, resultado.get("matricula"))
        if aluno:
            aplicar_aluno_no_resultado(resultado, aluno)
        self.resultado_atual = resultado
        return resultado

    def aplicar_matricula_manual(self, _evento=None) -> bool:
        matricula = self.campos.get("matricula", tk.StringVar()).get()
        aluno = encontrar_aluno_por_matricula(self.alunos, matricula)
        if not aluno:
            return False

        resultado = dict(self.resultado_atual or {})
        aplicar_aluno_no_resultado(resultado, aluno)
        self.resultado_atual = resultado
        self.campos["matricula"].set(aluno.matricula)
        self.campos["turma"].set(aluno.turma)
        self.campos["nome_aluno"].set(aluno.nome)
        self.log(f"Aluno definido pela matricula: {aluno.nome}")
        return True

    def lancar_gemul(self):
        try:
            caminho = self.obter_documento_atual()
            self.aplicar_matricula_manual()
            resultado = self.sincronizar_resultado_dos_campos()
            if not resultado.get("matricula") or not resultado.get("nome_aluno"):
                raise RuntimeError("Identifique o aluno antes de lancar no GEMUL.")
            if not resultado.get("data_inicio") or not resultado.get("data_fim"):
                raise RuntimeError("Preencha data inicial e data final antes de lancar no GEMUL.")
            self.log("Preenchendo formulario do GEMUL...")
            self.browser.preencher_justificativa(resultado, caminho)
        except Exception as exc:
            self.log(f"Erro ao lancar no GEMUL: {exc}")
            messagebox.showerror("GEMUL", str(exc))

    def salvar_documento(self):
        try:
            self.arquivo_atual = self.obter_documento_atual()
            self.aplicar_matricula_manual()
            resultado = self.resultado_atual or {}
            nome_base = resultado.get("nome_arquivo") or "ATESTADO_SEM_ALUNO_IDENTIFICADO"
            pasta = Path(self.pasta_saida_var.get() or APP_DIR)
            pasta.mkdir(parents=True, exist_ok=True)
            destino = pasta / f"{limpar_nome_arquivo(nome_base)}{self.arquivo_atual.suffix.lower()}"
            contador = 2
            while destino.exists():
                destino = pasta / f"{limpar_nome_arquivo(nome_base)} ({contador}){self.arquivo_atual.suffix.lower()}"
                contador += 1
            origem_salvar = self.arquivo_atual
            if self.arquivo_atual.suffix.lower() in IMAGE_EXTENSIONS:
                origem_salvar = preparar_imagem_para_gemul(self.arquivo_atual, self.log)
                destino = destino.with_suffix(".jpg")
            shutil.copy2(origem_salvar, destino)
            self.log(f"Arquivo salvo: {destino}")
            messagebox.showinfo("Salvar", f"Arquivo salvo:\n{destino}")
        except Exception as exc:
            self.log(f"Erro ao salvar: {exc}")
            messagebox.showerror("Salvar", str(exc))


def escanear_wia(destino: Path) -> Path:
    pythoncom = None
    try:
        import win32com.client
        import pythoncom as py_com
    except Exception as exc:
        raise RuntimeError("pywin32/win32com nao esta disponivel para acessar o scanner WIA.") from exc

    pythoncom = py_com
    pythoncom.CoInitialize()
    try:
        dialog = win32com.client.Dispatch("WIA.CommonDialog")
        imagem = dialog.ShowAcquireImage(1, 0, 0, WIA_FORMAT_JPEG, False, True, False)
        if imagem is None:
            raise RuntimeError("Digitalizacao cancelada ou nenhum scanner retornou imagem.")
        if destino.exists():
            destino.unlink()
        imagem.SaveFile(str(destino))
        return destino
    finally:
        pythoncom.CoUninitialize()


def main():
    root = tk.Tk()
    try:
        app = LeitorAtestadoApp(root)
    except Exception as exc:
        messagebox.showerror("Inicializacao", str(exc))
        raise

    def ao_fechar():
        try:
            app.browser.fechar()
        finally:
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", ao_fechar)
    root.mainloop()


if __name__ == "__main__":
    main()
