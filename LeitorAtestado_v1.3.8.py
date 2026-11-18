import base64
import csv
import io
import json
import math
import os
import queue
import re
import shutil
import socket
import ssl
import tempfile
import threading
import time
import traceback
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
APP_VERSION = "1.3.8"
GEMUL_LOGIN_URL = "https://gemul-aparecida.com.br/login.asp"
GEMUL_ATESTADO_URL = "https://gemul-aparecida.com.br/app/sca_justificativa_faltas.asp"
OPENAI_MODEL = (os.getenv("OPENAI_MODEL") or "gpt-5-mini").strip() or "gpt-5-mini"
OPENAI_TIMEOUT = int(os.getenv("OPENAI_TIMEOUT", "90"))
OPENAI_TENTATIVAS = int(os.getenv("OPENAI_TENTATIVAS", "2"))
OPENAI_IMAGE_MAX_SIDE = int(os.getenv("OPENAI_IMAGE_MAX_SIDE", "1600"))
OPENAI_IMAGE_JPEG_QUALITY = int(os.getenv("OPENAI_IMAGE_JPEG_QUALITY", "78"))
OPENAI_PDF_MAX_PAGES = int(os.getenv("OPENAI_PDF_MAX_PAGES", "3"))
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff"}
DOCUMENT_EXTENSIONS = IMAGE_EXTENSIONS | {".txt", ".pdf"}
WIA_FORMAT_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
WIA_INTENT_GRAYSCALE = 2
WIA_BIAS_MINIMIZE_SIZE = 131072
LOG_FILE = APP_DIR / "leitor_atestado.log"
CONFIG_FILE = APP_DIR / "leitor_atestado_config.json"


class DigitalizacaoEmAndamentoError(RuntimeError):
    pass


class ScannerWiaError(RuntimeError):
    pass


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


def limpar_texto_gemul(valor: Any) -> str:
    texto = str(valor or "")
    texto = re.sub(r"[<>]", " ", texto)
    texto = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


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


def resumir_excecao(exc: BaseException, limite: int = 500) -> str:
    texto = str(exc).strip() or exc.__class__.__name__
    linhas = [linha.strip() for linha in texto.splitlines() if linha.strip()]
    resumo = linhas[0] if linhas else exc.__class__.__name__
    if len(resumo) > limite:
        resumo = resumo[: limite - 3].rstrip() + "..."
    return resumo


def registrar_excecao_log(contexto: str, exc: BaseException) -> None:
    registrar_log_arquivo(f"{contexto}: {resumir_excecao(exc)}")
    detalhes = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)).rstrip()
    if detalhes:
        registrar_log_arquivo(f"{contexto} - traceback completo:\n{detalhes}")


def ler_ultimas_linhas_log(quantidade: int = 80) -> List[str]:
    try:
        if not LOG_FILE.exists():
            return []
        texto = LOG_FILE.read_text(encoding="utf-8", errors="replace")
        return texto.replace("\x00", "").splitlines()[-quantidade:]
    except Exception:
        return []


def carregar_configuracao() -> Dict[str, Any]:
    try:
        if not CONFIG_FILE.exists():
            return {}
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def salvar_configuracao(dados: Dict[str, Any]) -> None:
    try:
        CONFIG_FILE.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        registrar_log_arquivo(f"Nao foi possivel salvar configuracao local: {resumir_excecao(exc)}")


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


def otimizar_imagem_bytes_para_api(dados: bytes, origem: str, log=None) -> bytes:
    try:
        from PIL import Image, ImageOps
    except Exception:
        if log:
            log("Pillow nao disponivel; enviando imagem original para a API.")
        return dados

    try:
        with Image.open(io.BytesIO(dados)) as imagem:
            imagem = ImageOps.exif_transpose(imagem)
            imagem.thumbnail((OPENAI_IMAGE_MAX_SIDE, OPENAI_IMAGE_MAX_SIDE))
            if imagem.mode not in {"RGB", "L"}:
                imagem = imagem.convert("RGB")
            buffer = io.BytesIO()
            imagem.save(buffer, format="JPEG", quality=OPENAI_IMAGE_JPEG_QUALITY, optimize=True)
            otimizada = buffer.getvalue()
    except Exception as exc:
        if log:
            log(f"Nao foi possivel otimizar a imagem; enviando original. Detalhe: {exc}")
        return dados

    if len(otimizada) >= len(dados):
        return dados
    if log:
        log(f"{origem} otimizada para API: {len(dados) // 1024} KB -> {len(otimizada) // 1024} KB.")
    return otimizada


def preparar_imagem_para_api(caminho: Path, log=None) -> bytes:
    return otimizar_imagem_bytes_para_api(caminho.read_bytes(), "Imagem", log)


def conteudo_imagem_jpeg(dados: bytes) -> Dict[str, str]:
    b64 = base64.b64encode(dados).decode("ascii")
    return {"type": "input_image", "image_url": f"data:image/jpeg;base64,{b64}"}


def converter_pdf_para_imagens_api(caminho: Path, log=None) -> List[bytes]:
    try:
        tamanho = caminho.stat().st_size
    except OSError as exc:
        raise RuntimeError(f"Nao foi possivel acessar o PDF: {caminho}") from exc
    if tamanho <= 0:
        raise RuntimeError(
            "O PDF selecionado esta vazio (0 KB). Digitalize novamente ou selecione outro arquivo; "
            "a conversao so funciona com um PDF que tenha conteudo."
        )

    limite_paginas = max(1, OPENAI_PDF_MAX_PAGES)
    imagens: List[bytes] = []
    try:
        import fitz

        with fitz.open(caminho) as documento:
            total = len(documento)
            if total == 0:
                raise RuntimeError("PDF sem paginas.")
            for indice in range(min(total, limite_paginas)):
                pagina = documento.load_page(indice)
                pixmap = pagina.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                dados_png = pixmap.tobytes("png")
                imagens.append(otimizar_imagem_bytes_para_api(dados_png, f"PDF pagina {indice + 1}", log))
            if log:
                log(f"PDF convertido para {len(imagens)} imagem(ns) para leitura visual pela IA.")
            return imagens
    except Exception as exc_fitz:
        if log:
            log(f"Conversao de PDF por PyMuPDF falhou; tentando pdf2image. Detalhe: {resumir_excecao(exc_fitz)}")

    try:
        from pdf2image import convert_from_path

        paginas = convert_from_path(str(caminho), dpi=180, first_page=1, last_page=limite_paginas)
        for indice, pagina in enumerate(paginas):
            buffer = io.BytesIO()
            if pagina.mode not in {"RGB", "L"}:
                pagina = pagina.convert("RGB")
            pagina.save(buffer, format="JPEG", quality=OPENAI_IMAGE_JPEG_QUALITY, optimize=True)
            imagens.append(otimizar_imagem_bytes_para_api(buffer.getvalue(), f"PDF pagina {indice + 1}", log))
        if imagens:
            if log:
                log(f"PDF convertido para {len(imagens)} imagem(ns) para leitura visual pela IA.")
            return imagens
    except Exception as exc_pdf2image:
        if log:
            log(f"Conversao de PDF por pdf2image falhou. Detalhe: {resumir_excecao(exc_pdf2image)}")

    raise RuntimeError(
        "Nao foi possivel converter o PDF para imagem. Instale/ative PyMuPDF ou pdf2image/Poppler, "
        "ou selecione o arquivo em JPG/PNG."
    )


def preparar_pdf_para_gemul(caminho: Path, log=None) -> Path:
    destino_dir = APP_DIR / "atestados_salvos"
    destino_dir.mkdir(parents=True, exist_ok=True)
    destino = destino_dir / f"{limpar_nome_arquivo(caminho.stem)}_pagina1_gemul.jpg"
    imagens = converter_pdf_para_imagens_api(caminho, log)
    if not imagens:
        raise RuntimeError("PDF sem pagina convertida para anexar no GEMUL.")
    destino.write_bytes(imagens[0])
    destino = preparar_imagem_para_gemul(destino, log)
    if log:
        log(f"Primeira pagina do PDF preparada para GEMUL: {destino}")
    return destino


def preparar_documento_para_gemul(caminho: Path, log=None) -> Path:
    if caminho.suffix.lower() in IMAGE_EXTENSIONS:
        return preparar_imagem_para_gemul(caminho, log)
    if caminho.suffix.lower() == ".pdf":
        return preparar_pdf_para_gemul(caminho, log)
    raise RuntimeError("O GEMUL aceita apenas JPG/PNG no campo de imagem; para PDF, use um PDF escaneado/visual.")


def preparar_imagem_para_gemul(caminho: Path, log=None) -> Path:
    if caminho.suffix.lower() not in IMAGE_EXTENSIONS:
        return caminho

    try:
        from PIL import Image, ImageOps
    except Exception:
        if log:
            log("Pillow nao disponivel; imagem original pode exceder o limite de 2MB do GEMUL.")
        return caminho

    limite_bytes = 2 * 1024 * 1024
    destino = caminho.with_name(f"{limpar_nome_arquivo(caminho.stem)}_gemul_upload.jpg")
    with Image.open(caminho) as imagem:
        imagem = ImageOps.exif_transpose(imagem)
        if imagem.mode != "RGB":
            imagem = imagem.convert("RGB")

        lado_maximo = max(imagem.size)
        while lado_maximo >= 700:
            tentativa = imagem.copy()
            tentativa.thumbnail((lado_maximo, lado_maximo))
            qualidade = 82
            while qualidade >= 38:
                tentativa.save(destino, format="JPEG", quality=qualidade, optimize=True, progressive=False)
                if destino.stat().st_size <= limite_bytes:
                    if log:
                        log(
                            f"Imagem preparada para GEMUL: {caminho.stat().st_size // 1024} KB "
                            f"-> {destino.stat().st_size // 1024} KB ({destino.name})."
                        )
                    return destino
                qualidade -= 8
            lado_maximo -= 250

    raise RuntimeError("Nao foi possivel reduzir a imagem para menos de 2MB.")


def montar_conteudo_documento(caminho: Path, log=None) -> List[Dict[str, str]]:
    extensao = caminho.suffix.lower()
    conteudo = [{"type": "input_text", "text": f"Arquivo analisado: {caminho.name}"}]
    if extensao in IMAGE_EXTENSIONS:
        conteudo.append(conteudo_imagem_jpeg(preparar_imagem_para_api(caminho, log)))
        return conteudo
    if extensao == ".txt":
        conteudo.append({"type": "input_text", "text": caminho.read_text(encoding="utf-8", errors="ignore")[:200000]})
        return conteudo
    if extensao == ".pdf":
        for indice, imagem in enumerate(converter_pdf_para_imagens_api(caminho, log), start=1):
            conteudo.append({"type": "input_text", "text": f"Pagina {indice} do PDF renderizada como imagem."})
            conteudo.append(conteudo_imagem_jpeg(imagem))
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
        self.perfil_temporario = None
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
            sessao.send("Browser.setWindowBounds", {"windowId": janela["windowId"], "bounds": {"windowState": "normal"}})
            sessao.send("Browser.setWindowBounds", {"windowId": janela["windowId"], "bounds": {"windowState": "maximized"}})
        except Exception as exc:
            self.log(f"Nao foi possivel maximizar automaticamente o navegador: {exc}")

    def trazer_para_frente(self):
        return self._executar_no_worker(self._trazer_para_frente)

    def page_down(self):
        return self._executar_no_worker(self._page_down)

    def _trazer_para_frente(self) -> None:
        if not self.page or self.page.is_closed():
            return
        try:
            self.page.bring_to_front()
        except Exception as exc:
            self.log(f"Nao foi possivel trazer a aba do GEMUL para frente: {resumir_excecao(exc)}")
        try:
            self.maximizar_janela()
        except Exception:
            pass
        try:
            import ctypes

            user32 = ctypes.windll.user32
            hwnds = []

            def enum_proc(hwnd, _lparam):
                if not user32.IsWindowVisible(hwnd):
                    return True
                tamanho = user32.GetWindowTextLengthW(hwnd)
                if tamanho <= 0:
                    return True
                buffer = ctypes.create_unicode_buffer(tamanho + 1)
                user32.GetWindowTextW(hwnd, buffer, tamanho + 1)
                titulo_norm = normalizar_texto(buffer.value or "")
                if "GEMUL" in titulo_norm or "GESTOR DE MUNICIPIO" in titulo_norm:
                    hwnds.append(hwnd)
                return True

            enum_proc_type = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
            user32.EnumWindows(enum_proc_type(enum_proc), 0)
            if hwnds:
                hwnd = hwnds[0]
                user32.ShowWindow(hwnd, 9)
                user32.SetForegroundWindow(hwnd)
                self.log("Janela do navegador/GEMUL trazida para primeiro plano.")
        except Exception as exc:
            self.log(f"Nao foi possivel alternar foco para o navegador via Windows: {resumir_excecao(exc)}")

    def _page_down(self) -> None:
        if not self.page or self.page.is_closed():
            return
        try:
            self.page.mouse.wheel(0, 850)
            self.page.evaluate("window.scrollBy({top: 850, left: 0, behavior: 'instant'});")
            self.page.wait_for_timeout(350)
            self.log("Rolagem para baixo enviada ao navegador GEMUL.")
        except Exception as exc:
            self.log(f"Nao foi possivel rolar a pagina do GEMUL: {resumir_excecao(exc)}")

    def _ir_para_formulario_justificativa(self):
        for tentativa in range(4):
            try:
                self.page.goto(GEMUL_ATESTADO_URL, wait_until="domcontentloaded", timeout=60000)
                self.page.locator("#frm").wait_for(state="visible", timeout=15000)
                return
            except Exception as exc:
                if "interrupted by another navigation" in str(exc):
                    self.log("GEMUL ainda estava processando a navegacao anterior; aguardando e tentando novamente.")
                    try:
                        self.page.wait_for_load_state("domcontentloaded", timeout=10000)
                    except PlaywrightTimeoutError:
                        pass
                    self.page.wait_for_timeout(2500)
                    continue
                raise

    def _limpar_formulario_visual(self):
        if not self.page or self.page.is_closed():
            return
        self.page.evaluate(
            """() => {
                const editor = document.querySelector("#editor");
                if (editor) editor.innerHTML = "";
                const txtTexto = document.querySelector("#txtTexto");
                if (txtTexto) txtTexto.value = "";
                ["txtCodigo", "txtCodAluno"].forEach((id) => {
                    const el = document.getElementById(id);
                    if (el) el.value = "";
                });
            }"""
        )
        self.log("Formulario do GEMUL limpo para novo lancamento.")

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

    def _desativar_validacao_automatica_campos_estaticos(self):
        if not self.page or self.page.is_closed():
            return
        self.page.evaluate(
            """() => {
                if (window.__leitorAtestadoCamposSemBlur) return;
                window.__leitorAtestadoCamposSemBlur = true;
                ["txtAnoLetivo", "txtDataInicial", "txtDataFinal"].forEach((id) => {
                    const el = document.getElementById(id);
                    if (!el) return;
                    el.removeAttribute("onblur");
                    el.onblur = null;
                    el.addEventListener("blur", (ev) => {
                        ev.stopImmediatePropagation();
                    }, true);
                });
            }"""
        )
        self.log("Validacao automatica ao sair do ano letivo e das datas foi desativada nesta tela do GEMUL.")

    def preencher_justificativa(self, resultado: Dict[str, Any], arquivo: Optional[Path] = None):
        return self._executar_no_worker(self._preencher_justificativa, resultado, arquivo)

    def anexar_imagem(self, arquivo: Path):
        return self._executar_no_worker(self._anexar_imagem, arquivo)

    def _anexar_imagem(self, arquivo: Path):
        if not self.page or self.page.is_closed():
            raise RuntimeError("Abra o GEMUL antes de anexar a imagem.")
        if not arquivo or not arquivo.exists():
            raise FileNotFoundError("Documento atual nao encontrado para anexar no GEMUL.")
        if arquivo.suffix.lower() not in IMAGE_EXTENSIONS:
            raise RuntimeError("O GEMUL aceita apenas JPG/PNG no campo de imagem.")

        self.page.locator("#editor").wait_for(state="attached", timeout=10000)
        self.page.locator("#editor").evaluate("(el) => { el.innerHTML = ''; }")
        self.page.locator("#txtTexto").evaluate("(el) => { el.value = ''; }")

        arquivo_upload = preparar_imagem_para_gemul(arquivo, self.log)
        if arquivo_upload.suffix.lower() not in {".jpg", ".jpeg", ".png"}:
            raise RuntimeError(f"Arquivo preparado para GEMUL nao e JPG/PNG: {arquivo_upload}")
        tamanho_upload = arquivo_upload.stat().st_size
        if tamanho_upload > 2 * 1024 * 1024:
            raise RuntimeError(f"Imagem ainda excede 2MB apos compactacao: {tamanho_upload // 1024} KB.")

        upload_xpath = self.page.locator('xpath=//*[@id="frm"]/div[4]/div[1]/div/input')
        upload_fallback = self.page.locator('input[type="file"][data-edit="insertImage"]')
        try:
            try:
                with self.page.expect_file_chooser(timeout=3000) as chooser_info:
                    self.page.locator("#pictureBtn").click(force=True)
                chooser_info.value.set_files(str(arquivo_upload))
            except Exception:
                try:
                    upload_xpath.set_input_files(str(arquivo_upload), timeout=5000)
                except Exception:
                    upload_fallback.set_input_files(str(arquivo_upload), timeout=5000)
            self.page.evaluate(
                """() => {
                    const input = document.evaluate('//*[@id="frm"]/div[4]/div[1]/div/input', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue
                        || document.querySelector('input[type="file"][data-edit="insertImage"]');
                    if (input) input.dispatchEvent(new Event("change", {bubbles: true}));
                }"""
            )
            self.page.wait_for_timeout(2000)
            self.page.locator("#editor img").first.wait_for(state="attached", timeout=5000)
            self._instalar_liberador_incluir()
            self.log(f"Imagem anexada no editor do GEMUL: {arquivo_upload.name} ({tamanho_upload // 1024} KB).")
        except Exception as exc:
            self.page.locator("#editor").evaluate("(el) => { el.innerHTML = ''; }")
            self.page.locator("#txtTexto").evaluate("(el) => { el.value = ''; }")
            raise RuntimeError(f"Falha ao anexar imagem no GEMUL. O editor foi limpo. Detalhe: {exc}") from exc

    def _instalar_liberador_incluir(self):
        if not self.page or self.page.is_closed():
            return
        self.page.evaluate(
            """() => {
                if (window.__leitorAtestadoLiberadorIncluir) return;
                window.__leitorAtestadoLiberadorIncluir = true;
                const limpar = (valor) => (valor || "").replace(/[<>]/g, " ").replace(/[\\x00-\\x08\\x0b\\x0c\\x0e-\\x1f]/g, " ").replace(/\\s+/g, " ").trim();
                const preparar = () => {
                    ["#txtJustificativa", "#txtDataInicial", "#txtDataFinal"].forEach((seletor) => {
                        const el = document.querySelector(seletor);
                        if (el) el.value = limpar(el.value);
                    });
                    const txtTexto = document.querySelector("#txtTexto");
                    if (txtTexto && /[<>]/.test(txtTexto.value || "")) {
                        txtTexto.value = "";
                    }
                };
                const liberar = () => {
                    const btn = document.querySelector("#btnIncluir");
                    if (!btn) return;
                    const temImagem = !!document.querySelector("#editor img");
                    if (!temImagem) return;
                    btn.disabled = false;
                    btn.removeAttribute("disabled");
                    btn.classList.remove("disabled");
                    btn.style.pointerEvents = "";
                    btn.style.opacity = "";
                    btn.title = "";
                };
                document.addEventListener("click", (ev) => {
                    if (ev.target && ev.target.closest && ev.target.closest("#btnIncluir")) preparar();
                }, true);
                document.addEventListener("change", (ev) => {
                    if (ev.target && ev.target.matches && ev.target.matches('input[type="file"][data-edit="insertImage"]')) {
                        setTimeout(liberar, 500);
                        setTimeout(liberar, 1500);
                    }
                }, true);
                const editor = document.querySelector("#editor");
                if (editor) {
                    new MutationObserver(liberar).observe(editor, {childList: true, subtree: true, attributes: true});
                }
                liberar();
            }"""
        )

    def clicar_incluir(self) -> bool:
        return self._executar_no_worker(self._clicar_incluir)

    def _aceitar_proximo_dialogo_gemul(self) -> None:
        try:
            def aceitar(dialog):
                try:
                    self.log(f"Popup do GEMUL aceito automaticamente: {dialog.message}")
                    dialog.accept()
                except Exception as exc:
                    self.log(f"Falha ao aceitar popup do GEMUL: {resumir_excecao(exc)}")

            self.page.once("dialog", aceitar)
        except Exception as exc:
            self.log(f"Nao foi possivel preparar aceite automatico de popup: {resumir_excecao(exc)}")

    def _aguardar_pos_incluir(self) -> None:
        try:
            self.page.wait_for_load_state("networkidle", timeout=8000)
        except PlaywrightTimeoutError:
            pass
        except Exception:
            pass
        self.page.wait_for_timeout(1200)

    def _clicar_incluir(self) -> bool:
        if not self.page or self.page.is_closed():
            raise RuntimeError("Navegador nao esta aberto.")
        self._instalar_liberador_incluir()
        self._aceitar_proximo_dialogo_gemul()
        seletores = [
            'xpath=//*[@id="btnIncluir"]',
            "#btnIncluir",
            "input[name='cmdIncluir']",
            "input[type='submit']",
            "button[type='submit']",
        ]
        ultimo_erro = ""
        for seletor in seletores:
            try:
                locator = self.page.locator(seletor).first
                if locator.count() == 0:
                    continue
                try:
                    locator.scroll_into_view_if_needed(timeout=3000)
                except Exception:
                    pass
                self.log(f"Tentando clicar no botao Incluir via {seletor}.")
                try:
                    locator.click(timeout=5000)
                except Exception as exc_click:
                    ultimo_erro = resumir_excecao(exc_click)
                    self.log(f"Clique normal no Incluir falhou; tentando clique forcado: {ultimo_erro}")
                    locator.click(timeout=5000, force=True)
                self._aguardar_pos_incluir()
                self.log("Clique no botao Incluir enviado ao GEMUL.")
                return True
            except Exception as exc:
                ultimo_erro = resumir_excecao(exc)
                self.log(f"Tentativa de Incluir falhou em {seletor}: {ultimo_erro}")
        try:
            sucesso = self.page.evaluate(
                """() => {
                    const botao = document.evaluate('//*[@id="btnIncluir"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue
                        || document.querySelector("#btnIncluir, input[name='cmdIncluir'], input[type='submit'], button[type='submit']");
                    if (!botao) return false;
                    botao.disabled = false;
                    botao.removeAttribute("disabled");
                    botao.classList.remove("disabled");
                    botao.style.pointerEvents = "";
                    botao.style.opacity = "";
                    botao.dispatchEvent(new MouseEvent("mouseover", {bubbles: true}));
                    botao.dispatchEvent(new MouseEvent("mousedown", {bubbles: true}));
                    botao.dispatchEvent(new MouseEvent("mouseup", {bubbles: true}));
                    botao.click();
                    return true;
                }"""
            )
            if sucesso:
                self._aguardar_pos_incluir()
                self.log("Clique no botao Incluir enviado via JavaScript.")
                return True
        except Exception as exc:
            ultimo_erro = resumir_excecao(exc)
            self.log(f"Clique JavaScript no Incluir falhou: {ultimo_erro}")
        raise RuntimeError(f"Nao foi possivel acionar o botao Incluir do GEMUL. Ultimo erro: {ultimo_erro}")

    def _preencher_justificativa(self, resultado: Dict[str, Any], arquivo: Optional[Path] = None):
        if not self.page or self.page.is_closed():
            raise RuntimeError("Abra o GEMUL antes de lancar a justificativa.")
        self._ir_para_formulario_justificativa()
        self.maximizar_janela()

        data_inicio = formatar_data_iso_para_br(str(resultado.get("data_inicio") or ""))
        data_fim = formatar_data_iso_para_br(str(resultado.get("data_fim") or ""))
        turma = str(resultado.get("turma") or "").strip()
        nome_aluno = str(resultado.get("nome_aluno") or "").strip()
        matricula = re.sub(r"\D", "", str(resultado.get("matricula") or ""))
        titulo = limpar_texto_gemul(resultado.get("titulo_documento") or "") or "Atestado medico"
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
        self._desativar_validacao_automatica_campos_estaticos()
        self.page.locator("#txtJustificativa").fill(titulo)

        arquivo_upload = None
        if arquivo and arquivo.exists() and arquivo.suffix.lower() in (IMAGE_EXTENSIONS | {".pdf"}):
            arquivo_upload = preparar_documento_para_gemul(arquivo, self.log)
            self._anexar_imagem(arquivo_upload)
        elif arquivo and arquivo.exists():
            self.log("Documento nao anexado no GEMUL: a pagina aceita JPG/PNG no campo de imagem.")

        self._instalar_liberador_incluir()
        if arquivo_upload:
            self.log("Formulario do GEMUL preenchido e imagem anexada. Revise e use o botao Incluir no GEMUL para enviar.")
        else:
            self.log("Formulario do GEMUL preenchido. O programa parou aqui porque nao havia imagem JPG/PNG/PDF para anexar.")
        return str(arquivo_upload) if arquivo_upload else None

    def abrir(self, usuario: str, senha: str):
        return self._executar_no_worker(self._abrir, usuario, senha)

    def _abrir(self, usuario: str, senha: str):
        if self.page and not self.page.is_closed():
            self.page.goto(GEMUL_ATESTADO_URL, wait_until="domcontentloaded", timeout=60000)
            self.maximizar_janela()
            return

        self._fechar()
        ultimo_erro = None
        edge_exe = Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
        planos = [
            ("chromium-padrao", None, None, False),
            ("chromium-persistente", None, None, True),
            ("chrome-persistente", "chrome", None, True),
            ("msedge-persistente", "msedge", None, True),
        ]
        if edge_exe.exists():
            planos.append(("edge-executavel", None, edge_exe, True))
        for tentativa, (nome_plano, canal, executavel, persistente) in enumerate(planos, start=1):
            try:
                self.log(f"Tentando abrir navegador GEMUL: {nome_plano}.")
                self._abrir_nova_sessao(usuario, senha, canal=canal, executavel=executavel, persistente=persistente)
                return
            except Exception as exc:
                ultimo_erro = exc
                self.log(f"Falha ao abrir navegador GEMUL ({nome_plano}, tentativa {tentativa}/{len(planos)}): {resumir_excecao(exc)}")
                self._fechar()
                if tentativa < len(planos):
                    time.sleep(1.5)

        raise RuntimeError(
            "Nao foi possivel abrir o navegador do GEMUL com Chromium, Chrome ou Edge. "
            f"Consulte o log completo em: {LOG_FILE}"
        ) from ultimo_erro

    def _abrir_nova_sessao(
        self,
        usuario: str,
        senha: str,
        canal: Optional[str] = None,
        executavel: Optional[Path] = None,
        persistente: bool = False,
    ):
        self.play = sync_playwright().start()
        opcoes = {"headless": False, "args": ["--start-maximized"]}
        if canal:
            opcoes["channel"] = canal
        if executavel:
            opcoes["executable_path"] = str(executavel)
        if persistente:
            self.perfil_temporario = Path(tempfile.mkdtemp(prefix="gemul_playwright_"))
            self.context = self.play.chromium.launch_persistent_context(
                str(self.perfil_temporario),
                accept_downloads=True,
                viewport=None,
                **opcoes,
            )
            self.browser = self.context.browser
            self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        else:
            self.browser = self.play.chromium.launch(**opcoes)
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
                try:
                    self.context.close()
                except Exception as exc:
                    self.log(f"Aviso ao fechar contexto do navegador: {resumir_excecao(exc)}")
            if self.browser:
                try:
                    self.browser.close()
                except Exception as exc:
                    self.log(f"Aviso ao fechar navegador: {resumir_excecao(exc)}")
            if self.play:
                try:
                    self.play.stop()
                except Exception as exc:
                    self.log(f"Aviso ao encerrar Playwright: {resumir_excecao(exc)}")
        finally:
            self.play = None
            self.browser = None
            self.context = None
            self.page = None
            if self.perfil_temporario:
                try:
                    shutil.rmtree(self.perfil_temporario, ignore_errors=True)
                except Exception as exc:
                    self.log(f"Aviso ao remover perfil temporario: {resumir_excecao(exc)}")
            self.perfil_temporario = None


class AnimacaoEngrenagens(tk.Canvas):
    def __init__(self, master):
        cor_fundo = ttk.Style(master).lookup("TFrame", "background") or "SystemButtonFace"
        super().__init__(master, width=142, height=52, highlightthickness=0, bg=cor_fundo)
        self.ativo = True
        self.trabalhando = False
        self.angulo = 0.0
        self.fumaca = 0.0
        self.after_id = None
        self._animar()

    def iniciar(self):
        self.trabalhando = True
        self.ativo = True
        if not self.after_id:
            self._animar()

    def parar(self):
        self.trabalhando = False
        self.ativo = True
        if not self.after_id:
            self._animar()

    def _animar(self):
        self.delete("all")
        passo = 0.22 if self.trabalhando else 0.055
        self._desenhar_base()
        self._desenhar_chamine()
        self._desenhar_fumaca()
        self._desenhar_engrenagem(42, 29, 15, self.angulo, "#5e6570")
        self._desenhar_engrenagem(75, 29, 12, -self.angulo * 1.25, "#8a6f47")
        self.angulo = (self.angulo + passo) % (math.pi * 2)
        self.fumaca = (self.fumaca + (0.16 if self.trabalhando else 0.05)) % 1.0
        if self.ativo:
            self.after_id = self.after(70, self._animar)
        else:
            self.after_id = None

    def _desenhar_base(self):
        self.create_rectangle(23, 39, 93, 45, fill="#4b3b2e", outline="#2c2520")
        self.create_line(26, 46, 92, 46, fill="#b08a52", width=2)

    def _desenhar_chamine(self):
        self.create_rectangle(101, 18, 116, 43, fill="#5a4637", outline="#2c2520", width=1)
        self.create_rectangle(98, 14, 119, 20, fill="#6f5540", outline="#2c2520", width=1)
        self.create_line(104, 20, 104, 42, fill="#b58a5b")

    def _desenhar_fumaca(self):
        quantidade = 5 if self.trabalhando else 2
        for i in range(quantidade):
            fase = (self.fumaca + i / max(1, quantidade)) % 1.0
            x = 108 + math.sin(fase * math.pi * 2 + i) * (4 + i)
            y = 13 - fase * 18
            raio = 3 + fase * (6 if self.trabalhando else 3)
            tom = 150 + int(70 * fase)
            cor = f"#{tom:02x}{tom:02x}{tom:02x}"
            self.create_oval(x - raio, y - raio, x + raio, y + raio, fill=cor, outline="")

    def _desenhar_engrenagem(self, cx: int, cy: int, raio: int, angulo: float, cor: str):
        dentes = 10
        pontos = []
        for i in range(dentes * 2):
            r = raio + (4 if i % 2 == 0 else 0)
            a = angulo + i * math.pi / dentes
            pontos.extend([cx + math.cos(a) * r, cy + math.sin(a) * r])
        self.create_polygon(pontos, fill=cor, outline="#26323d", width=1)
        self.create_oval(cx - raio + 5, cy - raio + 5, cx + raio - 5, cy + raio - 5, fill=self["bg"], outline="#26323d")
        for i in range(4):
            a = angulo + i * math.pi / 2
            self.create_line(cx, cy, cx + math.cos(a) * (raio - 3), cy + math.sin(a) * (raio - 3), fill="#26323d", width=2)
        self.create_oval(cx - 3, cy - 3, cx + 3, cy + 3, fill="#26323d", outline="")


class LeitorAtestadoApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"Leitor de Atestado v{APP_VERSION}")
        self.root.geometry("980x680")
        self.root.after(100, self.maximizar_janela_principal)
        self.alunos = carregar_alunos()
        self.browser = GemulBrowser(self.log)
        self.arquivo_atual: Optional[Path] = None
        self.resultado_atual: Optional[Dict[str, Any]] = None
        self.escaneando = False
        self.abrindo_gemul = False
        self.incluir_em_andamento = False
        self.operacoes_ativas = 0
        self.config = carregar_configuracao()

        self.pasta_saida_var = tk.StringVar(value=self.config.get("pasta_saida") or str(APP_DIR / "atestados_salvos"))
        self.usuario_var = tk.StringVar(value=os.getenv("GEMUL_USUARIO", self.config.get("gemul_usuario", "")))
        self.senha_var = tk.StringVar(value=os.getenv("GEMUL_SENHA", self.config.get("gemul_senha", "")))
        self.modelo_var = tk.StringVar(value=self.config.get("openai_model") or OPENAI_MODEL)
        self.arquivo_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value=f"{len(self.alunos)} alunos carregados.")
        self._montar_interface()
        self.carregar_log_recente()

    def maximizar_janela_principal(self):
        try:
            self.root.state("zoomed")
        except tk.TclError:
            try:
                self.root.attributes("-zoomed", True)
            except tk.TclError:
                pass

    def trazer_app_para_frente(self):
        def focar():
            try:
                self.root.deiconify()
                self.root.lift()
                self.root.focus_force()
                self.root.attributes("-topmost", True)
                self.root.after(350, lambda: self.root.attributes("-topmost", False))
            except Exception as exc:
                self.log(f"Nao foi possivel trazer a interface para frente via Tk: {resumir_excecao(exc)}")
            try:
                import ctypes

                hwnd = self.root.winfo_id()
                user32 = ctypes.windll.user32
                user32.ShowWindow(hwnd, 9)
                user32.SetForegroundWindow(hwnd)
            except Exception as exc:
                self.log(f"Nao foi possivel alternar foco para a interface via Windows: {resumir_excecao(exc)}")

        self.root.after(0, focar)

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
        ttk.Button(frame_doc, text="Reduzir imagem 2MB", command=self.reduzir_imagem_thread).grid(row=0, column=3, padx=8, pady=8)

        frame_ia = ttk.Frame(self.root)
        frame_ia.grid(row=3, column=0, sticky="ew", padx=10, pady=6)
        frame_ia.columnconfigure(0, weight=1)
        frame_ia.columnconfigure(1, minsize=170)
        frame_comandos = ttk.Frame(frame_ia)
        frame_comandos.grid(row=0, column=0, sticky="w")
        ttk.Label(frame_comandos, text="Modelo").pack(side="left")
        ttk.Entry(frame_comandos, textvariable=self.modelo_var, width=16).pack(side="left", padx=(6, 10))
        ttk.Button(frame_comandos, text="Ler com IA", command=self.ler_ia_thread).pack(side="left")
        ttk.Button(frame_comandos, text="Salvar com nome padrao", command=self.salvar_documento).pack(side="left", padx=(6, 0))
        ttk.Button(frame_comandos, text="Lancar no GEMUL", command=self.lancar_gemul_thread).pack(side="left", padx=(6, 0))
        ttk.Button(frame_comandos, text="Incluir no GEMUL", command=self.incluir_gemul_thread).pack(side="left", padx=(6, 0))
        ttk.Button(frame_comandos, text="Preparar imagem manual", command=self.reanexar_imagem_thread).pack(side="left", padx=(6, 0))
        self.animacao = AnimacaoEngrenagens(frame_ia)
        self.animacao.grid(row=0, column=1, sticky="e", padx=(12, 18))

        frame_resultado = ttk.LabelFrame(self.root, text="Resultado")
        frame_resultado.grid(row=4, column=0, sticky="ew", padx=10, pady=6)
        frame_resultado.columnconfigure(1, weight=1)
        frame_resultado.columnconfigure(3, weight=1)
        frame_resultado.columnconfigure(5, weight=1)
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
            ttk.Label(frame_resultado, text=rotulo).grid(row=i // 3, column=(i % 3) * 2, padx=8, pady=4, sticky="w")
            var = tk.StringVar(value="")
            self.campos[chave] = var
            entrada = ttk.Entry(frame_resultado, textvariable=var)
            entrada.grid(row=i // 3, column=(i % 3) * 2 + 1, padx=8, pady=4, sticky="ew")
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
        ttk.Button(frame_log, text="Abrir arquivo de log", command=self.abrir_arquivo_log).grid(
            row=1, column=0, sticky="e", pady=(6, 0)
        )
        ttk.Label(self.root, textvariable=self.status_var).grid(row=6, column=0, sticky="ew", padx=10, pady=(0, 8))

    def carregar_log_recente(self):
        linhas = ler_ultimas_linhas_log(80)
        if not linhas:
            return
        self.log_text.insert("end", "Ultimas linhas do log anterior:\n")
        for linha in linhas:
            self.log_text.insert("end", f"{linha}\n")
        self.log_text.see("end")

    def log(self, mensagem: str):
        registrar_log_arquivo(mensagem)

        def atualizar():
            carimbo = datetime.now().strftime("%H:%M:%S")
            self.log_text.insert("end", f"[{carimbo}] {mensagem}\n")
            self.log_text.see("end")
            self.status_var.set(mensagem)

        self.root.after(0, atualizar)

    def mostrar_erro(self, titulo: str, mensagem: str):
        self.root.after(0, lambda: messagebox.showerror(titulo, mensagem))

    def abrir_arquivo_log(self):
        try:
            if not LOG_FILE.exists():
                registrar_log_arquivo("Arquivo de log criado.")
            os.startfile(LOG_FILE)
        except Exception as exc:
            self.mostrar_erro("Log", f"Nao foi possivel abrir o log.\n\n{resumir_excecao(exc)}\n\nArquivo: {LOG_FILE}")

    def executar_thread(self, func):
        threading.Thread(target=func, daemon=True).start()

    def iniciar_operacao(self):
        def atualizar():
            self.operacoes_ativas += 1
            self.animacao.iniciar()

        self.root.after(0, atualizar)

    def finalizar_operacao(self):
        def atualizar():
            self.operacoes_ativas = max(0, self.operacoes_ativas - 1)
            if self.operacoes_ativas == 0:
                self.animacao.parar()

        self.root.after(0, atualizar)

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(initialdir=self.pasta_saida_var.get() or str(APP_DIR))
        if pasta:
            self.pasta_saida_var.set(pasta)
            self.salvar_configuracao_atual()

    def selecionar_arquivo(self):
        pasta_imagens = Path.home() / "Pictures"
        caminho = filedialog.askopenfilename(
            initialdir=str(pasta_imagens if pasta_imagens.exists() else APP_DIR),
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
            raise DigitalizacaoEmAndamentoError("Aguarde a digitalizacao terminar antes de usar a IA.")
        texto_caminho = (self.arquivo_var.get() or "").strip()
        caminho = self.arquivo_atual or (Path(texto_caminho) if texto_caminho else None)
        if caminho is None:
            raise FileNotFoundError("Selecione ou escaneie um documento antes de usar a IA.")
        caminho = Path(caminho)
        if not caminho.exists():
            raise FileNotFoundError(f"Documento nao encontrado: {caminho}")
        if not caminho.is_file():
            raise ValueError(f"O caminho selecionado nao e um arquivo: {caminho}")
        if caminho.stat().st_size <= 0:
            raise ValueError(
                f"O documento selecionado esta vazio (0 KB): {caminho}\n\n"
                "Digitalize novamente ou selecione outro arquivo antes de usar a IA."
            )
        if caminho.suffix.lower() not in DOCUMENT_EXTENSIONS:
            raise ValueError(
                f"Formato nao suportado para leitura por IA: '{caminho.suffix or 'sem extensao'}'. "
                "Use PNG, JPG, JPEG, WEBP, TIF, TIFF, PDF ou TXT."
            )
        self.arquivo_atual = caminho
        return caminho

    def abrir_gemul_thread(self):
        if self.abrindo_gemul:
            self.log("Abertura do GEMUL ja esta em andamento; aguarde.")
            return
        self.executar_thread(self.abrir_gemul)

    def abrir_gemul(self):
        try:
            self.abrindo_gemul = True
            self.salvar_configuracao_atual()
            self.log("Abrindo navegador...")
            self.browser.abrir(self.usuario_var.get().strip(), self.senha_var.get().strip())
        except Exception as exc:
            registrar_excecao_log("Erro ao abrir GEMUL", exc)
            resumo = resumir_excecao(exc)
            self.log(f"Erro ao abrir GEMUL: {resumo}. Detalhes completos no log.")
            self.mostrar_erro("GEMUL", f"{resumo}\n\nDetalhes completos no log:\n{LOG_FILE}")
        finally:
            self.abrindo_gemul = False

    def salvar_configuracao_atual(self):
        salvar_configuracao(
            {
                "gemul_usuario": self.usuario_var.get().strip(),
                "gemul_senha": self.senha_var.get().strip(),
                "openai_model": self.modelo_var.get().strip() or OPENAI_MODEL,
                "pasta_saida": self.pasta_saida_var.get().strip() or str(APP_DIR / "atestados_salvos"),
            }
        )

    def escanear_thread(self):
        if self.escaneando:
            self.log("Digitalizacao ja esta em andamento; aguarde.")
            return
        self.executar_thread(self.escanear)

    def escanear(self):
        try:
            self.iniciar_operacao()
            self.escaneando = True
            pasta = Path(self.pasta_saida_var.get() or APP_DIR)
            pasta.mkdir(parents=True, exist_ok=True)
            destino = pasta / f"scan_atestado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
            self.log("Abrindo interface do scanner WIA...")
            caminho = escanear_wia(destino)
            self.definir_arquivo(caminho)
            self.log("Digitalizacao concluida.")
        except Exception as exc:
            registrar_excecao_log("Erro ao escanear", exc)
            resumo = resumir_excecao(exc)
            if isinstance(exc, ScannerWiaError):
                self.log(f"Digitalizacao nao concluida: {resumo}")
            else:
                self.log(f"Erro ao escanear: {resumo}. Detalhes completos no log.")
                self.mostrar_erro("Scanner", f"{resumo}\n\nDetalhes completos no log:\n{LOG_FILE}")
        finally:
            self.escaneando = False
            self.finalizar_operacao()

    def ler_ia_thread(self):
        self.executar_thread(self.ler_ia)

    def lancar_gemul_thread(self):
        self.executar_thread(self.lancar_gemul)

    def incluir_gemul_thread(self):
        self.executar_thread(self.incluir_gemul)

    def reanexar_imagem_thread(self):
        self.executar_thread(self.reanexar_imagem_gemul)

    def reduzir_imagem_thread(self):
        self.executar_thread(self.reduzir_imagem_atual)

    def ler_ia(self):
        try:
            self.iniciar_operacao()
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
            if isinstance(exc, DigitalizacaoEmAndamentoError):
                self.log(str(exc))
            else:
                self.log(f"Erro na leitura por IA: {exc}")
                self.mostrar_erro("Leitura por IA", str(exc))
        finally:
            self.finalizar_operacao()

    def exibir_resultado(self, resultado: Dict[str, Any]):
        for chave, var in self.campos.items():
            valor = resultado.get(chave, "")
            if chave in {"data_inicio", "data_fim"}:
                valor = formatar_data_iso_para_br(str(valor))
            elif chave in {"titulo_documento", "tipo_justificativa"}:
                valor = limpar_texto_gemul(valor)
            var.set(str(valor))
        self.log(f"Aluno identificado: {resultado.get('nome_aluno', '')}")
        if not str(resultado.get("nome_aluno") or "").strip():
            detalhe = resultado.get("observacoes") or resultado.get("texto_encontrado") or "Sem detalhe retornado pela IA."
            self.log(f"IA nao identificou aluno. Detalhe: {limpar_texto_gemul(detalhe)[:300]}")

    def sincronizar_resultado_dos_campos(self) -> Dict[str, Any]:
        resultado = dict(self.resultado_atual or {})
        for chave, var in self.campos.items():
            valor = var.get().strip()
            if valor:
                if chave in {"titulo_documento", "tipo_justificativa"}:
                    valor = limpar_texto_gemul(valor)
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

    def copiar_caminho_imagem_preparada(self, caminho: Path):
        def copiar():
            self.root.clipboard_clear()
            self.root.clipboard_append(str(caminho))
            self.status_var.set(f"Caminho da imagem copiado: {caminho}")

        self.root.after(0, copiar)
        self.log(f"Caminho da imagem copiado para area de transferencia: {caminho}")

    def lancar_gemul(self):
        try:
            self.iniciar_operacao()
            caminho = self.obter_documento_atual()
            self.aplicar_matricula_manual()
            resultado = self.sincronizar_resultado_dos_campos()
            if not resultado.get("matricula") or not resultado.get("nome_aluno"):
                raise RuntimeError("Identifique o aluno antes de lancar no GEMUL.")
            if not resultado.get("data_inicio") or not resultado.get("data_fim"):
                raise RuntimeError("Preencha data inicial e data final antes de lancar no GEMUL.")
            self.log("Preenchendo formulario do GEMUL...")
            self.browser.trazer_para_frente()
            caminho_imagem = self.browser.preencher_justificativa(resultado, caminho)
            if caminho_imagem:
                self.copiar_caminho_imagem_preparada(Path(caminho_imagem))
            self.browser.page_down()
        except Exception as exc:
            registrar_excecao_log("Erro ao lancar no GEMUL", exc)
            resumo = resumir_excecao(exc)
            self.log(f"Erro ao lancar no GEMUL: {resumo}. Detalhes completos no log.")
            self.mostrar_erro("GEMUL", f"{resumo}\n\nDetalhes completos no log:\n{LOG_FILE}")
        finally:
            self.finalizar_operacao()

    def incluir_gemul(self):
        try:
            if self.incluir_em_andamento:
                self.log("Clique em Incluir ignorado: envio ja esta em andamento.")
                return
            self.incluir_em_andamento = True
            self.iniciar_operacao()
            self.browser.trazer_para_frente()
            if self.browser.clicar_incluir():
                self.log("Confira no navegador se o GEMUL exibiu confirmacao ou mensagem de erro.")
        except Exception as exc:
            registrar_excecao_log("Incluir GEMUL", exc)
            resumo = resumir_excecao(exc)
            self.log(f"Erro ao acionar Incluir: {resumo}. Detalhes completos no log.")
            self.mostrar_erro("Incluir GEMUL", f"{resumo}\n\nDetalhes completos no log:\n{LOG_FILE}")
        finally:
            self.incluir_em_andamento = False
            self.finalizar_operacao()
            self.trazer_app_para_frente()

    def reanexar_imagem_gemul(self):
        try:
            self.iniciar_operacao()
            caminho = self.obter_documento_atual()
            imagem = preparar_documento_para_gemul(caminho, self.log)
            self.copiar_caminho_imagem_preparada(imagem)
        except Exception as exc:
            self.log(f"Erro ao preparar imagem para o GEMUL: {exc}")
            self.mostrar_erro("GEMUL", str(exc))
        finally:
            self.finalizar_operacao()

    def reduzir_imagem_atual(self):
        try:
            self.iniciar_operacao()
            caminho = self.obter_documento_atual()
            if caminho.suffix.lower() not in IMAGE_EXTENSIONS:
                raise RuntimeError("Selecione uma imagem para reduzir.")
            imagem = preparar_imagem_para_gemul(caminho, self.log)
            self.arquivo_atual = imagem
            self.root.after(0, lambda: self.arquivo_var.set(str(imagem)))
            self.copiar_caminho_imagem_preparada(imagem)
            tamanho_kb = imagem.stat().st_size // 1024
            self.log(f"Imagem reduzida e selecionada: {imagem} ({tamanho_kb} KB).")
            messagebox.showinfo("Imagem reduzida", f"Imagem pronta com {tamanho_kb} KB:\n{imagem}")
        except Exception as exc:
            self.log(f"Erro ao reduzir imagem: {exc}")
            self.mostrar_erro("Reduzir imagem", str(exc))
        finally:
            self.finalizar_operacao()

    def salvar_documento(self):
        try:
            self.arquivo_atual = self.obter_documento_atual()
            self.aplicar_matricula_manual()
            resultado = self.resultado_atual or {}
            nome_base = resultado.get("nome_arquivo") or "ATESTADO_SEM_ALUNO_IDENTIFICADO"
            pasta = Path(self.pasta_saida_var.get() or APP_DIR)
            pasta.mkdir(parents=True, exist_ok=True)
            origem_salvar = self.arquivo_atual
            extensao_destino = self.arquivo_atual.suffix.lower()
            if self.arquivo_atual.suffix.lower() in (IMAGE_EXTENSIONS | {".pdf"}):
                origem_salvar = preparar_documento_para_gemul(self.arquivo_atual, self.log)
                extensao_destino = ".jpg"
            destino = pasta / f"{limpar_nome_arquivo(nome_base)}{extensao_destino}"
            contador = 2
            while destino.exists():
                destino = pasta / f"{limpar_nome_arquivo(nome_base)} ({contador}){extensao_destino}"
                contador += 1
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
        inicio = time.monotonic()
        dialog = win32com.client.Dispatch("WIA.CommonDialog")
        try:
            imagem = dialog.ShowAcquireImage(
                1,
                WIA_INTENT_GRAYSCALE,
                WIA_BIAS_MINIMIZE_SIZE,
                WIA_FORMAT_JPEG,
                False,
                True,
                False,
            )
        except Exception as exc:
            texto = str(exc)
            if "ocupado" in texto.lower() or "-2145320954" in texto:
                raise ScannerWiaError(
                    "O scanner WIA informou que o dispositivo esta ocupado. "
                    "Feche ou cancele a janela de digitalizacao aberta, aguarde alguns segundos e tente novamente."
                ) from exc
            if "erro nao especificado" in texto.lower() or "-2147467259" in texto:
                raise ScannerWiaError(
                    "O driver WIA retornou um erro generico e nao entregou imagem. "
                    "Feche a janela do scanner e, se o arquivo foi salvo pelo utilitario do fabricante, use 'Selecionar arquivo'."
                ) from exc
            raise ScannerWiaError(
                "O scanner WIA nao concluiu a digitalizacao. "
                "Feche a janela do scanner e tente novamente; se a imagem foi salva fora do programa, use 'Selecionar arquivo'."
            ) from exc
        if imagem is None:
            raise ScannerWiaError(
                "Digitalizacao cancelada ou nenhum scanner retornou imagem. "
                "Se a imagem foi salva fora do programa, use 'Selecionar arquivo'."
            )
        if destino.exists():
            destino.unlink()
        imagem.SaveFile(str(destino))
        if not destino.exists() or destino.stat().st_size <= 0:
            raise RuntimeError(
                "O scanner nao gerou uma imagem valida; o arquivo salvo ficou vazio. "
                "Tente digitalizar novamente."
            )
        registrar_log_arquivo(
            f"Digitalizacao WIA salva em {time.monotonic() - inicio:.1f}s: {destino.name} ({destino.stat().st_size // 1024} KB)."
        )
        return destino
    finally:
        time.sleep(0.4)
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
            app.salvar_configuracao_atual()
            app.browser.fechar()
        finally:
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", ao_fechar)
    root.mainloop()


if __name__ == "__main__":
    main()
