# Releases

## v1.3.17 - 2026-10-29

- Refaz o login quando a sessao expira com a janela aberta, inclusive antes de preencher o atestado.
- Mantem a sessao a cada tres minutos sem comandos do leitor, usando uma aba auxiliar para preservar o formulario em revisao.
- Abre o GEMUL automaticamente ao iniciar quando ja existem credenciais salvas.
- Confere a listagem diariamente apos abrir o GEMUL. Reutiliza o CSV do Historico_escolar quando o registro diario e os hashes confirmam que esta atualizado hoje; caso contrario, coleta uma base local com a mesma rotina de relatorio.
- Preserva a base anterior se a coleta falhar e registra o motivo no log. O programa de historico e seus arquivos nao sao alterados.
- A sincronizacao usa student_sync.py; ambos os arquivos devem acompanhar main.py.

## v1.3.16 - 2026-10-27

- Entrada principal consolidada em `main.py`, baseada na v1.3.15. Novas atualizacoes usam o mesmo arquivo.
- Mantem o nome padrao, a pasta de destino e a preparacao da imagem para o GEMUL.
- Remove o original somente depois de copiar e verificar o arquivo final; falhas preservam o original.
- Mantem as imagens intermediarias em pasta temporaria, removida ao encerrar o processo.
- Atualiza o documento selecionado para o arquivo salvo, permitindo continuar o envio ao GEMUL.
- Corrige a preparacao de imagens com dimensoes menores que 700 pixels.

## v1.3.14 - 2026-09-03

- Criada versao separada em `LeitorAtestado_v1.3.14.py`.
- Mantem a fumaca em velocidade mais lenta, com maior volume de particulas.
- Baixa a torre de sinalizacao para alinhar visualmente com engrenagens e chamine.
- Quando uma etapa falha, o botao correspondente fica piscando em vermelho junto com a luz vermelha da torre.
- Mantem a rotina WIA preservada e destaca falhas de escaneamento diretamente no botao `Escanear`.

## v1.3.13 - 2026-09-03

- Criada versao separada em `LeitorAtestado_v1.3.13.py`.
- Aumenta o destaque visual das cores da torre de sinalizacao industrial.
- Remove a haste/base externa da torre, mantendo apenas o corpo com luzes empilhadas.
- Escurece levemente o azul dos botoes pendentes da timeline.
- Limita o reinicio visual da timeline ao acionamento da etapa final `Preparar imagem`.
- Reserva a coluna direita superior exclusivamente para a animacao e amplia o canvas para a fumaca subir ate perto do topo.

## v1.3.12 - 2026-09-03

- Criada versao separada em `LeitorAtestado_v1.3.12.py`.
- Evita reiniciar a timeline quando o escaneamento define o arquivo gerado pelo proprio fluxo.
- Troca o semaforo simples por uma torre de sinalizacao industrial empilhada.
- A torre usa verde para operacao, amarelo para standby e vermelho pulsante quando alguma operacao falha.

## v1.3.11 - 2026-09-03

- Criada versao separada em `LeitorAtestado_v1.3.11.py`.
- Remove o botao manual `Reduzir imagem 2MB`, mantendo a preparacao automatica de imagens abaixo do limite do GEMUL nos fluxos que enviam/salvam.
- Reorganiza os comandos em uma timeline apos o campo `Modelo`, com etapas que mudam de cor quando concluidas.
- Renomeia `Escanear Kyocera/WIA` para `Escanear` e coloca essa acao como primeira etapa do fluxo.
- Ajusta a animacao para desacelerar mais lentamente e adiciona um semaforo industrial antes das engrenagens.
- Define `gpt-5` como modelo padrao quando nao houver configuracao salva.

## v1.3.10 - 2026-09-03

- Criada versao separada em `LeitorAtestado_v1.3.10.py`.
- Prioriza `lista_alunos.xlsx` sobre a planilha antiga com codigo de turma e reforca a validacao do aluno por nome com comparacao tolerante a erros de OCR.
- Suaviza a aceleracao/desaceleracao das engrenagens e da fumaca na animacao.
- Aciona a animacao tambem nos fluxos de login e salvar com nome padrao.

## v1.3.9 - 2026-09-01

- Criada versao separada em `LeitorAtestado_v1.3.9.py`.
- Reserva uma faixa lateral direita para a animacao, mantendo os dados do resultado comprimidos a esquerda.
- Aumenta o canvas da animacao, baixa a arte das engrenagens/chamine e amplia o campo vertical da fumaca.

## v1.3.8 - 2026-08-31

- Criada versao separada em `LeitorAtestado_v1.3.8.py`.
- Suprime popups de erro conhecidos do scanner WIA quando o driver informa ocupado, erro generico ou cancela sem retornar imagem.
- Ao clicar em `Ler com IA` durante digitalizacao, registra apenas aviso no log/status em vez de abrir caixas de erro repetidas.

## v1.3.7 - 2026-08-12

- Criada versao separada em `LeitorAtestado_v1.3.7.py`.
- Substitui o `PageDown` por rolagem real de mouse e `window.scrollBy` no GEMUL.
- O seletor de arquivos passa a abrir na pasta padrao `Pictures/Imagens` do usuario.
- `Salvar com nome padrao` passa a salvar JPG preparado quando a origem for PDF ou imagem.
- Ajusta layout dos comandos/resultado para reservar melhor a area da animacao.

## v1.3.6 - 2026-08-12

- Criada versao separada em `LeitorAtestado_v1.3.6.py`.
- Ao clicar em `Lancar no GEMUL`, traz o navegador/GEMUL para frente antes de iniciar o preenchimento.
- Normaliza todo anexo para JPG RGB baseline com extensao `.jpg`, mesmo quando a imagem original ja tem menos de 2MB.
- Reprocessa tambem PDFs convertidos para garantir arquivo final JPG aceito pelo GEMUL.
- Dispara evento `change` no campo de upload apos definir o arquivo.

## v1.3.5 - 2026-08-12

- Criada versao separada em `LeitorAtestado_v1.3.5.py`.
- Apos `Lancar no GEMUL`, envia uma tecla `PageDown` ao navegador para rolar ate a area final do formulario.
- Ao usar `Incluir no GEMUL`, traz a janela do navegador/GEMUL para frente durante o clique.
- Depois do clique em `Incluir`, tenta devolver o foco para a interface do programa.

## v1.3.4 - 2026-08-12

- Criada versao separada em `LeitorAtestado_v1.3.4.py`.
- Corrige leitura de PDF escaneado: agora o PDF e convertido em imagens antes de enviar para a IA.
- Mantem leitura direta de imagens JPG/PNG/JPEG/WEBP/TIF/TIFF.
- Prepara automaticamente a primeira pagina de PDFs como JPG para anexo manual no GEMUL.
- Adiciona animacao simples de engrenagens com chamine: marcha lenta em standby e mais fumaca durante operacoes.
- Salva localmente credenciais/modelo/pasta usados no teste para reabrir preenchido.
- Anexa automaticamente a imagem no GEMUL pelo campo de upload informado, validando limite de 2MB.
- Adiciona botao `Incluir no GEMUL` com clique robusto no `//*[@id="btnIncluir"]` e fallback por JavaScript.
- Melhora a mensagem quando o WIA informa que o scanner esta ocupado.
- Registra no log detalhes quando a IA nao identifica o aluno.

## v1.3.3 - 2026-06-23

- Criada versao separada em `LeitorAtestado_v1.3.3.py`.
- Antes de lancar novo atestado, limpa o editor de imagem, `txtTexto` e campos ocultos do GEMUL.
- Reaplica a limpeza apos recarregamentos causados por tipo/turma, evitando sobra de dados anteriores.
- Aumenta retry/espera quando o GEMUL ainda esta processando inclusao anterior em `sca_justificativa_faltas_cod.asp?acao=I`.

## v1.3.2 - 2026-06-23

- Criada versao separada em `LeitorAtestado_v1.3.2.py`.
- Desativa a validacao automatica do GEMUL ao sair de `Data Inicial` e `Data Final` depois do preenchimento.
- Evita o alerta de caractere `<` quando o usuario edita data apos anexar manualmente uma imagem.
- Mantem anexo da imagem e clique em `Incluir` como etapas manuais.

## v1.3.1 - 2026-06-23

- Criada versao separada em `LeitorAtestado_v1.3.1.py`.
- Adiciona botao `Reduzir imagem 2MB` para compactar qualquer imagem selecionada, inclusive arquivos antigos ou externos.
- Apos reduzir, atualiza o documento atual para a copia `_gemul_2mb.jpg` e copia o caminho para a area de transferencia.
- Mantem o fluxo estavel: preenchimento automatico dos dados, anexo e clique em `Incluir` manuais.

## v1.3.0 - 2026-06-23

- Criada versao separada em `LeitorAtestado_v1.3.0.py`.
- Define o fluxo final estavel: preencher todos os dados no GEMUL e parar antes do anexo/inclusao.
- Remove do fluxo operacional qualquer tentativa de anexar imagem ou clicar em `Incluir` automaticamente.
- Mantem preparo da imagem abaixo de 2MB e copia o caminho para facilitar o upload manual.
- Renomeia o botao auxiliar para `Preparar imagem manual`.

## v1.2.2 - 2026-06-23

- Criada versao separada em `LeitorAtestado_v1.2.2.py`.
- Adiciona botao `Clicar Incluir` para acionar o botao `Incluir` do GEMUL.
- Antes de clicar, valida se existe imagem anexada no editor do GEMUL.
- Aceita automaticamente a confirmacao normal de inclusao e registra qualquer alerta do GEMUL no log.
- Mantem o anexo da imagem como passo manual para evitar residuo oculto com caractere `<`.

## v1.2.1 - 2026-06-23

- Criada versao separada em `LeitorAtestado_v1.2.1.py`.
- Desativa o anexo automatico no editor WYSIWYG do GEMUL para eliminar residuo oculto com caractere `<`.
- Mantem preenchimento automatico de todos os campos textuais/datas/turma/aluno.
- Prepara a imagem compactada abaixo de 2MB e copia o caminho para a area de transferencia.
- Troca `Reanexar imagem` por `Copiar caminho imagem`, deixando o anexo manual como ultimo passo antes de clicar em `Incluir`.

## v1.2.0 - 2026-06-23

- Criada versao separada em `LeitorAtestado_v1.2.0.py`.
- Maximiza a janela principal do programa ao abrir.
- Troca o anexo do GEMUL para fluxo mais proximo do clique manual no seletor de arquivo.
- Instala protecao no botao `Incluir` para limpar caracteres `<` e `>` de campos textuais antes da validacao do GEMUL.
- Sanitiza titulo/tipo extraidos pela IA antes de preencher `Ocorrencia`.
- Trata navegacao interrompida do GEMUL quando uma inclusao anterior ainda esta processando.
- Ajusta maximizacao do navegador para restaurar estado normal antes de maximizar.

## v1.1.1 - 2026-06-23

- Criada versao separada em `LeitorAtestado_v1.1.1.py`.
- Adiciona botao `Reanexar imagem` para recolocar o anexo se uma validacao/reload do GEMUL apagar a imagem.
- Centraliza a rotina de anexo da imagem no navegador, limpando o editor antes de novo upload.
- Mantem o comportamento de nao clicar automaticamente em `Incluir`.

## v1.1.0 - 2026-06-23

- Criada versao separada em `LeitorAtestado_v1.1.0.py`, preservando `LeitorAtestado.py`.
- Garante imagem abaixo de 2MB para envio ao GEMUL, reduzindo dimensao e qualidade em etapas.
- Salva copia nomeada do atestado em JPG ja reduzido quando o documento for imagem.
- Antes de anexar no GEMUL, limpa residuos antigos do editor para evitar bloqueio por caractere `<`.
- Valida o tamanho do arquivo antes de anexar e interrompe com erro claro se nao for possivel reduzir.

## v1.0.2 - 2026-06-23

- Corrige preenchimento da `Data Final` no GEMUL quando a `Data Inicial` dispara carregamento/validacao.
- Passa a preencher datas com eventos de tela e verificacao de permanencia do valor no campo.

## v1.0.1 - 2026-06-23

- Corrige erro do Playwright `cannot switch to a different thread` ao lancar no GEMUL.
- Centraliza abertura, preenchimento e fechamento do navegador em uma unica thread interna.
- Mantem o preenchimento do GEMUL sem clicar automaticamente em `Incluir`.
- Reforca o anexo da imagem escaneada no campo de upload do editor do GEMUL.

## v1.0.0 - 2026-06-23

Marco funcional do Leitor de Atestado.

- Abre o GEMUL em navegador maximizado.
- Escaneia documentos via WIA/Kyocera com inicializacao COM correta para leituras sucessivas.
- Envia imagem otimizada para a OpenAI, preservando o arquivo original.
- Extrai aluno, matricula, turma, periodo, titulo do documento e tipo de justificativa.
- Permite corrigir/definir aluno digitando a matricula manualmente.
- Limpa dados ao receber novo documento, evitando reaproveitar resultado anterior.
- Salva documentos com nome padrao baseado em matricula, turma e aluno.
- Mantem log visual e persistente em `leitor_atestado.log`.
- Preenche automaticamente o formulario de Justificativa de Faltas no GEMUL para revisao manual antes de incluir.
