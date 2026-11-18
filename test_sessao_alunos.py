import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import main
import student_sync


class SessaoAlunosTests(unittest.TestCase):
    def test_janela_existente_refaz_login(self):
        browser = main.GemulBrowser(Mock())
        browser.page = Mock()
        browser.page.is_closed.return_value = False
        browser._login_se_necessario = Mock()
        browser.maximizar_janela = Mock()
        browser._abrir('usuario', 'senha')
        browser._login_se_necessario.assert_called_once_with(browser.page)
        self.assertEqual(browser.usuario, 'usuario')

    def test_login_recusado_interrompe_operacao(self):
        browser = main.GemulBrowser(Mock())
        browser.usuario, browser.senha = 'usuario', 'senha'
        page = Mock()
        page.locator.return_value.count.return_value = 1
        with self.assertRaisesRegex(RuntimeError, 'nao aceitou'):
            browser._login_se_necessario(page)

    def test_manutencao_fecha_auxiliar_mesmo_com_erro(self):
        browser = main.GemulBrowser(Mock())
        browser.page = Mock()
        browser.page.is_closed.return_value = False
        browser.context = Mock()
        auxiliar = browser.context.new_page.return_value
        auxiliar.goto.side_effect = RuntimeError('offline')
        with self.assertRaises(RuntimeError):
            browser._manter_sessao()
        auxiliar.close.assert_called_once()
        browser.page.goto.assert_not_called()

    def test_reutiliza_base_do_historico_sem_coletar(self):
        browser = main.GemulBrowser(Mock())
        with patch.object(main, 'base_atualizada_hoje', return_value=True), \
                patch.object(main, 'carregar_alunos_csv', return_value=['aluno']), \
                patch.object(main, 'atualizar_base_diaria') as coleta:
            self.assertEqual(browser._atualizar_alunos(), ['aluno'])
            coleta.assert_not_called()

    def test_csv_compativel_e_integridade_diaria(self):
        with tempfile.TemporaryDirectory() as directory:
            record = dict.fromkeys(student_sync.FIELDS, '')
            record.update(matricula='123456789', nome='ALUNA TESTE', turma='5 A')
            student_sync.salvar_base(directory, [record])
            self.assertTrue(student_sync.base_atualizada_hoje(directory))
            path = Path(directory) / 'lista_alunos.csv'
            self.assertEqual(main.carregar_alunos_csv(path)[0].nome, 'ALUNA TESTE')
            path.write_text('corrompido', encoding='utf-8')
            self.assertFalse(student_sync.base_atualizada_hoje(directory))


if __name__ == '__main__':
    unittest.main()
