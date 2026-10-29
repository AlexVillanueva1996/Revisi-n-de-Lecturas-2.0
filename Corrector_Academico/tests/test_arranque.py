import json
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch, Mock

from src.arranque import leer_instalacion, motor_local, solicitar, verificar_lenguas


class Arranque(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def test_solo_loopback(self):
        for url in ['https://api.languagetool.org/v2/check','http://localhost.ejemplo/v2/check',
                    'http://127.0.0.1@externo.test/v2/check','http://user@127.0.0.1/v2/check']:
            with self.subTest(url=url),self.assertRaises(ValueError): solicitar(url)

    def test_marcador_incompleto_no_arranca(self):
        path=self.root/'instalacion.json'; path.write_text('{"version":2}')
        with self.assertRaises(ValueError): leer_instalacion(path)

    def test_marcador_valido_con_bom(self):
        exe=self.root/'programa'; exe.touch()
        lt=self.root/'lt'; lt.mkdir(); (lt/'languagetool-server.jar').touch()
        data={k:str(exe) for k in ('python','java','dot','soffice')}
        data.update({'version':2,'languagetool':str(lt)})
        path=self.root/'instalacion.json'; path.write_text(json.dumps(data),encoding='utf-8-sig')
        self.assertEqual(leer_instalacion(path)['version'],2)

    def test_cierra_unicamente_el_proceso_creado(self):
        process=Mock(); process.poll.return_value=None
        info={'java':'java','languagetool':str(self.root)}
        with patch('src.arranque.subprocess.Popen',return_value=process) as spawn, \
             patch('src.arranque.solicitar',return_value=[{'code':'es'},{'code':'en'}]):
            with motor_local(info,self.root/'logs',timeout=1) as url:
                self.assertTrue(url.startswith('http://127.0.0.1:'))
            process.terminate.assert_called_once(); process.wait.assert_called_once()
            self.assertNotIn('--public',spawn.call_args.args[0])

    def test_fallo_java_se_declara_y_no_reusa_otro_servidor(self):
        process=Mock(); process.poll.return_value=1; process.returncode=1
        with patch('src.arranque.subprocess.Popen',return_value=process), self.assertRaises(RuntimeError):
            with motor_local({'java':'java','languagetool':str(self.root)},self.root/'logs',timeout=.1): pass

    def test_ambos_idiomas_obligatorios(self):
        with patch('src.arranque.solicitar',return_value={'matches':[{'rule':'prueba'}]}) as request:
            self.assertEqual(verificar_lenguas('http://127.0.0.1:12345'),{'en-US':1,'es':1})
            self.assertEqual(request.call_count,2)
        with patch('src.arranque.solicitar',return_value={'matches':[]}):
            with self.assertRaises(ValueError): verificar_lenguas('http://127.0.0.1:12345')


if __name__=='__main__': unittest.main()
