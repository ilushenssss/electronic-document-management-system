"""Регрессии макета Word и редактирования существующих адресатов."""
import io
import json
import os
import sys
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from sz import actions, docx
from sz.store import Store
from test_core import seed, memo_data

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = '{%s}' % NS['w']


class LayoutTest(unittest.TestCase):
    def test_document_parts_and_layout(self):
        for rich in (None, json.dumps([{'runs': [{'text': 'Тело', 'bold': True}],
                                       'align': 'both', 'indent': True}])):
            memo = {'body': 'Тело\nВторой абзац', 'body_rich': rich,
                    'executor_name': 'Иванов Иван Иванович', 'executor_phone': '12-34 "доб." & 5'}
            data = docx.build_docx(memo, [{'head_position': 'Начальнику\nотдела',
                                         'head_name': 'Петров Пётр Петрович'}])
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for part in z.namelist():
                    ET.fromstring(z.read(part))
                document = ET.fromstring(z.read('word/document.xml'))
                margins = document.find('.//w:pgMar', NS)
                for name, cm in [('top', 1.5), ('left', 2.5), ('bottom', 1), ('right', 1)]:
                    self.assertEqual(margins.get(W + name), str(docx.twips(cm)))
                body_paras = [p for p in document.findall('.//w:p', NS)
                              if 'Тело' in ''.join(p.itertext())]
                self.assertTrue(body_paras)
                for p in body_paras:
                    self.assertEqual(p.find('w:pPr/w:spacing', NS).get(W + 'line'), '360')
                position = next(p for p in document.findall('.//w:p', NS)
                                if 'Начальнику' in ''.join(p.itertext()))
                self.assertIsNotNone(position.find('.//w:br', NS))
                self.assertNotIn('Исп.', ''.join(document.itertext()))
                footer = ET.fromstring(z.read('word/footer1.xml'))
                instructions = ''.join(e.text or '' for e in footer.findall('.//w:instrText', NS))
                self.assertIn(' IF ', instructions)
                self.assertIn(' PAGE ', instructions)
                self.assertIn(' NUMPAGES ', instructions)
                self.assertIn('Исп. Иванов И.И.', ''.join(footer.itertext()))
                self.assertIn('Тел. 12-34', ''.join(footer.itertext()))
                self.assertIsNotNone(document.find('.//w:footerReference', NS))
                settings = ET.fromstring(z.read('word/settings.xml'))
                self.assertEqual(settings.find('w:updateFields', NS).get(W + 'val'), 'true')

    def test_edit_existing_addressee_and_preserve_memo_snapshot(self):
        with tempfile.TemporaryDirectory() as root:
            s = Store(root)
            s.open()
            try:
                seed(s)
                old_id, error = actions.create_memo(s, memo_data(), [1])
                self.assertIsNone(error)
                row = s.one('SELECT * FROM addressees WHERE id = 1')
                row['head_position'] = 'Начальнику\r\nотдела кадров'
                actions.save_directory_row(s, 'addressees', row, 1)
                self.assertEqual(s.one('SELECT COUNT(*) AS n FROM addressees')['n'], 1)
                self.assertEqual(s.memo_addressees(old_id)[0]['head_position'], 'Начальнику отдела кадров')
            finally:
                s.close()
            s = Store(root)
            s.open()
            try:
                self.assertEqual(s.one('SELECT * FROM addressees WHERE id = 1')['head_position'],
                                 'Начальнику\nотдела кадров')
                new_id, error = actions.create_memo(s, memo_data(), [1])
                self.assertIsNone(error)
                self.assertEqual(s.memo_addressees(new_id)[0]['head_position'], 'Начальнику\nотдела кадров')
            finally:
                s.close()


if __name__ == '__main__':
    unittest.main()
