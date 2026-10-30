"""Постраничная выборка: порядок, поиск по всему реестру и границы страниц."""
import tempfile
import unittest

from sz import actions
from sz.store import Store
from test_core import seed, memo_data


def fill_registry(store, count):
    seed(store)
    memo_id, error = actions.create_memo(store, memo_data(body='Обычный текст'), [1])
    assert error is None, error
    template = store.one('SELECT * FROM memos WHERE id = ?', (memo_id,))
    keys = [key for key in template if key != 'id']

    def insert(s):
        for number in range(2, count + 1):
            row = dict(template, number=number, reg_number='НТЦ-%04d-26' % number,
                       body='ОБОРУДОВАНИЕ 100%_готово' if number <= 3 else 'Обычный текст')
            rid = s.run('INSERT INTO memos (%s) VALUES (%s)' % (
                ','.join(keys), ','.join('?' for _ in keys)), [row[key] for key in keys])
            s.run('INSERT INTO memo_addressees (memo_id, ord, department, head_position, '
                  'head_name, head_name_dative) VALUES (?, 0, ?, ?, ?, ?)',
                  (rid, 'Бухгалтерия' if number <= 3 else 'Отдел кадров',
                   'Начальнику', 'Петров Пётр Петрович', 'Петрову П.П.'))
    store.write(insert)


class PagingTest(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.TemporaryDirectory()
        self.store = Store(self.root.name)
        self.store.open()
        fill_registry(self.store, 123)

    def tearDown(self):
        self.store.close()
        self.root.cleanup()

    def test_pages_are_complete_ordered_and_limited_in_sql(self):
        statements = []
        self.store.db.set_trace_callback(statements.append)
        pages = [actions.search_memos_page(self.store, page=i) for i in range(3)]
        self.assertEqual([len(p['memos']) for p in pages], [50, 50, 23])
        self.assertTrue(all(p['matched'] == p['total'] == 123 for p in pages))
        self.assertEqual([m['number'] for p in pages for m in p['memos']], list(range(123, 0, -1)))
        selections = [q for q in statements if q.startswith('SELECT * FROM memos')]
        self.assertEqual(len(selections), 3)
        for query in selections:
            self.assertIn('LIMIT 50 OFFSET', query)
        self.assertEqual(pages[-1]['page_count'], 3)

    def test_search_and_filters_apply_before_paging(self):
        result = actions.search_memos_page(self.store, q='оборудование', year=2026, topic_id=1,
                                          executor_name='Сидоров Сидор Сидорович',
                                          addressee_department='Бухгалтерия', page=2)
        self.assertEqual(result['page'], 0)
        self.assertEqual(result['matched'], 2)
        self.assertEqual([m['number'] for m in result['memos']], [3, 2])
        literal = actions.search_memos_page(self.store, q='100%_ГОТОВО')
        self.assertEqual(literal['matched'], 2)
        self.assertEqual(actions.search_memos_page(self.store, q='100_ГОТОВО')['matched'], 0)
        self.assertEqual(actions.search_memos_page(self.store, executor_name='Нет такого')['memos'], [])

    def test_empty_and_out_of_range_pages_after_deletion(self):
        self.assertEqual(actions.search_memos_page(self.store, page=-1)['page'], 0)
        self.assertEqual(actions.search_memos_page(self.store, page=99)['page'], 2)
        self.store.write(lambda s: s.run('DELETE FROM memos WHERE number <= 23'))
        result = actions.search_memos_page(self.store, page=2)
        self.assertEqual(result['page'], 1)
        self.assertEqual(len(result['memos']), 50)
        self.store.write(lambda s: s.run('DELETE FROM memos'))
        result = actions.search_memos_page(self.store, page=1)
        self.assertEqual((result['total'], result['matched'], result['page'], result['page_count']),
                         (0, 0, 0, 1))
        self.assertEqual(result['memos'], [])


if __name__ == '__main__':
    unittest.main()
