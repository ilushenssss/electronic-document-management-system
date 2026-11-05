"""Номер СЭД и PDF: отдельное сохранение, безопасная замена и переименование."""
import os
import re
import shutil
import struct
import tempfile
import time
import uuid

from . import docx
from .store import StoreError, split_path


class ScanExists(StoreError):
    pass


def outlook_pdf_name(mime):
    """Имя единственного виртуального PDF из Windows FILEGROUPDESCRIPTOR.

    Содержимое вложения на стадии dragEnter не запрашиваем: Outlook может
    получать его из почты только при отпускании мыши.
    """
    for descriptor, encoding, filename_size in [('FileGroupDescriptorW', 'utf-16-le', 520),
                                                 ('FileGroupDescriptor', 'cp1251', 260)]:
        fmt = 'application/x-qt-windows-mime;value="%s"' % descriptor
        if not mime.hasFormat(fmt):
            continue
        data = bytes(mime.data(fmt))
        # cItems (4 байта), FILEDESCRIPTOR: имя начинается со смещения 72.
        if len(data) < 76 + filename_size or struct.unpack_from('<I', data)[0] != 1:
            return None
        name = data[76:76 + filename_size].decode(encoding, errors='replace').split('\0', 1)[0]
        if name.lower().endswith('.pdf'):
            return name
    return None


def outlook_pdf_bytes(mime):
    if not outlook_pdf_name(mime):
        raise StoreError('Перетащите одно вложение PDF из Outlook.')
    formats = ['application/x-qt-windows-mime;value="FileContents";index=0',
               'application/x-qt-windows-mime;value="FileContents"']
    for fmt in formats:
        if mime.hasFormat(fmt):
            data = bytes(mime.data(fmt))
            if data:
                if b'%PDF-' not in data[:1024]:
                    raise StoreError('Вложение Outlook не содержит заголовка PDF.')
                return data
    raise StoreError('Outlook не передал содержимое PDF. Сохраните вложение на диск '
                     'и используйте «Выбрать PDF» или перетащите файл из Проводника.')


def clean_number(value):
    number = str(value or '').strip().upper()
    if len(number) > 120 or not re.fullmatch(r'[0-9]+/[0-9]+/[0-9]+/СЗ', number):
        raise StoreError('Введите номер СЭД в формате 1/441062/1100/СЗ: три группы цифр и /СЗ.')
    return number


def pdf_name(number):
    return clean_number(number).replace('/', '_') + '.pdf'


def memo_folder(store, memo):
    if memo.get('file_path'):
        return split_path(memo['file_path'])[:-1]
    if memo.get('scan_path'):
        return split_path(memo['scan_path'])[:-1]
    return split_path(store.output_dir) + [str(memo['year']), docx.file_base_name(memo)]


def target_path(store, memo, number):
    return '/'.join(memo_folder(store, memo) + [pdf_name(number)])


def document_path(store, memo):
    """Также находит PDF, вручную помещённый под именем номера СЭД."""
    if not memo:
        return None
    if memo.get('scan_path'):
        return store.document_path(memo['scan_path'])
    if memo.get('sed_number'):
        return store.document_path(target_path(store, memo, memo['sed_number']))
    return None


def check_pdf(path):
    if not os.path.isfile(path) or os.path.splitext(path)[1].lower() != '.pdf':
        raise StoreError('Выберите один файл PDF.')
    with open(path, 'rb') as file:
        if b'%PDF-' not in file.read(1024):
            raise StoreError('Файл не содержит заголовка PDF. Выберите скан в формате PDF.')


def save(store, memo_id, number, source_path=None, expected=None, overwrite=False):
    """Сохраняет номер отдельно или вместе с копией PDF; исходник загрузки не удаляет.

    Файл готовится до записи. При ошибке БД/замены восстанавливается прежний
    документ. Смена номера переименовывает уже привязанный скан.
    """
    number = clean_number(number)
    if source_path is not None:
        source_path = os.path.abspath(source_path)
        check_pdf(source_path)
    temp_path = backup = None
    installed = moved = transaction = False
    with store._lock():
        memo = store.one('SELECT * FROM memos WHERE id = ?', (memo_id,))
        if not memo:
            raise StoreError('СЗ удалена. Скан не загружен.')
        current = (memo['sed_number'], memo['scan_path'])
        if expected is not None and tuple(expected) != current:
            raise StoreError('Номер СЭД или скан уже изменён другим пользователем. Закройте окно '
                             'ввода и перечитайте карточку СЗ.')
        rel = target_path(store, memo, number)
        destination = store.abs_path(rel)
        old_path = document_path(store, memo)
        source = source_path or old_path
        same_file = bool(source and os.path.exists(destination) and os.path.samefile(source, destination))
        if source and not same_file and os.path.exists(destination) and not overwrite:
            raise ScanExists('В папке СЗ уже есть файл %s. Заменить его? '
                             'Прежняя версия будет сохранена.' % os.path.basename(destination))
        if not source and os.path.exists(destination):
            check_pdf(destination)  # PDF мог быть помещён в папку вручную
        if source:
            check_pdf(source)
        try:
            if source and not same_file:
                os.makedirs(os.path.dirname(destination), exist_ok=True)
                if source_path:
                    with tempfile.NamedTemporaryFile(dir=os.path.dirname(destination), suffix='.tmp', delete=False) as file:
                        temp_path = file.name
                        with open(source, 'rb') as original:
                            shutil.copyfileobj(original, file)
                        file.flush()
                        os.fsync(file.fileno())
                if os.path.exists(destination):
                    directory = os.path.join(os.path.dirname(destination), 'Предыдущие версии')
                    os.makedirs(directory, exist_ok=True)
                    backup = os.path.join(directory, '%s-%s-%s.pdf' % (
                        os.path.splitext(os.path.basename(destination))[0], time.strftime('%Y%m%d-%H%M%S'),
                        uuid.uuid4().hex[:8]))
                    shutil.copy2(destination, backup)
            scan_rel = rel if source or os.path.isfile(destination) else ''
            store.db.execute('BEGIN IMMEDIATE')
            transaction = True
            store.run('UPDATE memos SET sed_number = ?, scan_path = ? WHERE id = ?', (number, scan_rel, memo_id))
            if source and not same_file:
                os.replace(temp_path if source_path else source, destination)
                installed = True
                moved = not bool(source_path)
                temp_path = None
            store.db.execute('COMMIT')
            transaction = False
        except BaseException as error:
            if transaction:
                store.db.execute('ROLLBACK')
            # Возврат файла нужен, если замена прошла, а COMMIT завершился ошибкой.
            try:
                if installed:
                    if moved:
                        os.replace(destination, source)
                    else:
                        os.unlink(destination)
                    if backup:
                        shutil.copy2(backup, destination)
            except OSError as restore_error:
                raise StoreError('Не удалось завершить сохранение и восстановить файл. '
                                 'Проверьте папку СЗ и «Предыдущие версии»: %s' % restore_error) from error
            if isinstance(error, PermissionError):
                raise StoreError('PDF открыт или нет прав на запись. Закройте файл и повторите; '
                                 'номер СЭД не изменён.') from error
            raise
        finally:
            if temp_path and os.path.exists(temp_path):
                os.unlink(temp_path)
        store._version = store._data_version()
        return {'sed_number': number, 'scan_path': scan_rel}
