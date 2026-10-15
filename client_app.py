# -*- coding: utf-8 -*-
"""
Клиентское приложение «Журнал регистрации приказов»
====================================================
Десктопный клиент (Tkinter, стандартная библиотека Python) к серверу
веб-приложения. Подключается по REST API v1 (http://адрес_сервера:порт/api/v1/...).

Работает на Windows и ALT Linux без установки сторонних библиотек.

Запуск:
    python3 client_app.py                     # окно с настройкой подключения
    python3 client_app.py --server http://192.168.0.10:8080

Возможности клиента (в пределах прав роли пользователя):
  - просмотр журналов Опека / Нуждаемость / Профилактика, поиск;
  - создание, редактирование и удаление приказов (+ примечание);
  - прикрепление, скачивание и удаление PDF-сканов;
  - справочник «Виды приказов» (просмотр; редактирование — куратору и выше);
  - справочник «Сотрудники» (просмотр; логин виден админу/начальнику отдела);
  - отчёты за период (все три), выгрузка таблицы в CSV/Excel;
  - самостоятельная смена пароля.

Пароль хранится только в памяти процесса (в реестр/файлы не пишется).

Если графическая библиотека Tkinter недоступна, запуск из командной строки
позволяет работать с сервером в текстовом режиме:
    python3 client_app.py --server http://хост:порт --cli --login admin
"""
import sys
import json
import argparse
import urllib.request
import urllib.error
import urllib.parse
import base64
import uuid
import csv
import os
from datetime import date

try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog, simpledialog
    HAS_TK = True
except Exception:            # Tkinter отсутствует — доступен режим --cli
    tk = None
    HAS_TK = False

CLIENT_VERSION = '1.4.0'


# ================================================================ HTTP-клиент
class ApiError(Exception):
    def __init__(self, msg, status=None):
        super().__init__(msg)
        self.status = status


class ServerClient:
    """Минимальный HTTP-клиент к REST API сервера (stdlib urllib)."""

    def __init__(self, base_url, login, password):
        self.base = base_url.rstrip('/')
        self.login = login
        token = f'{login}:{password}'.encode('utf-8')
        self.auth = 'Basic ' + base64.b64encode(token).decode('ascii')

    def _request(self, path, method='GET', data=None, files=None, raw=False):
        url = self.base + path
        headers = {'Authorization': self.auth}
        body = None
        if files is not None:                      # multipart/form-data
            boundary = '----pyclient' + uuid.uuid4().hex
            parts = []
            for name, value in (data or {}).items():
                parts.append(f'--{boundary}\r\nContent-Disposition: form-data; '
                             f'name="{name}"\r\n\r\n{value}\r\n'.encode('utf-8'))
            for name, fname, content in files:
                parts.append(('--%s\r\nContent-Disposition: form-data; name="%s"; '
                              'filename="%s"\r\nContent-Type: application/octet-stream'
                              '\r\n\r\n' % (boundary, name, fname)).encode('utf-8'))
                parts.append(content)
                parts.append(b'\r\n')
            parts.append(('--%s--\r\n' % boundary).encode('utf-8'))
            body = b''.join(parts)
            headers['Content-Type'] = 'multipart/form-data; boundary=' + boundary
        elif data is not None:                     # JSON
            body = json.dumps(data, ensure_ascii=False).encode('utf-8')
            headers['Content-Type'] = 'application/json; charset=utf-8'
        req = urllib.request.Request(url, data=body, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                payload = resp.read()
                if raw:
                    return payload, resp.headers
                ct = resp.headers.get('Content-Type', '')
                if 'json' in ct:
                    return json.loads(payload.decode('utf-8'))
                return payload
        except urllib.error.HTTPError as e:
            payload = e.read()
            if e.code == 401:
                raise ApiError('Неверный логин или пароль', 401)
            if e.code == 403:
                raise ApiError('Нет прав для этого действия', 403)
            try:
                msg = json.loads(payload.decode('utf-8')).get('error')
            except Exception:
                msg = None
            raise ApiError(msg or f'Ошибка сервера (HTTP {e.code})', e.code)
        except urllib.error.URLError as e:
            raise ApiError(f'Сервер недоступен: {e.reason}', None)

    # --- convenience wrappers -------------------------------------------
    def get(self, path):
        return self._request(path)

    def post(self, path, data=None, files=None):
        return self._request(path, 'POST', data, files)

    def put(self, path, data):
        return self._request(path, 'PUT', data)

    def delete(self, path):
        return self._request(path, 'DELETE')

    def get_raw(self, path):
        return self._request(path, raw=True)


def check_server(base_url):
    """Пинг публичного /api/v1/info — для проверки адреса при входе."""
    c = ServerClient(base_url, '__probe__', '__probe__')
    try:
        info = c.get('/api/v1/info')
        return info
    except ApiError as e:
        if e.status == 401:                 # сервер живой, но требует авторизацию
            return {}
        raise


# ================================================================== окно входа

# ==================================================================== CLI-режим
def _pad(s, w):
    s = str(s)
    return s if len(s) >= w else s + ' ' * (w - len(s))


def _table(rows, headers):
    widths = [max(len(str(h)), *(len(str(r[i])) for r in rows)) if rows else len(str(h))
              for i, h in enumerate(headers)]
    line = '+-' + '-+-'.join('-' * w for w in widths) + '-+'
    out = [line, '| ' + ' | '.join(_pad(h, w) for h, w in zip(headers, widths)) + ' |',
           line]
    for r in rows:
        out.append('| ' + ' | '.join(_pad(c, w) for c, w in zip(r, widths)) + ' |')
    out.append(line)
    return '\n'.join(out)


def cli_main(base, login, password):
    """Интерактивный текстовый клиент (для систем без Tkinter / по SSH)."""
    c = ServerClient(base, login, password)
    sess = c.get('/api/v1/session')
    journals = {j['title']: j['key'] for j in sess['journals']}
    print(f"Вход выполнен: {sess['full_name']} ({sess['role']}), "
          f"сервер v{sess['version']}")
    help_txt = ('l <№ журнала>            — список приказов (с поиском)\n'
                'a <№ журнала>            — создать приказ (по шагам)\n'
                'e <id приказа>           — изменить приказ\n'
                'd <id приказа>           — удалить приказ\n'
                'u <id приказа> <файл.pdf> — прикрепить скан\n'
                'p <id приказа> [файл]    — скачать скан\n'
                't                        — справочник видов приказов\n'
                'r                        — отчёт «Приказы за период»\n'
                'w                        — сотрудники\n'
                'x                        — сменить пароль\n'
                'h                        — эта справка, q — выход')
    print(help_txt)
    while True:
        try:
            cmd = input('\nклиент> ').strip()
        except EOFError:
            break
        if not cmd:
            continue
        parts = cmd.split()
        op, arg = parts[0].lower(), parts[1] if len(parts) > 1 else ''
        try:
            if op in ('q', 'exit'):
                break
            elif op == 'h':
                print(help_txt)
            elif op in ('l', 'list'):
                if arg not in journals:
                    print('Укажите номер журнала:',
                          ', '.join(f'{i+1}:{t}' for i, t in enumerate(journals)))
                    continue
                title = [t for t in journals if t == arg][0] if arg in journals else arg
                q = input('Строка поиска (Enter — все): ').strip()
                url = f'/api/v1/journal/{journals[title]}/orders'
                if q:
                    url += '?q=' + urllib.parse.quote(q)
                rows = c.get(url)
                print(_table([[r['id'], r['order_number'], r['order_date'],
                               r['order_type_name'] or '', r['person_fio'],
                               'да' if r['has_scan'] else ''] for r in rows],
                             ['ID', '№', 'Дата', 'Вид приказа', 'Ф.И.О.', 'Скан']))
            elif op == 'a':
                titles = list(journals)
                try:
                    title = titles[int(arg) - 1]
                except (ValueError, IndexError):
                    print('Номера журналов:',
                          ', '.join(f'{i+1}={t}' for i, t in enumerate(titles)))
                    continue
                data = {}
                data['order_number'] = input('№ приказа (пусто = авто): ').strip()
                data['order_date'] = input(
                    f'Дата ГГГГ-ММ-ДД [{date.today()}]: ').strip() or date.today().isoformat()
                data['person_fio'] = input('Ф.И.О. в отношении кого: ').strip()
                data['specialist_fio'] = input('ФИО специалиста: ').strip()
                types = c.get(f"/api/v1/order_types?journal={journals[title]}")
                for i, t in enumerate(types, 1):
                    print(f'  {i}. {t["name"]}')
                tn = input('Номер вида приказа (0 — без): ').strip()
                data['order_type_id'] = types[int(tn) - 1]['id'] \
                    if tn.isdigit() and 0 < int(tn) <= len(types) else None
                data['note'] = input('Примечание: ').strip()
                res = c.post(f"/api/v1/journal/{journals[title]}/orders", data)
                print(f"Создан приказ № {res['order_number']} (id={res['id']})")
            elif op == 'e':
                order = c.get(f'/api/v1/orders/{arg}')
                print(f"Редактирование № {order['order_number']}; Enter = оставить как есть")
                upd = {}
                for f, label in (('order_number', '№'), ('order_date', 'Дата'),
                                 ('person_fio', 'Ф.И.О.'),
                                 ('specialist_fio', 'Специалист'), ('note', 'Примечание')):
                    v = input(f'{label} [{order[f]}]: ').strip()
                    if v:
                        upd[f] = v
                if upd:
                    c.put(f'/api/v1/orders/{arg}', upd)
                    print('Сохранено')
            elif op == 'd':
                if input(f'Удалить приказ id={arg}? (д/н) ').strip().lower() == 'д':
                    c.delete(f'/api/v1/orders/{arg}')
                    print('Удалён')
            elif op == 'u':
                with open(arg and parts[2], 'rb') as fh:
                    content = fh.read()
                c.post(f'/api/v1/orders/{arg}/scan', files=[
                    ('file', os.path.basename(parts[2]), content)])
                print('Скан прикреплён')
            elif op == 'p':
                payload, _ = c.get_raw(f'/api/v1/orders/{arg}/scan')
                out = parts[2] if len(parts) > 2 else f'scan_{arg}.pdf'
                with open(out, 'wb') as fh:
                    fh.write(payload)
                print(f'Скан сохранён: {out} ({len(payload)} байт)')
            elif op == 't':
                rows = c.get('/api/v1/order_types')
                print(_table([[r['id'], r['kind_title'], r['name']] for r in rows],
                             ['ID', 'Журнал', 'Наименование']))
            elif op == 'r':
                d1 = input('С [первый день месяца]: ').strip()
                d2 = input('По [сегодня]: ').strip()
                qs = f'?d1={d1 or date.today().replace(day=1)}&d2={d2 or date.today()}'
                data = c.get('/api/v1/reports/by_period' + qs)
                print(_table([[r['kind'], r['count']] for r in data['rows']]
                             + [['ИТОГО', data['total']]],
                             ['Вид приказа', 'Кол-во']))
            elif op == 'w':
                rows = c.get('/api/v1/employees')
                print(_table([[r['full_name'], r['position'], r['role']]
                              for r in rows], ['ФИО', 'Должность', 'Роль']))
            elif op == 'x':
                import getpass
                old = getpass.getpass('Текущий пароль: ')
                new = getpass.getpass('Новый пароль: ')
                c.post('/api/v1/password', {'old_password': old, 'new_password': new,
                                            'confirm_password': new})
                print('Пароль изменён')
            else:
                print('Неизвестная команда, h — справка')
        except ApiError as e:
            print('Ошибка:', e)
        except (ValueError, IndexError):
            print('Некорректные параметры команды (h — справка)')
        except FileNotFoundError as e:
            print('Файл не найден:', e)


# ====================================================================== main
def main():
    ap = argparse.ArgumentParser(description='Клиент журнала приказов')
    ap.add_argument('--server', default='',
                    help='адрес сервера, напр. http://192.168.0.10:8080')
    ap.add_argument('--cli', action='store_true',
                    help='текстовый режим (без графического окна)')
    ap.add_argument('--login', default='', help='логин для режима --cli')
    args = ap.parse_args()

    if args.cli or not HAS_TK:
        if not HAS_TK and not args.cli:
            print('Tkinter недоступен — переключаюсь в текстовый режим '
                  '(--cli). Для графического интерфейса установите python3-tk.')
        base = args.server or input('Адрес сервера [http://localhost:8080]: ').strip() \
            or 'http://localhost:8080'
        login = args.login or input('Логин: ').strip()
        import getpass
        password = getpass.getpass('Пароль: ')
        try:
            cli_main(base, login, password)
        except KeyboardInterrupt:
            print('\nВыход')
        return

    LoginWindow, _, AppWindow = _build_gui_classes()
    win = LoginWindow(args.server)
    win.mainloop()
    if not win.result:
        return
    client, sess = win.result
    app = AppWindow(client, sess)
    app.mainloop()


if __name__ == '__main__':
    main()
