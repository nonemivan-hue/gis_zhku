# -*- coding: utf-8 -*-
"""ГИС ЖКУ - учёт лицевых счетов, организаций и адресов. Версия 1.7.1"""
import os
import re
import sqlite3
import shutil
import hashlib
import secrets
import logging
from datetime import datetime, timedelta

from flask import Flask, render_template, request, redirect, url_for, session, g, send_file, jsonify
import openpyxl
from docx import Document

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(APP_DIR, 'data')
DB_PATH = os.path.join(DATA_DIR, 'app.db')
BACKUP_DIR = os.path.join(DATA_DIR, 'backups')
TPL_OUT_DIR = os.path.join(APP_DIR, 'шаблоны')
APP_VERSION = '1.7.1'

for _d in (DATA_DIR, BACKUP_DIR, TPL_OUT_DIR):
    os.makedirs(_d, exist_ok=True)

logging.basicConfig(filename=os.path.join(DATA_DIR, 'app.log'), level=logging.INFO,
                    format='%(asctime)s %(levelname)s %(message)s')

app = Flask(__name__)
app.secret_key = secrets.token_hex(32)

ROLE_ACCESS = {
    'admin': {'*'},
    'operator': {'organizations', 'accounts', 'addresses', 'search', 'report', 'import'},
    'viewer': {'search', 'report'},
}
ROLE_TITLES = {'admin': 'Администратор', 'operator': 'Оператор', 'viewer': 'Просмотр'}

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    fullname TEXT DEFAULT '',
    role TEXT NOT NULL,
    password_hash TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS sessions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sid TEXT UNIQUE NOT NULL,
    username TEXT NOT NULL,
    role TEXT NOT NULL,
    fullname TEXT DEFAULT '',
    ip TEXT DEFAULT '',
    login_time TEXT NOT NULL,
    last_activity TEXT NOT NULL,
    logged_out INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS orgs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    ogrn TEXT DEFAULT '',
    kpp TEXT DEFAULT '',
    registry_id TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS accounts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    els TEXT NOT NULL,
    consumer_no TEXT DEFAULT '',
    jku_id TEXT DEFAULT '',
    gis_id TEXT DEFAULT '',
    acc_type TEXT DEFAULT '',
    acc_status TEXT DEFAULT '',
    date_created TEXT DEFAULT '',
    date_closed TEXT DEFAULT '',
    close_reason TEXT DEFAULT '',
    note TEXT DEFAULT '',
    is_tenant TEXT DEFAULT '',
    split_accounts TEXT DEFAULT '',
    total_area TEXT DEFAULT '',
    living_area TEXT DEFAULT '',
    heated_area TEXT DEFAULT '',
    residents TEXT DEFAULT '',
    org_ref TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS addresses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    els TEXT NOT NULL,
    address_full TEXT DEFAULT '',
    index_code TEXT DEFAULT '',
    region TEXT DEFAULT '',
    city TEXT DEFAULT '',
    settlement TEXT DEFAULT '',
    structure TEXT DEFAULT '',
    street TEXT DEFAULT '',
    house TEXT DEFAULT '',
    corpus TEXT DEFAULT '',
    apartment TEXT DEFAULT '',
    room TEXT DEFAULT '',
    fias_house TEXT DEFAULT '',
    house_type TEXT DEFAULT '',
    room_no TEXT DEFAULT '',
    gis_obj_id TEXT DEFAULT '',
    fund_status TEXT DEFAULT '',
    share TEXT DEFAULT '',
    house_area TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ts TEXT NOT NULL,
    username TEXT NOT NULL,
    action TEXT NOT NULL,
    details TEXT DEFAULT ''
);
"""

INDEXES = [
    "CREATE INDEX IF NOT EXISTS idx_accounts_els ON accounts(els)",
    "CREATE INDEX IF NOT EXISTS idx_accounts_consumer_no ON accounts(consumer_no)",
    "CREATE INDEX IF NOT EXISTS idx_accounts_org_ref ON accounts(org_ref)",
    "CREATE INDEX IF NOT EXISTS idx_addresses_els ON addresses(els)",
    "CREATE INDEX IF NOT EXISTS idx_addresses_city ON addresses(city)",
    "CREATE INDEX IF NOT EXISTS idx_addresses_settlement ON addresses(settlement)",
    "CREATE INDEX IF NOT EXISTS idx_addresses_structure ON addresses(structure)",
    "CREATE INDEX IF NOT EXISTS idx_addresses_street ON addresses(street)",
    "CREATE INDEX IF NOT EXISTS idx_addresses_house ON addresses(house)",
    "CREATE INDEX IF NOT EXISTS idx_addresses_apartment ON addresses(apartment)",
    "CREATE INDEX IF NOT EXISTS idx_orgs_registry_id ON orgs(registry_id)",
    "CREATE INDEX IF NOT EXISTS idx_sessions_sid ON sessions(sid)",
    "CREATE INDEX IF NOT EXISTS idx_audit_ts ON audit_log(ts)",
]


def get_db():
    if 'db' not in g:
        con = sqlite3.connect(DB_PATH, timeout=10)
        con.row_factory = sqlite3.Row
        con.execute('PRAGMA busy_timeout = 5000')
        g.db = con
    return g.db


@app.teardown_appcontext
def close_db(exc):
    db = g.pop('db', None)
    if db is not None:
        db.close()


SETTLEMENT_PREFIXES = ('с.', 'п.', 'пос.', 'кп.', 'мкр.', 'пгт.', 'ст-ца.', 'станица ')


def parse_address(full):
    res = {'index': '', 'region': '', 'city': '', 'settlement': '', 'structure': '',
           'street': '', 'house': '', 'corpus': '', 'apartment': '', 'room': ''}
    if not full:
        return res
    parts = [p.strip() for p in str(full).split(',') if p.strip()]
    for p in parts:
        low = p.lower()
        if re.fullmatch(r'\d{6}', p):
            res['index'] = p
            continue
        if low.endswith(' край') or low.endswith(' область') or low.endswith(' республика'):
            res['region'] = p
            continue
        m = re.match(r'^(жк|снт|днп|тсж|жск|кп|гск)[\s.]+(.+)$', p, re.I)
        if m:
            res['structure'] = p.strip()
            continue
        m = re.match(r'^ст[\s-]*ца[.\s]+(.+)$', low)
        if m:
            res['settlement'] = p[m.start(1):].strip()
            continue
        if low.startswith('станица '):
            res['settlement'] = p[8:].strip()
            continue
        matched = False
        for pref in SETTLEMENT_PREFIXES:
            if low.startswith(pref):
                res['settlement'] = p[len(pref):].strip()
                matched = True
                break
        if matched:
            continue
        m = re.match(r'^г[.\s]+(.+)$', p, re.I)
        if m:
            res['city'] = m.group(1).strip()
            continue
        if low.startswith('город '):
            res['city'] = p[6:].strip()
            continue
        m = re.match(r'^(ул|улица|пер|переулок|пр-кт|проспект|пр-д|проезд|ш|шоссе|б-р|бульвар)[.\s]+(.*)$', p, re.I)
        if m:
            res['street'] = m.group(2).strip() or p.strip()
            continue
        m = re.match(r'^(д|дом)[.\s]+(.*)$', p, re.I)
        if m:
            res['house'] = m.group(2).strip() or p.strip()
            continue
        m = re.match(r'^(корп|корпус)[.\s]+(.*)$', p, re.I)
        if m:
            res['corpus'] = m.group(2).strip() or p.strip()
            continue
        m = re.match(r'^(кв|квартира)[.\s]+(.*)$', p, re.I)
        if m:
            res['apartment'] = m.group(2).strip() or p.strip()
            continue
        m = re.match(r'^(ком|комн|комната)[.\s]+(.*)$', p, re.I)
        if m:
            res['room'] = m.group(2).strip() or p.strip()
            continue
        m = re.match(r'^к[.\s]+(.*)$', p, re.I)
        if m:
            if res['structure']:
                res['corpus'] = m.group(1).strip() or p.strip()
            else:
                res['room'] = m.group(1).strip() or p.strip()
            continue
        if re.fullmatch(r'\d+(/\d+)?[а-я]?', low):
            if not res['house']:
                res['house'] = p.strip()
            elif not res['corpus']:
                res['corpus'] = p.strip()
            continue
        if re.search(r'[а-яё]', low):
            if not res['city']:
                res['city'] = p.strip()
            elif not res['street']:
                res['street'] = re.sub(r'^(ул|улица|пер|переулок|пр-кт|проспект|пр-д|проезд|ш|шоссе|б-р|бульвар)[.\s]+', '', p, flags=re.I)
            continue
    return res


def norm(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip().lower()


UNIVERSAL_HEADERS = [
    '№ записи',
    'Номер ЛС (иной идентификатор потребителя)',
    'Идентификатор ЖКУ',
    'Единый лицевой счет',
    'Идентификатор лицевого счета в ГИС ЖКХ',
    'Тип лицевого счета',
    'Статус лицевого счета',
    'Дата создания',
    'Дата закрытия',
    'Причина закрытия',
    'Примечание (заполняется при проведении операции закрытия ЛС)',
    'Является нанимателем?',
    'Лицевые счета на помещение(я) разделены?',
    'Общая площадь, кв. м',
    'Жилая площадь, кв. м',
    'Отапливаемая площадь, кв. м',
    'Количество проживающих, чел.',
    'Наименование организации',
    'ОГРН организации',
    'КПП Организации',
    'Идентификатор организации в реестре организаций',
    'Адрес помещения',
    'Глобальный уникальный идентификатор дома по ФИАС/Идентификационный код дома в ГИС ЖКХ  Адрес дома, помещения, компнаты',
    'Тип помещения/блок Адрес дома, помещения, комнаты',
    'Номер помещения/блока Адрес дома, помещения, комнаты',
    'Номер комнаты',
    'Идентификатор дома, помещения, блока, комнаты, присвоенный ГИС ЖКХ',
    'Статус объекта жилищного фонда',
    'Доля внесения платы, размер доли в %',
    'Общая площадь дома, помещения, комнаты, кв. м',
]

UNIVERSAL_RULES = {
    'rec_no': ['№ записи'],
    'consumer_no': ['Номер ЛС (иной идентификатор потребителя)'],
    'jku_id': ['Идентификатор ЖКУ'],
    'els': ['Единый лицевой счет'],
    'gis_id': ['Идентификатор лицевого счета в ГИС ЖКХ'],
    'acc_type': ['Тип лицевого счета'],
    'acc_status': ['Статус лицевого счета'],
    'date_created': ['Дата создания'],
    'date_closed': ['Дата закрытия'],
    'close_reason': ['Причина закрытия'],
    'note': ['Примечание (заполняется при проведении операции закрытия ЛС)'],
    'is_tenant': ['Является нанимателем?'],
    'split_accounts': ['Лицевые счета на помещение(я) разделены?'],
    'total_area': ['Общая площадь, кв. м'],
    'living_area': ['Жилая площадь, кв. м'],
    'heated_area': ['Отапливаемая площадь, кв. м'],
    'residents': ['Количество проживающих, чел.'],
    'org_name': ['Наименование организации'],
    'org_ogrn': ['ОГРН организации'],
    'org_kpp': ['КПП Организации'],
    'org_ref': ['Идентификатор организации в реестре организаций'],
    'address_full': ['Адрес помещения'],
    'fias_house': ['Глобальный уникальный идентификатор дома по ФИАС/Идентификационный код дома в ГИС ЖКХ  Адрес дома, помещения, компнаты'],
    'house_type': ['Тип помещения/блок Адрес дома, помещения, комнаты'],
    'room_no': ['Номер помещения/блока Адрес дома, помещения, комнаты'],
    'room': ['Номер комнаты'],
    'gis_obj_id': ['Идентификатор дома, помещения, блока, комнаты, присвоенный ГИС ЖКХ'],
    'fund_status': ['Статус объекта жилищного фонда'],
    'share': ['Доля внесения платы, размер доли в %'],
    'house_area': ['Общая площадь дома, помещения, комнаты, кв. м'],
}

ORG_RULES = {
    'name': ['Наименование организации', 'Наименование'],
    'ogrn': ['ОГРН организации', 'ОГРН'],
    'kpp': ['КПП Организации', 'КПП'],
    'registry_id': ['Идентификатор организации в реестре организаций', 'Идентификатор в реестре'],
}

ACC_RULES = {
    'els': ['Единый лицевой счет', 'ЕЛС', 'Лицевой счет'],
    'consumer_no': ['Номер ЛС (иной идентификатор потребителя)', 'Номер ЛС', 'Номер лицевого счета'],
    'jku_id': ['Идентификатор ЖКУ'],
    'gis_id': ['Идентификатор лицевого счета в ГИС ЖКХ'],
    'acc_type': ['Тип лицевого счета'],
    'acc_status': ['Статус лицевого счета'],
    'date_created': ['Дата создания'],
    'date_closed': ['Дата закрытия'],
    'close_reason': ['Причина закрытия'],
    'note': ['Примечание (заполняется при проведении операции закрытия ЛС)', 'Примечание'],
    'is_tenant': ['Является нанимателем?'],
    'split_accounts': ['Лицевые счета на помещение(я) разделены?'],
    'total_area': ['Общая площадь, кв. м', 'Общая площадь'],
    'living_area': ['Жилая площадь, кв. м', 'Жилая площадь'],
    'heated_area': ['Отапливаемая площадь, кв. м', 'Отапливаемая площадь'],
    'residents': ['Количество проживающих, чел.', 'Количество проживающих'],
    'org_ref': ['Идентификатор организации в реестре организаций', 'Идентификатор организации'],
}

ADDR_RULES = {
    'els': ['Единый лицевой счет', 'ЕЛС'],
    'address_full': ['Адрес помещения', 'Адрес'],
    'fias_house': ['Глобальный уникальный идентификатор дома по ФИАС/Идентификационный код дома в ГИС ЖКХ  Адрес дома, помещения, компнаты', 'ГИД по ФИАС'],
    'house_type': ['Тип помещения/блок Адрес дома, помещения, комнаты', 'Тип помещения'],
    'room_no': ['Номер помещения/блока Адрес дома, помещения, комнаты', 'Номер помещения'],
    'room': ['Номер комнаты'],
    'gis_obj_id': ['Идентификатор дома, помещения, блока, комнаты, присвоенный ГИС ЖКХ', 'Идентификатор в ГИС ЖКХ'],
    'fund_status': ['Статус объекта жилищного фонда'],
    'share': ['Доля внесения платы, размер доли в %', 'Доля внесения платы'],
    'house_area': ['Общая площадь дома, помещения, комнаты, кв. м', 'Общая площадь дома'],
}


def match_columns(headers, rules):
    result = {}
    hmap = {}
    for i, h in enumerate(headers):
        if h is None:
            continue
        n = norm(h)
        if n and n not in hmap:
            hmap[n] = i
    for key, variants in rules.items():
        for v in variants:
            idx = hmap.get(norm(v))
            if idx is not None:
                result[key] = idx
                break
    return result

def init_db():
    con = sqlite3.connect(DB_PATH, timeout=10)
    con.row_factory = sqlite3.Row
    con.executescript(SCHEMA)
    for idx in INDEXES:
        con.execute(idx)
    con.commit()
    con.close()


def hash_password(pw, salt=None):
    salt = salt or secrets.token_hex(8)
    h = hashlib.sha256((salt + pw).encode('utf-8')).hexdigest()
    return salt + '$' + h


def check_password(pw, stored):
    if not stored or '$' not in stored:
        return False
    salt, _ = stored.split('$', 1)
    return hash_password(pw, salt) == stored


def seed_if_empty():
    con = sqlite3.connect(DB_PATH, timeout=10)
    con.row_factory = sqlite3.Row
    n = con.execute('SELECT COUNT(*) AS c FROM users').fetchone()['c']
    if n == 0:
        con.execute('INSERT INTO users (username, fullname, role, password_hash) VALUES (?,?,?,?)',
                    ('admin', 'Администратор', 'admin', hash_password('admin123')))
        con.execute('INSERT INTO users (username, fullname, role, password_hash) VALUES (?,?,?,?)',
                    ('operator', 'Оператор', 'operator', hash_password('operator123')))
        con.execute('INSERT INTO users (username, fullname, role, password_hash) VALUES (?,?,?,?)',
                    ('viewer', 'Просмотр', 'viewer', hash_password('viewer123')))
    con.commit()
    con.close()


def audit(con, username, action, details=''):
    con.execute('INSERT INTO audit_log (ts, username, action, details) VALUES (?,?,?,?)',
                (datetime.now().strftime('%Y-%m-%d %H:%M:%S'), username or 'аноним', action, details or ''))
    con.commit()


def ensure_org(con, name, ogrn, kpp, registry_id):
    registry_id = (registry_id or '').strip()
    name = (name or '').strip()
    ogrn = (ogrn or '').strip()
    kpp = (kpp or '').strip()
    existing = None
    if registry_id:
        existing = con.execute('SELECT * FROM orgs WHERE registry_id = ?', (registry_id,)).fetchone()
    if existing is None and ogrn:
        existing = con.execute('SELECT * FROM orgs WHERE ogrn = ?', (ogrn,)).fetchone()
    if existing is None and name:
        existing = con.execute('SELECT * FROM orgs WHERE UPPER(name) = UPPER(?)', (name,)).fetchone()
    if existing is not None:
        con.execute('UPDATE orgs SET name=?, ogrn=?, kpp=?, registry_id=? WHERE id=?',
                    (name or existing['name'], ogrn or existing['ogrn'],
                     kpp or existing['kpp'], registry_id or existing['registry_id'], existing['id']))
        return existing['id'], False
    cur = con.execute('INSERT INTO orgs (name, ogrn, kpp, registry_id) VALUES (?,?,?,?)',
                      (name, ogrn, kpp, registry_id))
    return cur.lastrowid, True


def ensure_account(con, acc):
    els = (acc.get('els') or '').strip()
    row = con.execute('SELECT * FROM accounts WHERE els = ?', (els,)).fetchone()
    if row is not None:
        con.execute('UPDATE accounts SET consumer_no=?, jku_id=?, gis_id=?, acc_type=?, acc_status=?, date_created=?, date_closed=?, close_reason=?, note=?, is_tenant=?, split_accounts=?, total_area=?, living_area=?, heated_area=?, residents=?, org_ref=? WHERE els=?',
                    (acc.get('consumer_no', ''), acc.get('jku_id', ''), acc.get('gis_id', ''),
                     acc.get('acc_type', ''), acc.get('acc_status', ''), acc.get('date_created', ''),
                     acc.get('date_closed', ''), acc.get('close_reason', ''), acc.get('note', ''),
                     acc.get('is_tenant', ''), acc.get('split_accounts', ''), acc.get('total_area', ''),
                     acc.get('living_area', ''), acc.get('heated_area', ''), acc.get('residents', ''),
                     acc.get('org_ref', ''), els))
        return row['id'], False
    cur = con.execute('INSERT INTO accounts (els, consumer_no, jku_id, gis_id, acc_type, acc_status, date_created, date_closed, close_reason, note, is_tenant, split_accounts, total_area, living_area, heated_area, residents, org_ref) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                      (els, acc.get('consumer_no', ''), acc.get('jku_id', ''), acc.get('gis_id', ''),
                       acc.get('acc_type', ''), acc.get('acc_status', ''), acc.get('date_created', ''),
                       acc.get('date_closed', ''), acc.get('close_reason', ''), acc.get('note', ''),
                       acc.get('is_tenant', ''), acc.get('split_accounts', ''), acc.get('total_area', ''),
                       acc.get('living_area', ''), acc.get('heated_area', ''), acc.get('residents', ''),
                       acc.get('org_ref', '')))
    return cur.lastrowid, True


def ensure_address(con, addr):
    els = (addr.get('els') or '').strip()
    row = con.execute('SELECT id FROM addresses WHERE els = ?', (els,)).fetchone()
    vals = (addr.get('address_full', ''), addr.get('index', ''), addr.get('region', ''),
            addr.get('city', ''), addr.get('settlement', ''), addr.get('structure', ''),
            addr.get('street', ''), addr.get('house', ''), addr.get('corpus', ''),
            addr.get('apartment', ''), addr.get('room', ''), addr.get('fias_house', ''),
            addr.get('house_type', ''), addr.get('room_no', ''), addr.get('gis_obj_id', ''),
            addr.get('fund_status', ''), addr.get('share', ''), addr.get('house_area', ''))
    if row is not None:
        con.execute('UPDATE addresses SET address_full=?, index_code=?, region=?, city=?, settlement=?, structure=?, street=?, house=?, corpus=?, apartment=?, room=?, fias_house=?, house_type=?, room_no=?, gis_obj_id=?, fund_status=?, share=?, house_area=? WHERE els=?',
                    vals + (els,))
        return row['id'], False
    cur = con.execute('INSERT INTO addresses (els, address_full, index_code, region, city, settlement, structure, street, house, corpus, apartment, room, fias_house, house_type, room_no, gis_obj_id, fund_status, share, house_area) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                      (els,) + vals)
    return cur.lastrowid, True


def universal_import(path, mode='add'):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    stats = {'accounts_new': 0, 'accounts_upd': 0, 'orgs_new': 0, 'orgs_upd': 0,
             'addresses_new': 0, 'addresses_upd': 0, 'skipped': 0, 'errors': []}
    if len(rows) < 2:
        stats['errors'].append('Файл пуст или не содержит данных')
        return stats
    headers = [str(h) if h is not None else '' for h in rows[0]]
    cols = match_columns(headers, UNIVERSAL_RULES)
    if 'els' not in cols:
        stats['errors'].append('Не найдена колонка «Единый лицевой счет»')
        return stats
    con = sqlite3.connect(DB_PATH, timeout=10)
    con.row_factory = sqlite3.Row
    if mode == 'replace':
        con.execute('DELETE FROM accounts')
        con.execute('DELETE FROM addresses')
        con.execute('DELETE FROM orgs')
    for row in rows[1:]:
        def val(key):
            ci = cols.get(key)
            if ci is None or ci >= len(row):
                return ''
            v = row[ci]
            return '' if v is None else str(v).strip()
        els = val('els')
        if not els:
            if any(val(k) for k in ('consumer_no', 'org_name', 'address_full')):
                stats['skipped'] += 1
            continue
        if val('org_ref') or val('org_name'):
            _oid, created = ensure_org(con, val('org_name'), val('org_ogrn'), val('org_kpp'), val('org_ref'))
            if created:
                stats['orgs_new'] += 1
            else:
                stats['orgs_upd'] += 1
        acc = {'els': els, 'consumer_no': val('consumer_no'), 'jku_id': val('jku_id'),
               'gis_id': val('gis_id'), 'acc_type': val('acc_type'), 'acc_status': val('acc_status'),
               'date_created': val('date_created'), 'date_closed': val('date_closed'),
               'close_reason': val('close_reason'), 'note': val('note'),
               'is_tenant': val('is_tenant'), 'split_accounts': val('split_accounts'),
               'total_area': val('total_area'), 'living_area': val('living_area'),
               'heated_area': val('heated_area'), 'residents': val('residents'),
               'org_ref': val('org_ref')}
        _aid, created = ensure_account(con, acc)
        if created:
            stats['accounts_new'] += 1
        else:
            stats['accounts_upd'] += 1
        addr = {'els': els, 'address_full': val('address_full')}
        addr.update(parse_address(val('address_full')))
        addr['fias_house'] = val('fias_house')
        addr['house_type'] = val('house_type')
        addr['room_no'] = val('room_no')
        addr['gis_obj_id'] = val('gis_obj_id')
        addr['fund_status'] = val('fund_status')
        addr['share'] = val('share')
        addr['house_area'] = val('house_area')
        _adid, created = ensure_address(con, addr)
        if created:
            stats['addresses_new'] += 1
        else:
            stats['addresses_upd'] += 1
    con.commit()
    con.close()
    return stats


def section_import(path, table, rules, mode='add'):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    stats = {'new': 0, 'upd': 0, 'skipped': 0, 'errors': []}
    if len(rows) < 2:
        stats['errors'].append('Файл пуст или не содержит данных')
        return stats
    headers = [str(h) if h is not None else '' for h in rows[0]]
    cols = match_columns(headers, rules)
    key = 'registry_id' if table == 'orgs' else 'els'
    if key not in cols:
        stats['errors'].append('Не найдена ключовая колонка (идентификатор)')
        return stats
    con = sqlite3.connect(DB_PATH, timeout=10)
    con.row_factory = sqlite3.Row
    if mode == 'replace':
        con.execute('DELETE FROM ' + table)
    fields = list(cols.keys())
    for row in rows[1:]:
        rec = {}
        for f, ci in cols.items():
            v = row[ci] if ci < len(row) else None
            rec[f] = '' if v is None else str(v).strip()
        keyval = rec.get(key, '')
        if not keyval:
            stats['skipped'] += 1
            continue
        existing = con.execute('SELECT id FROM ' + table + ' WHERE ' + key + ' = ?', (keyval,)).fetchone()
        if existing:
            sets = ', '.join(f + ' = ?' for f in fields if f != key)
            args = [rec[f] for f in fields if f != key] + [existing['id']]
            con.execute('UPDATE ' + table + ' SET ' + sets + ' WHERE id = ?', args)
            stats['upd'] += 1
        else:
            names = ', '.join(fields)
            q = ', '.join('?' for _ in fields)
            con.execute('INSERT INTO ' + table + ' (' + names + ') VALUES (' + q + ')',
                        [rec[f] for f in fields])
            stats['new'] += 1
    con.commit()
    con.close()
    return stats


def report_rows(els=None, consumer_no=None):
    con = sqlite3.connect(DB_PATH, timeout=10)
    con.row_factory = sqlite3.Row
    sql = ("SELECT a.els, a.consumer_no, a.jku_id, a.gis_id, a.acc_type, a.acc_status, "
           "a.date_created, a.date_closed, a.total_area, a.living_area, a.heated_area, "
           "a.residents, a.org_ref, o.name AS org_name, o.ogrn, o.kpp, "
           "ad.address_full, ad.city, ad.settlement, ad.street, ad.house, ad.corpus, "
           "ad.apartment, ad.house_area "
           "FROM accounts a "
           "LEFT JOIN addresses ad ON ad.els = a.els "
           "LEFT JOIN orgs o ON o.registry_id = a.org_ref ")
    cond = []
    args = []
    if els:
        cond.append('a.els = ?')
        args.append(els.strip())
    if consumer_no:
        cond.append('a.consumer_no = ?')
        args.append(consumer_no.strip())
    if cond:
        sql += 'WHERE ' + ' AND '.join(cond)
    sql += ' GROUP BY a.els, a.consumer_no, a.gis_id'
    rows = con.execute(sql, args).fetchall()
    con.close()
    return [dict(r) for r in rows]


def export_word(rows):
    doc = Document()
    doc.add_heading('Отчет по лицевым счетам', level=1)
    doc.add_paragraph('Дата формирования: ' + datetime.now().strftime('%d.%m.%Y %H:%M'))
    for r in rows:
        doc.add_heading('ЕЛС ' + str(r.get('els') or ''), level=2)
        t = doc.add_table(rows=0, cols=2)
        t.style = 'Table Grid'
        pairs = [
            ('Номер ЛС', r.get('consumer_no', '')),
            ('Идентификатор ЖКУ', r.get('jku_id', '')),
            ('Тип лицевого счета', r.get('acc_type', '')),
            ('Статус', r.get('acc_status', '')),
            ('Дата создания', r.get('date_created', '')),
            ('Организация', r.get('org_name', '')),
            ('ОГРН организации', r.get('ogrn', '')),
            ('КПП организации', r.get('kpp', '')),
            ('Адрес помещения', r.get('address_full', '')),
            ('Общая площадь дома, кв. м', r.get('house_area', '')),
            ('Общая площадь, кв. м', r.get('total_area', '')),
            ('Жилая площадь, кв. м', r.get('living_area', '')),
            ('Отапливаемая площадь, кв. м', r.get('heated_area', '')),
            ('Проживает, чел.', r.get('residents', '')),
        ]
        for k, v in pairs:
            cells = t.add_row().cells
            cells[0].text = k
            cells[1].text = str(v or '')
    path = os.path.join(DATA_DIR, 'report_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '.docx')
    doc.save(path)
    return path


TEMPLATE_SPECS = {
    'organizations': {
        'file': 'Шаблон_Организации.xlsx',
        'headers': ['Наименование организации', 'ОГРН организации', 'КПП Организации',
                    'Идентификатор организации в реестре организаций'],
        'example': ['МУНИЦИПАЛЬНОЕ УНИТАРНОЕ ПРЕДПРИЯТИЕ МУНИЦИПАЛЬНОГО ОБРАЗОВАНИЯ ГОРОД-КУРОРТ ГЕЛЕНДЖИК "ТЕПЛОВЫЕ СЕТИ"',
                    '1022300769688', '230401001', '4c19d941-b8ce-4f9d-8d7e-7ca0a42ac173'],
        'notes': ['Полное наименование организации', '13 цифр без пробелов', '9 цифр',
                  'Глобальный идентификатор из ГИС ЖКХ; единый лицевой счет в этом справочнике не указывается'],
    },
    'accounts': {
        'file': 'Шаблон_Счета.xlsx',
        'headers': ['Единый лицевой счет', 'Номер ЛС (иной идентификатор потребителя)', 'Идентификатор ЖКУ',
                    'Идентификатор лицевого счета в ГИС ЖКХ', 'Тип лицевого счета', 'Статус лицевого счета',
                    'Дата создания', 'Дата закрытия', 'Причина закрытия',
                    'Примечание (заполняется при проведении операции закрытия ЛС)', 'Является нанимателем?',
                    'Лицевые счета на помещение(я) разделены?', 'Общая площадь, кв. м',
                    'Жилая площадь, кв. м', 'Отапливаемая площадь, кв. м', 'Количество проживающих, чел.',
                    'Идентификатор организации в реестре организаций'],
        'example': ['40КР079116', '210260530', '40КР079116-68', '0000181f-5f18-4f65-b437-d0ea2391f684',
                    'ЛС ДРСО', 'Активный', '17.08.2022', '', '', '', 'нет', 'нет', '75.8', '40.9', '75.8', '1',
                    '4c19d941-b8ce-4f9d-8d7e-7ca0a42ac173'],
        'notes': ['Ключ счета — обязателен', 'Иной идентификатор потребителя', 'Идентификатор начислений ЖКУ',
                  'Глобальный идентификатор счета', 'ЛС ДРСО / ЛС УО', 'Активный или Закрыт',
                  'ДД.ММ.ГГГГ', 'не заполняется для активных', 'заполняется при закрытии',
                  'свободный текст', 'да/нет', 'да/нет', 'число', 'число', 'число', 'целое число',
                  'идентификатор организации из справочника организаций'],
    },
    'addresses': {
        'file': 'Шаблон_Адреса.xlsx',
        'headers': ['Единый лицевой счет', 'Адрес помещения',
                    'Глобальный уникальный идентификатор дома по ФИАС/Идентификационный код дома в ГИС ЖКХ  Адрес дома, помещения, компнаты',
                    'Тип помещения/блок Адрес дома, помещения, комнаты',
                    'Номер помещения/блока Адрес дома, помещения, комнаты', 'Номер комнаты',
                    'Идентификатор дома, помещения, блока, комнаты, присвоенный ГИС ЖКХ',
                    'Статус объекта жилищного фонда', 'Доля внесения платы, размер доли в %',
                    'Общая площадь дома, помещения, комнаты, кв. м'],
        'example': ['40КР079116', '353488, Краснодарский край, г. Геленджик, с. Архипо-Осиповка, ул. Рабочая, д. 1, кв. 3',
                    'b8a8a46a-5cac-4cfc-a189-4856a858a072', 'Жилое помещение', '3', '', '9CxC0005300003',
                    '1', '100', '400.8'],
        'notes': ['Ключ — единый лицевой счет', 'Полный адрес одной строкой: индекс, край, город, населенный пункт, улица, дом, корпус, квартира',
                  'ГИД дома из ФИАС/ГИС ЖКХ', 'Жилое помещение / блок и т.п.', 'номер помещения',
                  'если есть', 'идентификатор в ГИС ЖКХ', '0 или 1', 'число', 'число'],
    },
    'universal': {
        'file': 'Шаблон_Универсальный.xlsx',
        'headers': UNIVERSAL_HEADERS,
        'example': ['1', '210260530', '40КР079116-68', '40КР079116', '0000181f-5f18-4f65-b437-d0ea2391f684',
                    'ЛС ДРСО', 'Активный', '17.08.2022', '', '', '', 'нет', 'нет', '', '', '', '',
                    'МУП "ТЕПЛОВЫЕ СЕТИ"', '1022300769688', '230401001', '4c19d941-b8ce-4f9d-8d7e-7ca0a42ac173',
                    '353488, Краснодарский край, г. Геленджик, с. Архипо-Осиповка, ул. Рабочая, д. 1, кв. 3',
                    'b8a8a46a-5cac-4cfc-a189-4856a858a072', 'Жилое помещение', '3', '', '9CxC0005300003',
                    '1', '', '400.8'],
        'notes': ['Порядковый номер строки', 'Номер ЛС потребителя', 'Идентификатор ЖКУ',
                  'Ключ счета — обязателен', 'ГИД счета в ГИС ЖКХ', 'ЛС ДРСО / ЛС УО', 'Активный или Закрыт',
                  'ДД.ММ.ГГГГ', 'не заполняется для активных', 'заполняется при закрытии', 'свободный текст',
                  'да/нет', 'да/нет', 'число', 'число', 'число', 'целое число',
                  'организация загружается из этой колонки', 'ОГРН', 'КПП',
                  'идентификатор организации — по нему схлопываются дубли',
                  'адрес разбирается автоматически на уровни', 'ГИД дома', 'тип помещения',
                  'номер помещения', 'номер комнаты', 'идентификатор в ГИС ЖКХ', 'статус',
                  'доля платы', 'площадь дома'],
    },
}


def build_template_workbook(section):
    spec = TEMPLATE_SPECS[section]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Шаблон'
    for i, h in enumerate(spec['headers'], start=1):
        ws.cell(row=1, column=i, value=h)
    for i, e in enumerate(spec['example'], start=1):
        ws.cell(row=2, column=i, value=e)
    for col in range(1, len(spec['headers']) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 28
    ws2 = wb.create_sheet('Инструкция')
    ws2.cell(row=1, column=1, value='Колонка')
    ws2.cell(row=1, column=2, value='Пример заполнения')
    ws2.cell(row=1, column=3, value='Пояснение')
    for i, h in enumerate(spec['headers'], start=1):
        ws2.cell(row=i + 1, column=1, value=h)
        ws2.cell(row=i + 1, column=2, value=spec['example'][i - 1] if i - 1 < len(spec['example']) else '')
        ws2.cell(row=i + 1, column=3, value=spec['notes'][i - 1] if i - 1 < len(spec['notes']) else '')
    ws2.column_dimensions['A'].width = 45
    ws2.column_dimensions['B'].width = 45
    ws2.column_dimensions['C'].width = 60
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


import io

def backup_db():
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    dst = os.path.join(BACKUP_DIR, 'app_' + stamp + '.db')
    shutil.copy2(DB_PATH, dst)
    return dst


def restore_db(src_path):
    if not src_path or not os.path.isfile(src_path):
        return False, 'Файл не найден'
    try:
        chk = sqlite3.connect(src_path)
        chk.execute('SELECT COUNT(*) FROM users')
        chk.close()
    except Exception:
        return False, 'Файл не является базой данных приложения'
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    pre = os.path.join(BACKUP_DIR, 'pre_restore_' + stamp + '.db')
    shutil.copy2(DB_PATH, pre)
    shutil.copy2(src_path, DB_PATH)
    return True, pre

def save_template_files():
    for section, spec in TEMPLATE_SPECS.items():
        bio = build_template_workbook(section)
        with open(os.path.join(TPL_OUT_DIR, spec['file']), 'wb') as f:
            f.write(bio.getvalue())


def active_users(con):
    limit = (datetime.now() - timedelta(minutes=15)).strftime('%Y-%m-%d %H:%M:%S')
    rows = con.execute("SELECT username, role, fullname, ip, login_time, last_activity FROM sessions WHERE logged_out = 0 AND last_activity >= ? ORDER BY last_activity DESC", (limit,)).fetchall()
    return [dict(r) for r in rows]


def current_user():
    return session.get('user')


def can(section):
    u = session.get('user')
    if not u:
        return False
    acc = ROLE_ACCESS.get(u.get('role'), set())
    return '*' in acc or section in acc


@app.before_request
def before_request():
    sid = session.get('sid')
    if sid:
        try:
            con = get_db()
            con.execute('UPDATE sessions SET last_activity = ? WHERE sid = ? AND logged_out = 0',
                        (datetime.now().strftime('%Y-%m-%d %H:%M:%S'), sid))
            con.commit()
        except sqlite3.Error:
            pass


@app.context_processor
def inject_globals():
    return {'can': can, 'current_user': current_user, 'APP_VERSION': APP_VERSION,
            'ROLE_TITLES': ROLE_TITLES}


@app.route('/login', methods=['GET', 'POST'])
def login():
    error = ''
    if request.method == 'POST':
        username = (request.form.get('username') or '').strip()
        password = request.form.get('password') or ''
        con = get_db()
        row = con.execute('SELECT * FROM users WHERE username = ?', (username,)).fetchone()
        if row and check_password(password, row['password_hash']):
            sid = secrets.token_hex(16)
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ip = (request.headers.get('X-Forwarded-For') or request.remote_addr or '')[:64]
            con.execute('INSERT INTO sessions (sid, username, role, fullname, ip, login_time, last_activity) VALUES (?,?,?,?,?,?,?)',
                        (sid, row['username'], row['role'], row['fullname'], ip, now, now))
            con.commit()
            session['sid'] = sid
            session['user'] = {'username': row['username'], 'role': row['role'], 'fullname': row['fullname']}
            audit(con, row['username'], 'Вход в систему')
            return redirect(url_for('index'))
        error = 'Неверный логин или пароль'
    return render_template('login.html', error=error)


@app.route('/logout')
def logout():
    sid = session.get('sid')
    if sid:
        con = get_db()
        con.execute('UPDATE sessions SET logged_out = 1 WHERE sid = ?', (sid,))
        con.commit()
    session.pop('sid', None)
    session.pop('user', None)
    return redirect(url_for('login'))


@app.route('/')
def index():
    if not can('search'):
        return redirect(url_for('login'))
    con = get_db()
    stats = {
        'orgs': con.execute('SELECT COUNT(*) AS c FROM orgs').fetchone()['c'],
        'accounts': con.execute('SELECT COUNT(*) AS c FROM accounts').fetchone()['c'],
        'addresses': con.execute('SELECT COUNT(*) AS c FROM addresses').fetchone()['c'],
        'users': con.execute('SELECT COUNT(*) AS c FROM users').fetchone()['c'],
    }
    users = active_users(con) if can('sessions') else []
    return render_template('index.html', stats=stats, active_users=users)


@app.route('/api/sessions')
def api_sessions():
    if not can('sessions'):
        return jsonify({'ok': False, 'users': []})
    con = get_db()
    return jsonify({'ok': True, 'users': active_users(con)})


@app.route('/api/cascade')
def api_cascade():
    if not can('search'):
        return jsonify({'ok': False, 'items': []})
    level = request.args.get('level', '')
    city = request.args.get('city', '')
    settlement = request.args.get('settlement', '')
    structure = request.args.get('structure', '')
    street = request.args.get('street', '')
    con = get_db()
    out = []
    if level == 'cities':
        rows = con.execute("SELECT DISTINCT city FROM addresses WHERE COALESCE(city, '') <> '' ORDER BY city").fetchall()
        out = [r[0] for r in rows]
    elif level == 'settlements':
        rows = con.execute("SELECT DISTINCT settlement FROM addresses WHERE city = ? AND COALESCE(settlement, '') <> '' ORDER BY settlement", (city,)).fetchall()
        out = [r[0] for r in rows]
    elif level == 'structures':
        rows = con.execute("SELECT DISTINCT structure FROM addresses WHERE city = ? AND settlement = ? AND COALESCE(structure, '') <> '' ORDER BY structure", (city, settlement)).fetchall()
        out = [r[0] for r in rows]
    elif level == 'streets':
        sql = "SELECT DISTINCT street FROM addresses WHERE city = ? AND COALESCE(street, '') <> ''"
        args = [city]
        if settlement:
            sql += ' AND settlement = ?'
            args.append(settlement)
        if structure:
            sql += ' AND structure = ?'
            args.append(structure)
        sql += ' ORDER BY street'
        rows = con.execute(sql, args).fetchall()
        out = [r[0] for r in rows]
    elif level == 'buildings':
        sql = "SELECT DISTINCT house, corpus FROM addresses WHERE city = ? AND street = ?"
        args = [city, street]
        if settlement:
            sql += ' AND settlement = ?'
            args.append(settlement)
        sql += ' ORDER BY CAST(house AS INTEGER), corpus'
        rows = con.execute(sql, args).fetchall()
        out = [{'house': r[0], 'corpus': r[1]} for r in rows]
    elif level == 'accounts':
        sql = "SELECT a.els, a.consumer_no, ad.house, ad.corpus, ad.apartment FROM accounts a JOIN addresses ad ON ad.els = a.els WHERE ad.city = ? AND ad.street = ? AND ad.house = ?"
        args = [city, street, request.args.get('house', '')]
        rows = con.execute(sql, args).fetchall()
        out = [dict(r) for r in rows]
    return jsonify({'ok': True, 'items': out})


@app.route('/search')
def search():
    if not can('search'):
        return redirect(url_for('login'))
    return render_template('search.html')


@app.route('/report')
def report():
    if not can('report'):
        return redirect(url_for('login'))
    els = (request.args.get('els') or '').strip()
    consumer_no = (request.args.get('consumer_no') or '').strip()
    rows = []
    if els or consumer_no:
        rows = report_rows(els or None, consumer_no or None)
    return render_template('report.html', rows=rows, els=els, consumer_no=consumer_no)


@app.route('/report/word')
def report_word():
    if not can('report'):
        return redirect(url_for('login'))
    els = (request.args.get('els') or '').strip()
    consumer_no = (request.args.get('consumer_no') or '').strip()
    rows = report_rows(els or None, consumer_no or None)
    if not rows:
        return redirect(url_for('report', els=els, consumer_no=consumer_no))
    path = export_word(rows)
    con = get_db()
    audit(con, (current_user() or {}).get('username', ''), 'Экспорт отчета в Word', 'ЕЛС: ' + els)
    return send_file(path, as_attachment=True, download_name='report.docx')


DIR_SECTIONS = {
    'organizations': {
        'title': 'Организации',
        'table': 'orgs',
        'columns': [('name', 'Наименование'), ('ogrn', 'ОГРН'), ('kpp', 'КПП'),
                    ('registry_id', 'Идентификатор в реестре')],
        'rules': ORG_RULES,
        'key': 'registry_id',
        'required': 'name',
    },
    'accounts': {
        'title': 'Лицевые счета',
        'table': 'accounts',
        'columns': [('els', 'Единый лицевой счет'), ('consumer_no', 'Номер ЛС'),
                    ('jku_id', 'Идентификатор ЖКУ'), ('gis_id', 'Идентификатор в ГИС ЖКХ'),
                    ('acc_type', 'Тип'), ('acc_status', 'Статус'), ('date_created', 'Дата создания'),
                    ('total_area', 'Общая площадь, кв. м'), ('living_area', 'Жилая площадь, кв. м'),
                    ('heated_area', 'Отапливаемая площадь, кв. м'), ('residents', 'Проживает, чел.'),
                    ('org_ref', 'Организация (идентификатор в реестре)')],
        'rules': ACC_RULES,
        'key': 'els',
        'required': 'els',
    },
    'addresses': {
        'title': 'Адреса',
        'table': 'addresses',
        'columns': [('els', 'Единый лицевой счет'), ('address_full', 'Адрес'), ('city', 'Город'),
                    ('settlement', 'Населенный пункт'), ('street', 'Улица'), ('house', 'Дом'),
                    ('corpus', 'Корпус'), ('apartment', 'Квартира'), ('fias_house', 'ГИД ФИАС'),
                    ('house_area', 'Общая площадь дома, кв. м')],
        'rules': ADDR_RULES,
        'key': 'els',
        'required': 'els',
    },
}


@app.route('/<sec>', methods=['GET', 'POST'])
def ref_list(sec):
    if sec not in DIR_SECTIONS:
        return redirect(url_for('index'))
    if not can(sec):
        return redirect(url_for('login'))
    meta = DIR_SECTIONS[sec]
    con = get_db()
    if request.method == 'POST':
        action = request.form.get('action', '')
        if action == 'delete':
            rid = request.form.get('id', '')
            con.execute('DELETE FROM ' + meta['table'] + ' WHERE id = ?', (rid,))
            audit(con, (current_user() or {}).get('username', ''), 'Удаление записи', meta['title'] + ' id=' + rid)
            con.commit()
            return redirect(url_for('ref_list', sec=sec))
        rid = request.form.get('id', '')
        values = {}
        for key, _label in meta['columns']:
            values[key] = (request.form.get(key) or '').strip()
        if not values.get(meta['required']):
            return redirect(url_for('ref_list', sec=sec))
        if action == 'save' and rid:
            sets = ', '.join(k + ' = ?' for k in values)
            con.execute('UPDATE ' + meta['table'] + ' SET ' + sets + ' WHERE id = ?', list(values.values()) + [rid])
            audit(con, (current_user() or {}).get('username', ''), 'Изменение записи', meta['title'] + ' id=' + rid)
        else:
            names = ', '.join(values.keys())
            q = ', '.join('?' for _ in values)
            con.execute('INSERT INTO ' + meta['table'] + ' (' + names + ') VALUES (' + q + ')', list(values.values()))
            audit(con, (current_user() or {}).get('username', ''), 'Добавление записи', meta['title'])
        con.commit()
        return redirect(url_for('ref_list', sec=sec))
    rows = [dict(r) for r in con.execute('SELECT * FROM ' + meta['table'] + ' ORDER BY id LIMIT 500').fetchall()]
    org_map = {}
    if sec == 'accounts':
        org_map = {r['registry_id']: r['name'] for r in con.execute('SELECT registry_id, name FROM orgs').fetchall() if r['registry_id']}
    return render_template('list.html', sec=sec, meta=meta, rows=rows, org_map=org_map)


@app.route('/<sec>/template')
def ref_template(sec):
    if sec not in DIR_SECTIONS or not can(sec):
        return redirect(url_for('login'))
    bio = build_template_workbook(sec)
    filename = TEMPLATE_SPECS[sec]['file']
    return send_file(bio, as_attachment=True, download_name=filename)


@app.route('/<sec>/import', methods=['POST'])
def ref_import(sec):
    if sec not in DIR_SECTIONS or not can('import'):
        return redirect(url_for('login'))
    f = request.files.get('file')
    if not f or not f.filename:
        return redirect(url_for('ref_list', sec=sec))
    mode = request.form.get('mode', 'add')
    tmp = os.path.join(DATA_DIR, 'upload_' + secrets.token_hex(4) + '.xlsx')
    f.save(tmp)
    stats = section_import(tmp, DIR_SECTIONS[sec]['table'], DIR_SECTIONS[sec]['rules'], mode)
    os.remove(tmp)
    con = get_db()
    audit(con, (current_user() or {}).get('username', ''), 'Импорт справочника', sec + ': ' + str(stats))
    return render_template('import_result.html', stats=stats, back=url_for('ref_list', sec=sec))


@app.route('/import', methods=['GET', 'POST'])
def import_universal():
    if not can('import'):
        return redirect(url_for('login'))
    if request.method == 'POST':
        f = request.files.get('file')
        if not f or not f.filename:
            return redirect(url_for('import_universal'))
        mode = request.form.get('mode', 'add')
        tmp = os.path.join(DATA_DIR, 'upload_' + secrets.token_hex(4) + '.xlsx')
        f.save(tmp)
        stats = universal_import(tmp, mode)
        os.remove(tmp)
        con = get_db()
        audit(con, (current_user() or {}).get('username', ''), 'Универсальный импорт', str(stats))
        return render_template('import_result.html', stats=stats, back=url_for('index'))
    return render_template('import.html')


@app.route('/import/template')
def import_template():
    if not can('import'):
        return redirect(url_for('login'))
    bio = build_template_workbook('universal')
    return send_file(bio, as_attachment=True, download_name=TEMPLATE_SPECS['universal']['file'])


@app.route('/admin')
def admin():
    if not can('sessions'):
        return redirect(url_for('login'))
    con = get_db()
    users = [dict(r) for r in con.execute('SELECT id, username, fullname, role FROM users ORDER BY username').fetchall()]
    sessions_list = [dict(r) for r in con.execute('SELECT * FROM sessions WHERE logged_out = 0 ORDER BY last_activity DESC').fetchall()]
    audit_rows = [dict(r) for r in con.execute('SELECT * FROM audit_log ORDER BY id DESC LIMIT 50').fetchall()]
    backups = sorted(os.listdir(BACKUP_DIR), reverse=True)[:10]
    return render_template('admin.html', users=users, sessions_list=sessions_list, audit_rows=audit_rows, backups=backups)


@app.route('/admin/action', methods=['POST'])
def admin_action():
    if not can('sessions'):
        return redirect(url_for('login'))
    con = get_db()
    action = request.form.get('action', '')
    user = (current_user() or {}).get('username', '')
    if action == 'add_user':
        username = (request.form.get('username') or '').strip()
        fullname = (request.form.get('fullname') or '').strip()
        role = request.form.get('role', 'viewer')
        password = request.form.get('password') or ''
        if username and password and role in ROLE_TITLES:
            con.execute('INSERT OR REPLACE INTO users (username, fullname, role, password_hash) VALUES (?,?,?,?)',
                        (username, fullname, role, hash_password(password)))
            audit(con, user, 'Добавление пользователя', username + ' (' + role + ')')
    elif action == 'delete_user':
        username = request.form.get('username', '')
        if username != 'admin':
            con.execute('DELETE FROM users WHERE username = ?', (username,))
            audit(con, user, 'Удаление пользователя', username)
    elif action == 'backup':
        path = backup_db()
        audit(con, user, 'Резервное копирование', os.path.basename(path))
    elif action == 'clear_sessions':
        con.execute('UPDATE sessions SET logged_out = 1')
        audit(con, user, 'Очистка активных сессий')
    elif action == 'restore':
        f = request.files.get('restore_file')
        if not f or not f.filename:
            audit(con, user, 'Восстановление из копии', 'Файл не выбран')
        else:
            tmp = os.path.join(DATA_DIR, 'restore_' + secrets.token_hex(4) + '.db')
            f.save(tmp)
            ok, info = restore_db(tmp)
            os.remove(tmp)
            if ok:
                audit(con, user, 'Восстановление из копии', 'Успешно; предохранительная копия: ' + os.path.basename(info))
            else:
                audit(con, user, 'Восстановление из копии', 'Ошибка: ' + info)
    con.commit()
    return redirect(url_for('admin'))



SEARCH_FORM_FILE = os.path.join(APP_DIR, 'forms', 'Форма отчета.xlsx')
SEARCH_TABLE_FILE = os.path.join(APP_DIR, 'forms', 'Отчет таблица.xlsx')


def address_search_groups(query):
    words = [w for w in norm(query).split(' ') if w]
    if not words:
        return []
    pat = '.*'.join(re.escape(w.replace('ё', 'е')) for w in words)
    rx = re.compile(pat)
    con = get_db()
    rows = con.execute(
        "SELECT ad.*, a.consumer_no, a.total_area, a.living_area, a.heated_area, "
        "o.name AS org_name FROM addresses ad "
        "LEFT JOIN accounts a ON a.els = ad.els "
        "LEFT JOIN orgs o ON o.registry_id = a.org_ref "
        "WHERE COALESCE(ad.address_full, '') <> ''").fetchall()
    groups = {}
    for r in rows:
        full = norm(r['address_full']).replace('ё', 'е')
        if not rx.search(full):
            continue
        g = groups.setdefault(r['address_full'], {
            'address_full': r['address_full'], 'city': r['city'] or '',
            'settlement': r['settlement'] or '', 'structure': r['structure'] or '',
            'street': r['street'] or '', 'house': r['house'] or '',
            'corpus': r['corpus'] or '', 'apartment': r['apartment'] or '',
            'accounts': []})
        g['accounts'].append({'els': r['els'], 'consumer_no': r['consumer_no'] or '',
                              'total_area': r['total_area'] or '',
                              'living_area': r['living_area'] or '',
                              'heated_area': r['heated_area'] or '',
                              'org_name': r['org_name'] or ''})
    out = list(groups.values())
    for g in out:
        g['els_list'] = ', '.join(a['els'] for a in g['accounts'])
    return out


def build_report_form(group):
    wb = openpyxl.load_workbook(SEARCH_FORM_FILE)
    ws = wb.active
    ws['C1'] = group['address_full']
    accs = group['accounts']
    if accs:
        first = accs[0]
        ws['C3'] = first['els']
        ws['G3'] = first['total_area']
        ws['I3'] = first['living_area']
        ws['K3'] = first['heated_area']
    start = 6
    data_rows = []
    for rng in ws.merged_cells.ranges:
        if rng.min_col == 2 and rng.max_col == 8 and rng.min_row > 5:
            if rng.min_row < start:
                start = rng.min_row
            data_rows.append(rng.min_row)
    data_rows.sort()
    for i, a in enumerate(accs):
        if i < len(data_rows):
            r = data_rows[i]
        else:
            r = (data_rows[-1] if data_rows else start) + (i - len(data_rows) + 1)
            if r not in data_rows:
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
                ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=11)
        ws.cell(row=r, column=2, value=a.get('org_name') or '')
        ws.cell(row=r, column=9, value=a.get('consumer_no') or '')
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    path = os.path.join(DATA_DIR, 'form_' + stamp + '.xlsx')
    wb.save(path)
    return path


TABLE_COLS = [
    ('city', ['Город']),
    ('settlement', ['Населенный пункт', 'Населённый пункт']),
    ('street', ['Улица']),
    ('house', ['Дом', 'Дом/корпус']),
    ('corpus', ['Корпус']),
    ('apartment', ['Квартира']),
    ('els', ['ЕЛС', 'Единый лицевой счет']),
    ('consumer_no', ['Номер ЛС', 'Номер лицевого счета', 'Номер ЛС организации']),
    ('org_name', ['Организация', 'Наименование организации']),
]


def build_address_table_report(groups):
    wb = openpyxl.load_workbook(SEARCH_TABLE_FILE)
    ws = wb.active
    col_map = {}
    for cell in ws[1]:
        if cell.value is None:
            continue
        h = norm(str(cell.value)).replace('ё', 'е')
        for key, variants in TABLE_COLS:
            if h in [norm(v).replace('ё', 'е') for v in variants] and key not in col_map.values():
                col_map[cell.column] = key
                break
    r = 2
    for g in groups:
        for a in g['accounts']:
            for col, key in col_map.items():
                if key == 'els':
                    val = a['els']
                elif key == 'consumer_no':
                    val = a['consumer_no']
                elif key == 'org_name':
                    val = a['org_name']
                else:
                    val = g.get(key, '')
                ws.cell(row=r, column=col, value=val)
            r += 1
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    path = os.path.join(DATA_DIR, 'table_' + stamp + '.xlsx')
    wb.save(path)
    return path


@app.route('/address_search')
def address_search():
    if not can('search'):
        return redirect(url_for('login'))
    query = (request.args.get('query') or '').strip()
    groups = address_search_groups(query) if query else []
    return render_template('address_search.html', query=query, groups=groups)


@app.route('/address_search/report')
def address_search_report():
    if not can('report'):
        return redirect(url_for('login'))
    query = (request.args.get('query') or '').strip()
    try:
        n = int(request.args.get('n', '0') or 0)
    except ValueError:
        n = 0
    groups = address_search_groups(query) if query else []
    if not groups or n < 0 or n >= len(groups):
        return redirect(url_for('address_search', query=query))
    path = build_report_form(groups[n])
    con = get_db()
    audit(con, (current_user() or {}).get('username', ''), 'Отчет по адресу (Форма отчета)', groups[n]['address_full'])
    return send_file(path, as_attachment=True, download_name='Форма отчета.xlsx')


@app.route('/address_search/report_all')
def address_search_report_all():
    if not can('report'):
        return redirect(url_for('login'))
    query = (request.args.get('query') or '').strip()
    groups = address_search_groups(query) if query else []
    if not groups:
        return redirect(url_for('address_search', query=query))
    path = build_address_table_report(groups)
    con = get_db()
    audit(con, (current_user() or {}).get('username', ''), 'Отчет таблица по найденным адресам', 'Адресов: ' + str(len(groups)))
    return send_file(path, as_attachment=True, download_name='Отчет таблица.xlsx')


if __name__ == '__main__':
    init_db()
    seed_if_empty()
    save_template_files()
    logging.info('Приложение запущено, версия %s', APP_VERSION)
    app.run(host='127.0.0.1', port=5000, debug=False)
