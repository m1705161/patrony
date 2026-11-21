# -*- coding: utf-8 -*-
"""
Презентация «Шардирование и репликация» — черновой макет без фирменного стиля.
Каждый инструмент стека — отдельный слайд. Термины с расшифровкой.
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

INK = RGBColor(0x22, 0x28, 0x33)
MUTE = RGBColor(0x5B, 0x66, 0x76)
ACC = RGBColor(0x1F, 0x6F, 0xEB)
ACC_D = RGBColor(0x0D, 0x4B, 0xA8)
PINK = RGBColor(0xD0, 0x3A, 0x7E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PAPER = RGBColor(0xFF, 0xFF, 0xFF)
CARD = RGBColor(0xF2, 0xF5, 0xF9)
CARD2 = RGBColor(0xE4, 0xEC, 0xF7)
LINE = RGBColor(0xC9, 0xD3, 0xDE)

SW, SH = 13.333, 7.5


def new_deck():
    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    return prs


def slide(prs, bg=PAPER):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    f = s.background.fill
    f.solid()
    f.fore_color.rgb = bg
    return s


def _run(r, text, size, bold=False, color=INK):
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color


def text(s, x, y, w, h, txt, size, bold=False, color=INK, align=PP_ALIGN.LEFT, line=1.0):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    for i, ln in enumerate(txt.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.line_spacing = line
        r = p.add_run(); _run(r, ln, size, bold, color)
    return tf


def shape(s, x, y, w, h, fill=None, line_color=None, rounded=False):
    t = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    sp = s.shapes.add_shape(t, Inches(x), Inches(y), Inches(w), Inches(h))
    if rounded:
        try: sp.adjustments[0] = 0.09
        except Exception: pass
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line_color is None: sp.line.fill.background()
    else: sp.line.color.rgb = line_color; sp.line.width = Pt(1.0)
    return sp


def oval(s, x, y, w, h, fill=None):
    sp = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb = fill
    sp.line.fill.background()
    return sp


def arrow(s, x1, y1, x2, y2, color=ACC, w=2.0):
    conn = s.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    conn.line.color.rgb = color
    conn.line.width = Pt(w)
    ln = conn.line._get_or_add_ln()
    tail = ln.makeelement(qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'})
    ln.append(tail)
    return conn


def label(s, x, y, w, h, txt, size, bold=False, color=WHITE, align=PP_ALIGN.CENTER):
    sp = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = sp.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.03); tf.margin_right = Inches(0.03)
    tf.margin_top = 0; tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = align; p.line_spacing = 1.0
    r = p.add_run(); _run(r, txt, size, bold, color)
    return sp


def node(s, x, y, w, h, txt, fill=ACC, tcolor=WHITE, size=12):
    shape(s, x, y, w, h, fill=fill, rounded=True)
    label(s, x, y, w, h, txt, size, True, tcolor)


def page_num(s, n):
    text(s, 12.55, 7.08, 0.6, 0.3, str(n), 10, False, MUTE, PP_ALIGN.RIGHT)


def takeaway(s, t):
    shape(s, 0.6, 6.42, 12.1, 0.5, fill=CARD2)
    label(s, 0.8, 6.42, 11.7, 0.5, t, 12, True, ACC_D, PP_ALIGN.LEFT)


def header(s, title, sub=None):
    text(s, 0.6, 0.42, 12.1, 0.55, title, 23, True, INK)
    if sub:
        text(s, 0.6, 0.98, 12.1, 0.35, sub, 12.5, False, MUTE)


# ============================== 1-8 ==============================

def s01_title(s):
    shape(s, 0, 0, SW, SH, fill=ACC)
    text(s, 1.0, 1.3, 11.3, 0.4, "ЛЕКЦИЯ 1 · АРХИТЕКТУРНЫЕ ПАТТЕРНЫ", 15, True, CARD2)
    text(s, 1.0, 1.8, 11.3, 2.2, "Шардирование и репликация", 44, True, WHITE)
    text(s, 1.0, 4.0, 11.3, 0.6, "как теория переходит в инженерные расчёты", 22, False, WHITE)
    text(s, 1.0, 5.5, 11.3, 1.0,
         "Разбираем два главных приёма, которые позволяют базе данных\nрасти с нагрузкой: как резать данные и как их копировать.",
         15, False, WHITE)


def s02_scale(s):
    text(s, 0.6, 0.45, 12.1, 0.6, "Одна мощная машина упирается в потолок", 25, True, INK)
    text(s, 0.6, 1.25, 6.0, 0.4, "Вертикально — купить машину мощнее", 14, True, ACC)
    shape(s, 1.5, 1.75, 2.2, 2.6, fill=CARD, line_color=LINE)
    label(s, 1.5, 1.75, 2.2, 2.6, "1 большая\nмашина", 14, True, INK)
    text(s, 0.6, 4.55, 5.6, 1.9,
         "Плюс: просто — один сервер, всё локально.\nМинус: цена растёт быстрее пользы,\nи рано или поздно больше ядер и\nпамяти в одну машину не поставить.", 11.5, False, MUTE)
    text(s, 7.1, 1.25, 6.0, 0.4, "Горизонтально — много обычных машин", 14, True, ACC)
    x = 7.1
    for i in range(4):
        shape(s, x, 1.75, 1.3, 1.6, fill=CARD2, line_color=LINE)
        label(s, x, 1.75, 1.3, 1.6, "узел", 11, True, INK)
        x += 1.43
    text(s, 7.1, 3.55, 5.6, 1.9,
         "Берём много дешёвых серверов и раздаём\nданные между ними. Когда не хватает —\nпросто добавляем ещё. Так и работает\nвесь современный интернет.", 11.5, False, MUTE)
    takeaway(s, "Итог: расти надо не «вверх», а «вширь» — именно этому и посвящена лекция.")
    page_num(s, 2)


def s03_cap(s):
    text(s, 0.6, 0.45, 12.1, 0.6, "Из трёх желаний одновременно сбудутся только два", 25, True, INK)
    cx, cy = 4.3, 4.0
    tri = s.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(cx-3.05), Inches(cy-1.6), Inches(6.1), Inches(3.6))
    tri.fill.solid(); tri.fill.fore_color.rgb = CARD
    tri.line.color.rgb = ACC; tri.line.width = Pt(2)
    label(s, cx-1.6, cy-1.75, 3.2, 0.4, "Consistency", 13, True, ACC)
    label(s, cx-1.6, cy+1.15, 3.2, 0.4, "Availability", 13, True, ACC)
    label(s, cx-1.7, cy-0.25, 3.4, 0.5, "Partition\ntolerance", 11, True, INK)
    shape(s, 8.0, 1.3, 4.7, 4.4, fill=CARD, line_color=LINE, rounded=True)
    text(s, 8.2, 1.45, 4.3, 4.1,
         "Согласованность — все узлы читают\nодни и те же данные, без расхождений.\n\n"
         "Доступность — на каждый запрос\nобязательно приходит ответ.\n\n"
         "Устойчивость к разделению — сеть\nмежду машинами может рваться,\nи это нужно переживать.",
         11, False, INK)
    shape(s, 0.6, 6.45, 12.1, 0.5, fill=CARD2)
    label(s, 0.8, 6.45, 11.7, 0.5,
          "Сеть падает всегда, поэтому «разделение» не выбираем. Остаётся выбор: быстрый ответ или идеально одинаковые данные.",
          11.5, True, ACC_D, PP_ALIGN.LEFT)
    page_num(s, 3)


def s04_evol(s):
    text(s, 0.6, 0.45, 12.1, 0.6, "Три поколения баз: меняется обёртка, идея та же", 24, True, INK)
    boxes = [
        ("MPP", "Много машин, каждая считает\nсвою часть. Для аналитики.\nTeradata, Greenplum"),
        ("NoSQL", "Хранилища «ключ → значение».\nЖертвуют строгой точностью\nради масштаба. Dynamo, Cassandra"),
        ("NewSQL", "Обычный SQL и транзакции,\nно на кластере машин.\nSpanner, CockroachDB"),
    ]
    x = 0.8
    for i, (name, ex) in enumerate(boxes):
        shape(s, x, 1.6, 3.6, 3.2, fill=CARD if i % 2 == 0 else CARD2, line_color=LINE, rounded=True)
        text(s, x+0.22, 1.8, 3.2, 0.5, name, 19, True, ACC)
        text(s, x+0.22, 2.4, 3.2, 2.3, ex, 11.5, False, INK)
        if i < 2:
            arrow(s, x+3.6, 3.2, x+4.2, 3.2, PINK, 2)
        x += 4.2
    text(s, 0.8, 5.1, 11.5, 0.7,
         "Во всех трёх — одна и та же начинка: данные режут на части и копируют. Меняется только, как это названо.", 12.5, True, INK)
    page_num(s, 4)


def s05_shard(s):
    text(s, 0.6, 0.45, 12.1, 0.6, "Шардинг: разрезаем данные и разносим по машинам", 24, True, INK)
    shape(s, 1.0, 1.7, 2.6, 2.0, fill=ACC_D, rounded=True)
    label(s, 1.0, 1.7, 2.6, 2.0, "вся\nтаблица", 14, True, WHITE)
    arrow(s, 3.8, 2.2, 4.7, 1.5, ACC, 2)
    arrow(s, 3.8, 2.7, 4.7, 2.7, ACC, 2)
    arrow(s, 3.8, 3.2, 4.7, 3.9, ACC, 2)
    sy = 1.25
    for i in range(3):
        shape(s, 5.0, sy, 2.4, 0.9, fill=CARD2, line_color=ACC)
        label(s, 5.0, sy, 2.4, 0.9, "часть %d\n(на своей машине)" % (i+1), 10.5, True, INK)
        sy += 1.05
    shape(s, 8.0, 1.3, 4.7, 4.0, fill=CARD, line_color=LINE, rounded=True)
    text(s, 8.2, 1.45, 4.3, 3.7,
         "Партиционирование — разрезать таблицу\nна части. Можно даже в одной базе.\n\n"
         "Шардирование — то же самое, но каждая\nчасть живёт на отдельной машине.\n\n"
         "Зачем: чтобы одна таблица не упёрлась\nв ёмкость и скорость одного сервера.",
         11.5, False, INK)
    takeaway(s, "Одна логическая таблица — много физических кусков на разных машинах.")
    page_num(s, 5)


def s06_strategies(s):
    text(s, 0.6, 0.45, 12.1, 0.6, "Ключ решает, ровно ли лягут данные", 24, True, INK)
    cards = [
        ("Hash — по остатку", "Считаем число от ключа, берём остаток от деления на число машин. Ровно, но при добавлении машины почти всё пересчитывать."),
        ("Range — по диапазону", "Машина хранит свой кусок ключей (А–М, Н–Я). Легко искать по порядку, но свежие записи валят в один хвост."),
        ("Directory — по справочнику", "Отдельная таблица «кто где лежит». Полный контроль, но сам справочник — узкое место."),
        ("Consistent hashing", "При добавлении машины двигается малая часть данных, а не все. Спасение при росте."),
    ]
    for i, (name, d) in enumerate(cards):
        cx = 0.6 + (i % 2) * 6.15
        cy = 1.5 + (i // 2) * 2.4
        shape(s, cx, cy, 5.9, 2.15, fill=CARD if i % 2 == 0 else CARD2, line_color=LINE, rounded=True)
        text(s, cx+0.2, cy+0.12, 5.5, 0.4, name, 15, True, ACC)
        text(s, cx+0.2, cy+0.55, 5.5, 1.5, d, 11, False, INK)
    takeaway(s, "Плохо выбранный ключ = перекос: одна машина задыхается, остальные простаивают.")
    page_num(s, 6)


def s07_repl(s):
    text(s, 0.6, 0.45, 12.1, 0.6, "Репликация: храним несколько копий одних данных", 24, True, INK)
    node(s, 0.9, 1.7, 2.5, 1.7, "Главный\n(принимает запись)", ACC_D, WHITE, 11)
    node(s, 4.9, 1.1, 2.5, 1.2, "Копия 1", CARD2, INK)
    node(s, 4.9, 3.0, 2.5, 1.2, "Копия 2", CARD2, INK)
    arrow(s, 3.4, 2.55, 4.9, 1.7, ACC, 2)
    arrow(s, 3.4, 2.55, 4.9, 3.6, ACC, 2)
    label(s, 3.8, 1.7, 1.3, 0.4, "копирует", 10, False, MUTE)
    shape(s, 7.9, 1.3, 4.8, 4.4, fill=CARD, line_color=LINE, rounded=True)
    text(s, 8.1, 1.45, 4.4, 4.1,
         "Синхронная — ждём, пока копия\nподтвердит запись. Ничего не теряем,\nно запись дольше.\n\n"
         "Асинхронная — не ждём, запись быстрая,\nно при падении главного часть свежих\nданных можно потерять.\n\n"
         "Кворум — правило «сколько копий должно\nподтвердить», чтобы данные не разъехались.",
         11, False, INK)
    takeaway(s, "Репликация — это страховка от падения машины. Цена — лишний трафик и лишнее место на диске.")
    page_num(s, 7)


def s08_consist(s):
    text(s, 0.6, 0.45, 12.1, 0.6, "Кто главный — решает голосование, а не админ", 24, True, INK)
    labels = [("строго одинаково", 0.9), ("по порядку причин", 3.9), ("видно свои записи", 6.9), ("когда-нибудь", 10.2)]
    shape(s, 1.0, 2.3, 11.3, 0.09, fill=LINE)
    for nm, xx in labels:
        oval(s, xx+0.25, 2.2, 0.22, 0.22, fill=ACC)
        label(s, xx-0.5, 2.6, 2.0, 0.4, nm, 10.5, True, INK)
    text(s, 1.0, 3.3, 11.3, 1.0,
         "Когда главная машина падает, копиям нужно самим договориться, кто теперь главный. "
         "Такое голосование между машинами называется консенсусом, а самый понятный его алгоритм — Raft.",
         12.5, False, INK)
    shape(s, 1.0, 4.45, 11.3, 1.0, fill=CARD2, rounded=True)
    label(s, 1.2, 4.45, 10.9, 1.0,
          "Кластер из N = 2f + 1 машин переживает f одновременных падений. etcd на 3 машинах переживает одно.", 12, True, INK)
    page_num(s, 8)


# ============================== 9-13 ядро ==============================

def s09_partition(s):
    header(s, "Ядро: партиционирование в Postgres", "разрезаем одну таблицу на куски внутри одной базы")
    shape(s, 0.9, 1.6, 2.4, 1.9, fill=ACC_D, rounded=True)
    label(s, 0.9, 1.6, 2.4, 1.9, "таблица", 13, True, WHITE)
    arrow(s, 3.6, 2.1, 4.5, 1.5, ACC, 2)
    arrow(s, 3.6, 2.55, 4.5, 2.55, ACC, 2)
    arrow(s, 3.6, 3.0, 4.5, 3.6, ACC, 2)
    py = 1.25
    for i in range(3):
        shape(s, 4.8, py, 2.3, 0.85, fill=CARD2, line_color=ACC)
        label(s, 4.8, py, 2.3, 0.85, "кусок %d" % (i+1), 11, True, INK)
        py += 0.95
    text(s, 7.8, 1.6, 4.9, 2.3,
         "Три способа резать: RANGE (по диапазону), LIST (по списку значений), HASH (по остатку).\n\n"
         "Главная польза — pruning: запрос по ключу читает только нужный кусок, а не всю таблицу.",
         11, False, INK)
    shape(s, 0.6, 4.35, 12.1, 1.5, fill=CARD, rounded=True)
    text(s, 1.0, 4.55, 11.3, 1.0, "число кусков = общий объём ÷ размер одного куска", 20, True, ACC)
    text(s, 1.0, 5.2, 11.3, 0.5, "кусок держим небольшим, чтобы быстро бэкапить и чинить", 12, False, MUTE)
    page_num(s, 9)


def s10_citus(s):
    header(s, "Ядро: шардирование через Citus", "Postgres-кластер: один координатор и несколько хранителей")
    node(s, 0.9, 1.2, 3.0, 1.0, "Coordinator\n(принимает запрос)", ACC_D, WHITE, 11)
    wy = 2.5
    xs = [0.9, 4.1, 7.3]
    for i in range(3):
        node(s, xs[i], wy, 2.6, 1.2, "worker %d\n(хранит часть)" % (i+1), CARD2, INK, 10.5)
        arrow(s, 2.4, 2.2, xs[i]+1.3, wy-0.05, ACC, 1.5)
    text(s, 0.9, 3.9, 8.0, 1.4,
         "distribution column — колонка, по которой режут данные между воркерами. Это и есть hash-шардинг.",
         11.5, False, INK)
    shape(s, 0.9, 5.3, 11.7, 1.0, fill=CARD, rounded=True)
    text(s, 1.1, 5.45, 11.3, 0.75,
         "строк на кусок = всего ÷ shard_count      ·      копии = shard_count × replication_factor",
         16, True, ACC)
    takeaway(s, "shard_count (по умолчанию 32) делают больше числа воркеров — при добавлении машины переезжают целые куски, а не всё заново.")
    page_num(s, 10)


def s11_replpg(s):
    header(s, "Ядро: репликация в Postgres", "главный шлёт журнал изменений, копия его повторяет")
    node(s, 0.9, 1.5, 2.6, 1.5, "Primary\n(принимает запись)", ACC_D, WHITE, 10.5)
    node(s, 5.4, 1.1, 2.6, 1.3, "Replica\n(копия)", CARD2, INK, 10.5)
    arrow(s, 3.5, 2.2, 5.4, 1.75, ACC, 2)
    label(s, 4.0, 1.6, 1.7, 0.4, "WAL — журнал\nизменений", 9, False, MUTE)
    text(s, 0.9, 3.3, 6.8, 2.0,
         "Физическая: главный шлёт копии журнал изменений (WAL), копия повторяет те же операции.\n\n"
         "Логическая: подписываемся не на всю базу, а на отдельные таблицы — удобно для выгрузки в аналитику.",
         11.5, False, INK)
    shape(s, 8.0, 1.3, 4.7, 3.2, fill=CARD, line_color=LINE, rounded=True)
    text(s, 8.2, 1.45, 4.3, 2.9,
         "synchronous_commit — насколько жёстко ждать ответ копии, прежде чем сказать «запись принята».\n\n"
         "off → remote_write → on → remote_apply\n(от «не ждём» до «копия подтвердила»).",
         11, False, INK)
    takeaway(s, "Один параметр synchronous_commit задаёт весь компромисс: писать быстрее или хранить надёжнее.")
    page_num(s, 11)


def s12_patroni(s):
    header(s, "Ядро: автоматический отказоустойчивый Postgres", "Patroni сам назначает нового главного, когда старый упал")
    node(s, 0.7, 1.15, 2.3, 2.2, "etcd — помнит,\nкто сейчас\nглавный", CARD2, INK, 10.5)
    node(s, 4.0, 1.4, 2.3, 1.3, "Primary", ACC_D)
    node(s, 4.0, 3.0, 2.3, 1.0, "Replica", CARD2, INK)
    arrow(s, 3.0, 1.6, 4.0, 2.0, MUTE, 1.5)
    arrow(s, 3.0, 3.4, 4.0, 3.5, MUTE, 1.5)
    arrow(s, 6.3, 2.0, 7.8, 2.0, PINK, 2)
    text(s, 8.0, 1.1, 4.8, 1.1, "failover — смена главного.\nЗасекаем время: ожидаемо 30–60 с.", 11, True, INK)
    text(s, 8.0, 2.5, 4.8, 2.2,
         "RPO — сколько данных можем потерять.\nRTO — сколько времени сервис лежит.\n\n"
         "RTO = обнаружение (loop_wait 10с + ttl 30с) + назначение нового + переподключение клиентов.",
         11, False, MUTE)
    takeaway(s, "Надёжность меряется числами RPO и RTO, а не фразой «у нас есть реплика».")
    page_num(s, 12)


def s13_net(s):
    header(s, "Ядро: во сколько обходится репликация", "каждая копия — это реальные гигабайты на диске и в сети")
    shape(s, 1.3, 1.7, 10.7, 1.5, fill=CARD, line_color=ACC, rounded=True)
    label(s, 1.3, 1.7, 10.7, 1.5, "сетевой трафик = скорость записи × (количество копий − 1)", 21, True, ACC)
    shape(s, 1.3, 3.6, 5.1, 1.9, fill=CARD2, rounded=True)
    text(s, 1.5, 3.75, 4.7, 1.6,
         "Плюс basebackup — полная копия базы,\nкогда реплику поднимают с нуля.\nВ первый раз это может быть много гигабайт.",
         11.5, False, INK)
    shape(s, 6.9, 3.6, 5.1, 1.9, fill=CARD2, rounded=True)
    text(s, 7.1, 3.75, 4.7, 1.6,
         "Размножение хранения: каждая копия — полный дубль данных.\nПолная копия = RF× места,\nerasure coding — экономнее, но сложнее.",
         11.5, False, INK)
    takeaway(s, "Трафик и диск считаем заранее — иначе репликация съест весь бюджет незаметно.")
    page_num(s, 13)


def s14_example(s):
    header(s, "Ядро: считаем от данных до железа", "10 ТБ данных, 20 тысяч чтений в секунду")
    steps = ["считаем куски", "считаем копии", "считаем машины", "считаем сеть"]
    x = 1.2
    for i, st in enumerate(steps):
        last = i == len(steps)-1
        shape(s, x, 1.4, 2.6, 0.95, fill=ACC_D if last else CARD, line_color=LINE, rounded=True)
        label(s, x+0.05, 1.4, 2.5, 0.95, st, 12, True, WHITE if last else INK)
        if i < len(steps)-1:
            arrow(s, x+2.6, 1.85, x+2.95, 1.85, PINK, 2)
        x += 3.0
    rows = [
        ("Шардов", "40", "10 ТБ ÷ 256 ГБ на шард"),
        ("Реплик всего", "120", "40 шардов × 3 копии"),
        ("Узлов", "15–24", "по ёмкости и оперативной памяти"),
        ("Сеть на репликацию", "~96 Мбит/с", "скорость записи × 2 копии"),
        ("RPO / RTO", "≈0 / 30–60 с", "синхронная копия + Patroni"),
    ]
    ry = 2.7
    for name, val, how in rows:
        shape(s, 1.1, ry, 3.1, 0.6, fill=CARD, line_color=LINE, rounded=True)
        label(s, 1.1, ry, 3.1, 0.6, name, 12, True, INK)
        shape(s, 4.35, ry, 2.0, 0.6, fill=CARD2, line_color=LINE, rounded=True)
        label(s, 4.35, ry, 2.0, 0.6, val, 13, True, ACC)
        text(s, 6.55, ry+0.16, 6.0, 0.4, how, 11.5, False, MUTE)
        ry += 0.72
    page_num(s, 14)


# ============================== 15-20 инструменты стека ==============================

def card2(s, title, body, x, y, w, h, tone=CARD):
    shape(s, x, y, w, h, fill=tone, line_color=LINE, rounded=True)
    text(s, x+0.22, y+0.15, w-0.44, 0.4, title, 15, True, ACC)
    text(s, x+0.22, y+0.58, w-0.44, h-0.7, body, 11, False, INK)


def s15_pg(s):
    header(s, "Postgres: всё в одном движке", "то же шардирование и репликация, но без кластера")
    card2(s, "Партиции (шардинг)",
          "RANGE / LIST / HASH — режем таблицу\nна куски внутри одной базы. Запрос по\nключу читает только нужный кусок.",
          0.7, 1.6, 5.8, 2.2)
    card2(s, "Streaming (репликация)",
          "Главный шлёт журнал изменений (WAL)\nкопии, она повторяет те же операции.\nЭто и есть физическая репликация.",
          6.8, 1.6, 5.8, 2.2, CARD2)
    card2(s, "Logical replication",
          "Подписываемся не на всю базу, а на\nотдельные таблицы. Удобно для выгрузки\nв аналитику и обмена между базами.",
          0.7, 4.0, 5.8, 2.2, CARD2)
    card2(s, "Patroni + etcd",
          "Автоматический failover: когда главный\nупал, Patroni через etcd назначает\nнового. Всё без рук админа.",
          6.8, 4.0, 5.8, 2.2)
    takeaway(s, "Начинают всегда с одного Postgres — партиции и реплики работают и без кластера.")
    page_num(s, 15)


def s16_pg_citus(s):
    header(s, "Citus: шардирование поверх Postgres", "расширение, которое превращает несколько Postgres в один кластер")
    node(s, 0.9, 1.2, 3.0, 1.0, "Coordinator\n(принимает запрос)", ACC_D, WHITE, 11)
    xs = [0.9, 4.1, 7.3]
    for i in range(3):
        node(s, xs[i], 2.6, 2.6, 1.2, "worker %d\n(хранит часть)" % (i+1), CARD2, INK, 10.5)
        arrow(s, 2.4, 2.2, xs[i]+1.3, 2.55, ACC, 1.5)
    text(s, 0.9, 4.0, 8.0, 1.6,
         "Одна команда create_distributed_table(t, col) — и таблица сама режется и раскладывается по воркерам.",
         11.5, False, INK)
    card2(s, "Почему это удобно",
          "SQL остаётся обычным, а данные сами\nраспределяются. Распределение видно\nчерез citus_shards.", 9.0, 1.3, 3.6, 2.0)
    takeaway(s, "Citus = Postgres + автоматическое шардирование. В ядре мы считали именно его shard_count — 32.")
    page_num(s, 16)


def s17_greenplum(s):
    header(s, "Greenplum: аналитический double", "раздаёт данные AND дублирует их — оба приёма сразу")
    node(s, 0.9, 1.2, 3.0, 1.1, "Master\n(координатор)", ACC_D, WHITE, 11)
    xs = [0.9, 4.1, 7.3]
    for i in range(3):
        node(s, xs[i], 2.7, 2.8, 1.1, "segment %d\n(своя часть)" % (i+1), CARD2, INK, 10.5)
        node(s, xs[i], 4.1, 2.8, 0.9, "mirror %d\n(копия)" % (i+1), LINE, INK, 9.5)
        arrow(s, 2.4, 2.3, xs[i]+0.4, 2.65, ACC, 1.5)
        arrow(s, xs[i]+1.4, 3.8, xs[i]+1.4, 4.05, MUTE, 1.5)
    text(s, 0.9, 5.2, 11.7, 1.3,
         "DISTRIBUTED BY (колонка) режет данные по сегментам — это hash-шардинг. У каждого сегмента есть mirror-копия — это репликация.",
         11.5, False, INK)
    takeaway(s, "Если ключ распределения не совпал с ключом JOIN — данные гоняются между сегментами, это видно в плане запроса.")
    page_num(s, 17)


def s18_spark(s):
    header(s, "Spark: шардирование в памяти", "партиции — это и есть куски данных, только внутри вычислительной задачи")
    shape(s, 0.7, 1.5, 2.6, 1.4, fill=ACC_D, rounded=True)
    label(s, 0.7, 1.5, 2.6, 1.4, "большой\nдатасет", 13, True, WHITE)
    arrow(s, 3.5, 1.8, 4.3, 1.4, ACC, 2)
    arrow(s, 3.5, 2.2, 4.3, 2.2, ACC, 2)
    arrow(s, 3.5, 2.6, 4.3, 3.0, ACC, 2)
    py = 1.3
    for i in range(3):
        shape(s, 4.5, py, 2.2, 0.75, fill=CARD2, line_color=ACC)
        label(s, 4.5, py, 2.2, 0.75, "партиция %d" % (i+1), 10.5, True, INK)
        py += 0.85
    card2(s, "repartition / coalesce",
          "repartition — заново разложить по\nдругим кускам (дорого, но ровно).\ncoalesce — слить куски без перекладывания.",
          7.4, 1.4, 5.1, 2.2)
    card2(s, "Перекос (skew)",
          "Один «горячий» ключ собирает все\nстроки в одну партицию — одна задача\nтащит всё. Лечим salting и broadcast.",
          7.4, 3.8, 5.1, 2.2, CARD2)
    takeaway(s, "Выбор ключа партиции в Spark — тот же выбор ключа шарда, только на уровне задач, а не машин.")
    page_num(s, 18)


def s19_s3(s):
    header(s, "S3 / MinIO: хранилище объектов", "шардирование спрятано внутри, наружу торчит репликация")
    card2(s, "Репликация объектов",
          "SRR — между бакетами одного региона.\nCRR — между регионами (для надёжности\nи близости к пользователю).",
          0.7, 1.5, 5.8, 2.3)
    card2(s, "Сильная согласованность",
          "С декабря 2020 S3 отдаёт свежие\nданные сразу после записи. Раньше\nприходилось городить костыли.",
          6.8, 1.5, 5.8, 2.3, CARD2)
    card2(s, "Eraser coding (MinIO)",
          "Вместо полных копий данные режут\nна куски и добавляют куски чётности.\nМеньше места, выдерживает отказ дисков.",
          0.7, 4.0, 5.8, 2.3, CARD2)
    card2(s, "Зачем это здесь",
          "Шардирование тут автоматическое — просто кладём объекты. А вот репликацию и цену хранения выбираем сами.",
          6.8, 4.0, 5.8, 2.3)
    takeaway(s, "Объектное хранилище — пример, где шардирование инкапсулировано, а репликация — явный выбор инженера.")
    page_num(s, 19)


def s20_k8s(s):
    header(s, "Kubernetes: те же приёмы для приложения", "копирует не данные, а сами запущенные процессы")
    rows = [
        ("ReplicaSet", "держит заданное число копий приложения. Упала — поднимется новая. Это «репликация процесса»."),
        ("StatefulSet", "стабильные имена pod-0, pod-1 и свои диски. Нужно базе, где каждая копия — конкретный шард."),
        ("PodDisruptionBudget", "не даёт выселить слишком много копий за раз — иначе упадёт и кворум, и сервис."),
        ("anti-affinity", "раскидывает копии по разным машинам и зонам, чтобы одна поломка не убила все копии сразу."),
    ]
    ry = 1.45
    for name, how in rows:
        shape(s, 0.7, ry, 2.7, 1.05, fill=CARD, line_color=LINE, rounded=True)
        label(s, 0.7, ry, 2.7, 1.05, name, 12, True, INK)
        shape(s, 3.5, ry, 9.2, 1.05, fill=CARD2, line_color=LINE, rounded=True)
        text(s, 3.7, ry+0.16, 8.8, 0.8, how, 11.5, False, INK)
        ry += 1.17
    takeaway(s, "Kubernetes не знает про «шарды данных», но даёт ровно те примитивы, из которых шардированная база собирается.")
    page_num(s, 20)


# ============================== 21-23 практика и финал ==============================

def s21_citus_prac(s):
    header(s, "Практика: шардирование видно глазами", "поднимаем Citus в docker-compose")
    node(s, 0.9, 1.2, 3.0, 0.95, "coordinator\n(docker-compose)", ACC_D, WHITE, 10.5)
    xs = [0.9, 4.1, 7.3]
    for i in range(3):
        node(s, xs[i], 2.4, 2.6, 1.1, "worker %d" % (i+1), CARD2, INK, 11)
        arrow(s, 2.4, 2.15, xs[i]+1.3, 2.35, ACC, 1.5)
    text(s, 8.2, 1.15, 4.4, 0.9, "Что делаем по шагам:", 14, True, INK)
    steps = [
        "1. Включаем coordinator + workers",
        "2. create_distributed_table",
        "3. Смотрим, сколько строк на каждом воркере",
        "4. Меняем ключ на плохой — видим перекос",
    ]
    for i, st in enumerate(steps):
        text(s, 8.2, 1.75 + i*0.5, 4.5, 0.45, st, 11, False, INK)
    text(s, 0.9, 4.0, 7.8, 1.8,
         "Что это даёт: план запроса показывает, когда данные с одного ключа лежат рядом — база не гоняет их между машинами. "
         "Перекос видно цифрами, а не интуицией.",
         11.5, False, INK)
    takeaway(s, "Сколько строк уехало на каждый worker — это и есть шардирование, измеренное в числах.")
    page_num(s, 21)


def s22_patroni_prac(s):
    header(s, "Практика: RTO превращаем в секундомер", "засекаем, за сколько Patroni поднимет нового главного")
    steps = ["1. отключаем\nprimary", "2. истекает\nаренда (ttl)", "3. выборы\nнового", "4. новый\nprimary готов"]
    x = 0.7
    for i, st in enumerate(steps):
        last = i == len(steps)-1
        shape(s, x, 1.8, 2.6, 1.6, fill=ACC_D if last else CARD, line_color=LINE, rounded=True)
        label(s, x+0.08, 1.8, 2.44, 1.6, st, 11, True, WHITE if last else INK)
        if i < len(steps)-1:
            arrow(s, x+2.6, 2.6, x+3.0, 2.6, PINK, 2)
        x += 3.0
    text(s, 0.7, 3.8, 11.5, 0.8,
         "Засекаем: от «docker stop primary» до момента, когда новый главный готов. Ожидаемо 30–60 секунд.", 13.5, True, INK)
    text(s, 0.7, 4.6, 11.5, 1.2,
         "Дальше pg_rewind возвращает упавший узел в строй как копию. При синхронной копии потерянных данных ноль — RPO = 0.",
         11.5, False, MUTE)
    page_num(s, 22)


def s23_close(s):
    shape(s, 0, 0, SW, SH, fill=ACC)
    text(s, 1.0, 0.9, 11.3, 1.0, "Что выносим с лекции", 32, True, WHITE)
    blocks = [
        ("Теория", ["шардирование режет данные\nрепликация их копирует", "это два независимых приёма", "CAP: из трёх гарантий — две"]),
        ("Расчёты", ["шардов — по объёму данных\nмашин — по ёмкости и памяти", "сеть — скорость × копии", "RPO/RTO — сколько теряем и лежим"]),
        ("Практика", ["Citus: шардирование вживую", "Patroni: failover по секундомеру", "оба на готовом docker-compose"]),
    ]
    x = 0.9
    for h, balls in blocks:
        shape(s, x, 2.1, 3.75, 3.8, fill=ACC_D, line_color=RGBColor(0x5A, 0x98, 0xF0), rounded=True)
        text(s, x+0.25, 2.3, 3.25, 0.5, h, 18, True, WHITE)
        text(s, x+0.25, 2.95, 3.25, 2.7, "\n".join("• " + b for b in balls), 12, False, WHITE)
        x += 3.95
    text(s, 1.0, 6.3, 11.3, 0.5, "Спасибо за внимание", 17, True, WHITE)


def main():
    prs = new_deck()
    s01_title(slide(prs, ACC))
    s02_scale(slide(prs))
    s03_cap(slide(prs))
    s04_evol(slide(prs))
    s05_shard(slide(prs))
    s06_strategies(slide(prs))
    s07_repl(slide(prs))
    s08_consist(slide(prs))
    s09_partition(slide(prs))
    s10_citus(slide(prs))
    s11_replpg(slide(prs))
    s12_patroni(slide(prs))
    s13_net(slide(prs))
    s14_example(slide(prs))
    s15_pg(slide(prs))
    s16_pg_citus(slide(prs))
    s17_greenplum(slide(prs))
    s18_spark(slide(prs))
    s19_s3(slide(prs))
    s20_k8s(slide(prs))
    s21_citus_prac(slide(prs))
    s22_patroni_prac(slide(prs))
    s23_close(slide(prs, ACC))

    out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "презентация-шардирование-репликация.pptx")
    prs.save(out)
    print("Saved:", out, "| slides:", len(prs.slides))


if __name__ == "__main__":
    main()
