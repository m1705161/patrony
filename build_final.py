# -*- coding: utf-8 -*-
"""
Итоговая сборка: каркас из Занятие_21 (титул, связь, правила, рекап, рефлексия)
+ мой контент по шардированию/репликации в том же стиле (Play, #0077FF).
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.parts.presentation import PresentationPart
from pptx.opc.packuri import PackURI

TEMPLATE = "/Users/a.krivyakov/Downloads/Занятие_21_Шардирование_и_репликация.pptx"

# палитра VK Education
BLUE  = RGBColor(0x00, 0x77, 0xFF)
BLUE_D = RGBColor(0x00, 0x52, 0xCC)
PINK  = RGBColor(0xFF, 0x38, 0x85)
CYAN  = RGBColor(0x7C, 0xED, 0xF8)
LIGHT = RGBColor(0xEB, 0xF3, 0xF9)
LIGHT2 = RGBColor(0xD8, 0xEB, 0xFC)
INK   = RGBColor(0x1F, 0x29, 0x37)
MUTE  = RGBColor(0x5B, 0x66, 0x76)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE  = RGBColor(0xC9, 0xD3, 0xDE)

FONT = "Play"
SW, SH = 13.333, 7.5


# ---------- основы ----------

def _patch_next_slide_partname(prs):
    def next_slide_partname(self):
        used = set()
        for sldId in self._element.sldIdLst:
            used.add(self.related_part(sldId.rId).partname)
        n = 1
        while True:
            pn = PackURI("/ppt/slides/slide%d.xml" % n)
            if pn not in used:
                return pn
            n += 1
    PresentationPart._next_slide_partname = property(next_slide_partname)


def delete_slides(prs, indices):
    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    for i in sorted(indices, reverse=True):
        el = ids[i]
        rId = el.get(qn('r:id'))
        prs.part.drop_rel(rId)
        sldIdLst.remove(el)


def reorder_slides(prs, order):
    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    for el in ids:
        sldIdLst.remove(el)
    for i in order:
        sldIdLst.append(ids[i])


def get_layout(prs, name):
    for lay in prs.slide_layouts:
        if lay.name == name:
            return lay
    raise KeyError(name)


def _run(r, text, size, bold=False, color=INK):
    r.text = text
    r.font.name = FONT
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


def arrow(s, x1, y1, x2, y2, color=BLUE, w=2.0):
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


def node(s, x, y, w, h, txt, fill=BLUE, tcolor=WHITE, size=12):
    shape(s, x, y, w, h, fill=fill, rounded=True)
    label(s, x, y, w, h, txt, size, True, tcolor)


def page_num(s, n):
    text(s, 12.55, 7.08, 0.6, 0.3, str(n), 10, False, MUTE, PP_ALIGN.RIGHT)


def takeaway(s, t):
    shape(s, 0.6, 6.42, 12.1, 0.5, fill=LIGHT, rounded=True)
    label(s, 0.8, 6.42, 11.7, 0.5, t, 12, True, PINK, PP_ALIGN.LEFT)


def add_header(s, title, sub=None):
    """Пишет заголовок в title-плейсхолдер лейаута (стиль шаблона)."""
    t = s.shapes.title
    if t is not None:
        t.left = Inches(0.74); t.top = Inches(0.42); t.width = Inches(11.9); t.height = Inches(0.95)
        tf = t.text_frame
        tf.word_wrap = True
        tf.clear()
        p = tf.paragraphs[0]
        r = p.add_run(); _run(r, title, 24, True, BLUE)
        if sub:
            sp = text(s, 0.74, 1.22, 11.9, 0.4, sub, 12.5, False, MUTE)
        return 1.72
    else:
        text(s, 0.74, 0.42, 11.9, 0.95, title, 24, True, BLUE)
        if sub:
            text(s, 0.74, 1.22, 11.9, 0.4, sub, 12.5, False, MUTE)
        return 1.72


def card2(s, title, body, x, y, w, h, tone=LIGHT):
    shape(s, x, y, w, h, fill=tone, line_color=LINE, rounded=True)
    text(s, x+0.22, y+0.15, w-0.44, 0.4, title, 15, True, BLUE)
    text(s, x+0.22, y+0.58, w-0.44, h-0.7, body, 11, False, INK)


# ---------- контент ----------

def c_scale(s):
    add_header(s, "Одна мощная машина упирается в потолок")
    text(s, 0.74, 1.85, 6.0, 0.4, "Вертикально — купить машину мощнее", 14, True, BLUE)
    shape(s, 1.6, 2.35, 2.2, 2.4, fill=LIGHT, line_color=LINE)
    label(s, 1.6, 2.35, 2.2, 2.4, "1 большая\nмашина", 14, True, INK)
    text(s, 0.74, 4.95, 5.6, 1.5,
         "Плюс: просто — один сервер, всё локально.\nМинус: цена растёт быстрее пользы,\nи больше ядер/памяти некуда вставить.",
         11.5, False, MUTE)
    text(s, 7.2, 1.85, 5.9, 0.4, "Горизонтально — много обычных машин", 14, True, BLUE)
    x = 7.2
    for i in range(4):
        shape(s, x, 2.35, 1.3, 1.55, fill=LIGHT2, line_color=LINE)
        label(s, x, 2.35, 1.3, 1.55, "узел", 11, True, INK)
        x += 1.45
    text(s, 7.2, 4.1, 5.6, 1.5,
         "Берём много дешёвых серверов и раздаём\nданные между ними. Не хватает —\nдобавляем ещё. Так работает весь интернет.",
         11.5, False, MUTE)
    takeaway(s, "Итог: расти надо не «вверх», а «вширь» — именно этому и посвящена лекция.")
    page_num(s, 1)


def c_cap(s):
    add_header(s, "Из трёх желаний одновременно сбудутся только два")
    cx, cy = 4.3, 4.0
    tri = s.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(cx-3.05), Inches(cy-1.6), Inches(6.1), Inches(3.6))
    tri.fill.solid(); tri.fill.fore_color.rgb = LIGHT
    tri.line.color.rgb = BLUE; tri.line.width = Pt(2)
    label(s, cx-1.6, cy-1.75, 3.2, 0.4, "Consistency", 13, True, BLUE)
    label(s, cx-1.6, cy+1.15, 3.2, 0.4, "Availability", 13, True, BLUE)
    label(s, cx-1.7, cy-0.25, 3.4, 0.5, "Partition\ntolerance", 11, True, INK)
    shape(s, 8.0, 1.3, 4.7, 4.4, fill=LIGHT, line_color=LINE, rounded=True)
    text(s, 8.2, 1.45, 4.3, 4.1,
         "Согласованность — все узлы читают\nодни и те же данные, без расхождений.\n\n"
         "Доступность — на каждый запрос\nобязательно приходит ответ.\n\n"
         "Устойчивость к разделению — сеть\nмежду машинами может рваться,\nи это нужно переживать.",
         11, False, INK)
    shape(s, 0.6, 6.45, 12.1, 0.5, fill=LIGHT, rounded=True)
    label(s, 0.8, 6.45, 11.7, 0.5,
          "Сеть падает всегда, поэтому «разделение» не выбираем. Остаётся выбор: быстрый ответ или одинаковые данные.",
          11.5, True, PINK, PP_ALIGN.LEFT)
    page_num(s, 2)


def c_evol(s):
    add_header(s, "Три поколения баз: меняется обёртка, идея та же")
    boxes = [
        ("MPP", "Много машин, каждая считает\nсвою часть. Для аналитики.\nTeradata, Greenplum"),
        ("NoSQL", "Хранилища «ключ → значение».\nЖертвуют точностью ради\nмасштаба. Dynamo, Cassandra"),
        ("NewSQL", "Обычный SQL и транзакции,\nно на кластере машин.\nSpanner, CockroachDB"),
    ]
    x = 0.8
    for i, (name, ex) in enumerate(boxes):
        shape(s, x, 1.9, 3.6, 3.0, fill=LIGHT if i % 2 == 0 else LIGHT2, line_color=LINE, rounded=True)
        text(s, x+0.22, 2.1, 3.2, 0.5, name, 19, True, BLUE)
        text(s, x+0.22, 2.7, 3.2, 2.0, ex, 11.5, False, INK)
        if i < 2:
            arrow(s, x+3.6, 3.4, x+4.2, 3.4, PINK, 2)
        x += 4.2
    text(s, 0.8, 5.2, 11.5, 0.7,
         "Во всех трёх одна начинка: данные режут на части и копируют. Меняется только, как это названо.", 12.5, True, INK)
    page_num(s, 3)


def c_shard(s):
    add_header(s, "Шардинг: разрезаем данные и разносим по машинам")
    shape(s, 1.0, 1.9, 2.6, 2.0, fill=BLUE_D, rounded=True)
    label(s, 1.0, 1.9, 2.6, 2.0, "вся\nтаблица", 14, True, WHITE)
    arrow(s, 3.8, 2.4, 4.7, 1.7, BLUE, 2)
    arrow(s, 3.8, 2.9, 4.7, 2.9, BLUE, 2)
    arrow(s, 3.8, 3.4, 4.7, 4.1, BLUE, 2)
    sy = 1.45
    for i in range(3):
        shape(s, 5.0, sy, 2.4, 0.9, fill=LIGHT2, line_color=BLUE)
        label(s, 5.0, sy, 2.4, 0.9, "часть %d\n(на своей машине)" % (i+1), 10.5, True, INK)
        sy += 1.05
    shape(s, 8.0, 1.5, 4.7, 3.8, fill=LIGHT, line_color=LINE, rounded=True)
    text(s, 8.2, 1.65, 4.3, 3.5,
         "Партиционирование — разрезать таблицу\nна части. Можно даже в одной базе.\n\n"
         "Шардирование — то же самое, но каждая\nчасть живёт на отдельной машине.\n\n"
         "Зачем: чтобы одна таблица не упёрлась\nв ёмкость и скорость одного сервера.",
         11.5, False, INK)
    takeaway(s, "Одна логическая таблица — много физических кусков на разных машинах.")
    page_num(s, 4)


def c_strategies(s):
    add_header(s, "Ключ решает, ровно ли лягут данные")
    cards = [
        ("Hash — по остатку", "Считаем число от ключа, берём остаток от деления на число машин. Ровно, но при добавлении машины почти всё пересчитывать."),
        ("Range — по диапазону", "Машина хранит свой кусок ключей (А–М, Н–Я). Легко искать по порядку, но свежие записи валят в один хвост."),
        ("Directory — по справочнику", "Отдельная таблица «кто где лежит». Полный контроль, но сам справочник — узкое место."),
        ("Consistent hashing", "При добавлении машины двигается малая часть данных, а не все. Спасение при росте."),
    ]
    for i, (name, d) in enumerate(cards):
        cx = 0.6 + (i % 2) * 6.15
        cy = 1.85 + (i // 2) * 2.2
        shape(s, cx, cy, 5.9, 2.0, fill=LIGHT if i % 2 == 0 else LIGHT2, line_color=LINE, rounded=True)
        text(s, cx+0.2, cy+0.12, 5.5, 0.4, name, 15, True, BLUE)
        text(s, cx+0.2, cy+0.55, 5.5, 1.4, d, 11, False, INK)
    takeaway(s, "Плохо выбранный ключ = перекос: одна машина задыхается, остальные простаивают.")
    page_num(s, 5)


def c_repl(s):
    add_header(s, "Репликация: храним несколько копий одних данных")
    node(s, 0.9, 1.9, 2.5, 1.7, "Главный\n(принимает запись)", BLUE_D, WHITE, 11)
    node(s, 4.9, 1.3, 2.5, 1.2, "Копия 1", CYAN, INK)
    node(s, 4.9, 3.2, 2.5, 1.2, "Копия 2", CYAN, INK)
    arrow(s, 3.4, 2.75, 4.9, 1.9, BLUE, 2)
    arrow(s, 3.4, 2.75, 4.9, 3.8, BLUE, 2)
    label(s, 3.9, 1.9, 1.3, 0.4, "копирует", 10, False, MUTE)
    shape(s, 7.9, 1.5, 4.8, 4.2, fill=LIGHT, line_color=LINE, rounded=True)
    text(s, 8.1, 1.65, 4.4, 3.9,
         "Синхронная — ждём, пока копия\nподтвердит запись. Ничего не теряем,\nно запись дольше.\n\n"
         "Асинхронная — не ждём, запись быстрая,\nно при падении главного часть свежих\nданных можно потерять.\n\n"
         "Кворум — «сколько копий должно\nподтвердить», чтобы данные не разъехались.",
         11, False, INK)
    takeaway(s, "Репликация — страховка от падения машины. Цена — лишний трафик и лишнее место на диске.")
    page_num(s, 6)


def c_consist(s):
    add_header(s, "Кто главный — решает голосование, а не админ")
    labels = [("строго одинаково", 0.9), ("по порядку причин", 3.9), ("видно свои записи", 6.9), ("когда-нибудь", 10.2)]
    shape(s, 1.0, 2.5, 11.3, 0.09, fill=LINE)
    for nm, xx in labels:
        oval(s, xx+0.25, 2.4, 0.22, 0.22, fill=BLUE)
        label(s, xx-0.5, 2.8, 2.0, 0.4, nm, 10.5, True, INK)
    text(s, 1.0, 3.5, 11.3, 1.0,
         "Когда главная машина падает, копиям нужно самим договориться, кто теперь главный. "
         "Такое голосование называется консенсусом, а самый понятный его алгоритм — Raft.",
         12.5, False, INK)
    shape(s, 1.0, 4.6, 11.3, 1.0, fill=LIGHT, rounded=True)
    label(s, 1.2, 4.6, 10.9, 1.0,
          "Кластер из N = 2f + 1 машин переживает f одновременных падений. etcd на 3 машинах переживает одно.", 12, True, INK)
    page_num(s, 7)


def c_partition(s):
    add_header(s, "Ядро: партиционирование в Postgres", "разрезаем одну таблицу на куски внутри одной базы")
    shape(s, 0.9, 2.0, 2.4, 1.9, fill=BLUE_D, rounded=True)
    label(s, 0.9, 2.0, 2.4, 1.9, "таблица", 13, True, WHITE)
    arrow(s, 3.6, 2.5, 4.5, 1.9, BLUE, 2)
    arrow(s, 3.6, 2.95, 4.5, 2.95, BLUE, 2)
    arrow(s, 3.6, 3.4, 4.5, 4.0, BLUE, 2)
    py = 1.65
    for i in range(3):
        shape(s, 4.8, py, 2.3, 0.85, fill=LIGHT2, line_color=BLUE)
        label(s, 4.8, py, 2.3, 0.85, "кусок %d" % (i+1), 11, True, INK)
        py += 0.95
    text(s, 7.8, 2.0, 4.9, 2.2,
         "RANGE (по диапазону), LIST (по списку значений), HASH (по остатку).\n\n"
         "Главная польза — pruning: запрос по ключу читает только нужный кусок.",
         11, False, INK)
    shape(s, 0.6, 4.55, 12.1, 1.4, fill=LIGHT, rounded=True)
    text(s, 1.0, 4.75, 11.3, 0.9, "число кусков = общий объём ÷ размер одного куска", 20, True, BLUE)
    text(s, 1.0, 5.4, 11.3, 0.5, "кусок держим небольшим, чтобы быстро бэкапить и чинить", 12, False, MUTE)
    page_num(s, 8)


def c_citus(s):
    add_header(s, "Ядро: шардирование через Citus", "Postgres-кластер: один координатор и несколько хранителей")
    node(s, 0.9, 1.6, 3.0, 1.0, "Coordinator\n(принимает запрос)", BLUE_D, WHITE, 11)
    wy = 2.9
    xs = [0.9, 4.1, 7.3]
    for i in range(3):
        node(s, xs[i], wy, 2.6, 1.2, "worker %d\n(хранит часть)" % (i+1), CYAN, INK, 10.5)
        arrow(s, 2.4, 2.6, xs[i]+1.3, wy, BLUE, 1.5)
    text(s, 0.9, 4.3, 11.0, 0.9,
         "distribution column — колонка, по которой режут данные между воркерами. Это и есть hash-шардинг.",
         11.5, False, INK)
    shape(s, 0.9, 5.25, 11.7, 1.0, fill=LIGHT, rounded=True)
    text(s, 1.1, 5.42, 11.3, 0.7,
         "строк на кусок = всего ÷ shard_count      ·      копии = shard_count × replication_factor",
         16, True, BLUE)
    takeaway(s, "shard_count (по умолчанию 32) делают больше числа воркеров — при добавлении машины переезжают целые куски, а не всё заново.")
    page_num(s, 9)


def c_replpg(s):
    add_header(s, "Ядро: репликация в Postgres", "главный шлёт журнал изменений, копия его повторяет")
    node(s, 0.9, 1.9, 2.6, 1.5, "Primary\n(принимает запись)", BLUE_D, WHITE, 10.5)
    node(s, 5.4, 1.5, 2.6, 1.3, "Replica\n(копия)", CYAN, INK, 10.5)
    arrow(s, 3.5, 2.6, 5.4, 2.15, BLUE, 2)
    label(s, 4.0, 2.0, 1.7, 0.4, "WAL — журнал\nизменений", 9, False, MUTE)
    text(s, 0.9, 3.7, 6.8, 2.0,
         "Физическая: главный шлёт копии журнал изменений (WAL), копия повторяет те же операции.\n\n"
         "Логическая: подписываемся не на всю базу, а на отдельные таблицы — удобно для выгрузки в аналитику.",
         11.5, False, INK)
    shape(s, 8.0, 1.7, 4.7, 3.0, fill=LIGHT, line_color=LINE, rounded=True)
    text(s, 8.2, 1.85, 4.3, 2.7,
         "synchronous_commit — насколько жёстко ждать ответ копии, прежде чем сказать «запись принята».\n\n"
         "off → remote_write → on → remote_apply",
         11, False, INK)
    takeaway(s, "Один параметр synchronous_commit задаёт весь компромисс: писать быстрее или хранить надёжнее.")
    page_num(s, 10)


def c_patroni(s):
    add_header(s, "Ядро: автоматический отказоустойчивый Postgres", "Patroni сам назначает нового главного, когда старый упал")
    node(s, 0.7, 1.5, 2.3, 2.2, "etcd — помнит,\nкто сейчас\nглавный", CYAN, INK, 10.5)
    node(s, 4.0, 1.7, 2.3, 1.3, "Primary", BLUE_D)
    node(s, 4.0, 3.3, 2.3, 1.0, "Replica", CYAN, INK)
    arrow(s, 3.0, 1.9, 4.0, 2.3, MUTE, 1.5)
    arrow(s, 3.0, 3.8, 4.0, 3.8, MUTE, 1.5)
    arrow(s, 6.3, 2.3, 7.8, 2.3, PINK, 2)
    text(s, 8.0, 1.4, 4.8, 1.1, "failover — смена главного.\nЗасекаем время: ожидаемо 30–60 с.", 11, True, INK)
    text(s, 8.0, 2.8, 4.8, 2.2,
         "RPO — сколько данных можем потерять.\nRTO — сколько времени сервис лежит.\n\n"
         "RTO = обнаружение (loop_wait 10с + ttl 30с) + назначение нового + переподключение клиентов.",
         11, False, MUTE)
    takeaway(s, "Надёжность меряется числами RPO и RTO, а не фразой «у нас есть реплика».")
    page_num(s, 11)


def c_net(s):
    add_header(s, "Ядро: во сколько обходится репликация", "каждая копия — это реальные гигабайты на диске и в сети")
    shape(s, 1.3, 2.1, 10.7, 1.5, fill=LIGHT, line_color=BLUE, rounded=True)
    label(s, 1.3, 2.1, 10.7, 1.5, "сетевой трафик = скорость записи × (количество копий − 1)", 21, True, BLUE)
    shape(s, 1.3, 4.0, 5.1, 1.9, fill=LIGHT2, rounded=True)
    text(s, 1.5, 4.15, 4.7, 1.6,
         "Плюс basebackup — полная копия базы,\nкогда реплику поднимают с нуля.\nВ первый раз это много гигабайт.",
         11.5, False, INK)
    shape(s, 6.9, 4.0, 5.1, 1.9, fill=LIGHT2, rounded=True)
    text(s, 7.1, 4.15, 4.7, 1.6,
         "Размножение хранения: каждая копия — полный дубль данных.\nПолная копия = RF× места,\nerasure coding — экономнее.",
         11.5, False, INK)
    takeaway(s, "Трафик и диск считаем заранее — иначе репликация съест весь бюджет незаметно.")
    page_num(s, 12)


def c_example(s):
    add_header(s, "Ядро: считаем от данных до железа", "10 ТБ данных, 20 тысяч чтений в секунду")
    steps = ["считаем куски", "считаем копии", "считаем машины", "считаем сеть"]
    x = 1.2
    for i, st in enumerate(steps):
        last = i == len(steps)-1
        shape(s, x, 1.8, 2.6, 0.95, fill=BLUE if last else LIGHT2, line_color=LINE, rounded=True)
        label(s, x+0.05, 1.8, 2.5, 0.95, st, 12, True, WHITE if last else INK)
        if i < len(steps)-1:
            arrow(s, x+2.6, 2.3, x+2.95, 2.3, PINK, 2)
        x += 3.0
    rows = [
        ("Шардов", "40", "10 ТБ ÷ 256 ГБ на шард"),
        ("Реплик всего", "120", "40 шардов × 3 копии"),
        ("Узлов", "15–24", "по ёмкости и оперативной памяти"),
        ("Сеть на репликацию", "~96 Мбит/с", "скорость записи × 2 копии"),
        ("RPO / RTO", "≈0 / 30–60 с", "синхронная копия + Patroni"),
    ]
    ry = 3.15
    for name, val, how in rows:
        shape(s, 1.1, ry, 3.1, 0.6, fill=LIGHT, line_color=LINE, rounded=True)
        label(s, 1.1, ry, 3.1, 0.6, name, 12, True, INK)
        shape(s, 4.35, ry, 2.0, 0.6, fill=LIGHT2, line_color=LINE, rounded=True)
        label(s, 4.35, ry, 2.0, 0.6, val, 13, True, BLUE)
        text(s, 6.55, ry+0.16, 6.0, 0.4, how, 11.5, False, MUTE)
        ry += 0.72
    page_num(s, 13)


def c_pg(s):
    add_header(s, "Postgres: всё в одном движке", "то же шардирование и репликация, но без кластера")
    card2(s, "Партиции (шардинг)", "RANGE / LIST / HASH — режем таблицу\nна куски внутри одной базы. Запрос по\nключу читает только нужный кусок.", 0.7, 1.9, 5.8, 2.1)
    card2(s, "Streaming (репликация)", "Главный шлёт журнал изменений (WAL)\nкопии, она повторяет операции.\nЭто и есть физическая репликация.", 6.8, 1.9, 5.8, 2.1, LIGHT2)
    card2(s, "Logical replication", "Подписываемся не на всю базу, а на\nотдельные таблицы. Удобно для\nвыгрузки в аналитику.", 0.7, 4.2, 5.8, 2.1, LIGHT2)
    card2(s, "Patroni + etcd", "Автоматический failover: когда главный\nупал, Patroni через etcd назначает\nнового. Всё без рук админа.", 6.8, 4.2, 5.8, 2.1)
    takeaway(s, "Начинают всегда с одного Postgres — партиции и реплики работают и без кластера.")
    page_num(s, 14)


def c_pg_citus(s):
    add_header(s, "Citus: шардирование поверх Postgres", "расширение, которое превращает несколько Postgres в один кластер")
    node(s, 0.9, 1.6, 3.0, 1.0, "Coordinator\n(принимает запрос)", BLUE_D, WHITE, 11)
    xs = [0.9, 4.1, 7.3]
    for i in range(3):
        node(s, xs[i], 2.9, 2.6, 1.2, "worker %d\n(хранит часть)" % (i+1), CYAN, INK, 10.5)
        arrow(s, 2.4, 2.6, xs[i]+1.3, 2.85, BLUE, 1.5)
    text(s, 0.9, 4.3, 8.0, 1.5,
         "Одна команда create_distributed_table(t, col) — и таблица сама режется и раскладывается по воркерам.",
         11.5, False, INK)
    card2(s, "Почему удобно", "SQL остаётся обычным, данные сами\nраспределяются. Распределение видно\nчерез citus_shards.", 9.0, 1.7, 3.6, 2.0, LIGHT2)
    takeaway(s, "Citus = Postgres + автоматическое шардирование. В ядре мы считали именно его shard_count — 32.")
    page_num(s, 15)


def c_greenplum(s):
    add_header(s, "Greenplum: аналитический double", "раздаёт данные AND дублирует их — оба приёма сразу")
    node(s, 0.9, 1.6, 3.0, 1.1, "Master\n(координатор)", BLUE_D, WHITE, 11)
    xs = [0.9, 4.1, 7.3]
    for i in range(3):
        node(s, xs[i], 3.0, 2.8, 1.1, "segment %d\n(своя часть)" % (i+1), CYAN, INK, 10.5)
        node(s, xs[i], 4.35, 2.8, 0.9, "mirror %d\n(копия)" % (i+1), LIGHT2, INK, 9.5)
        arrow(s, 2.4, 2.7, xs[i]+0.4, 2.95, BLUE, 1.5)
        arrow(s, xs[i]+1.4, 4.1, xs[i]+1.4, 4.35, MUTE, 1.5)
    text(s, 0.9, 5.4, 11.7, 1.2,
         "DISTRIBUTED BY (колонка) режет данные по сегментам — это hash-шардинг. У каждого сегмента есть mirror-копия — это репликация.",
         11.5, False, INK)
    takeaway(s, "Если ключ распределения не совпал с ключом JOIN — данные гоняются между сегментами, это видно в плане.")
    page_num(s, 16)


def c_spark(s):
    add_header(s, "Spark: шардирование в памяти", "партиции — это и есть куски данных, только внутри задачи")
    shape(s, 0.7, 1.9, 2.6, 1.4, fill=BLUE_D, rounded=True)
    label(s, 0.7, 1.9, 2.6, 1.4, "большой\nдатасет", 13, True, WHITE)
    arrow(s, 3.5, 2.2, 4.3, 1.8, BLUE, 2)
    arrow(s, 3.5, 2.6, 4.3, 2.6, BLUE, 2)
    arrow(s, 3.5, 3.0, 4.3, 3.4, BLUE, 2)
    py = 1.7
    for i in range(3):
        shape(s, 4.5, py, 2.2, 0.75, fill=LIGHT2, line_color=BLUE)
        label(s, 4.5, py, 2.2, 0.75, "партиция %d" % (i+1), 10.5, True, INK)
        py += 0.85
    card2(s, "repartition / coalesce", "repartition — заново разложить по\nдругим кускам (дорого, но ровно).\ncoalesce — слить без перекладывания.", 7.4, 1.8, 5.1, 2.1, LIGHT2)
    card2(s, "Перекос (skew)", "Один «горячий» ключ собирает все\nстроки в одну партицию — одна задача\nтащит всё. Лечим salting и broadcast.", 7.4, 4.1, 5.1, 2.1)
    takeaway(s, "Выбор ключа партиции в Spark — тот же выбор ключа шарда, только на уровне задач, а не машин.")
    page_num(s, 17)


def c_s3(s):
    add_header(s, "S3 / MinIO: хранилище объектов", "шардирование спрятано внутри, наружу торчит репликация")
    card2(s, "Репликация объектов", "SRR — между бакетами одного региона.\nCRR — между регионами (для надёжности\nи близости к пользователю).", 0.7, 1.9, 5.8, 2.2)
    card2(s, "Сильная согласованность", "С декабря 2020 S3 отдаёт свежие\nданные сразу после записи. Раньше\nприходилось городить костыли.", 6.8, 1.9, 5.8, 2.2, LIGHT2)
    card2(s, "Eraser coding (MinIO)", "Вместо полных копий данные режут\nна куски и добавляют куски чётности.\nМеньше места, терпит отказ дисков.", 0.7, 4.3, 5.8, 2.2, LIGHT2)
    card2(s, "Зачем это здесь", "Шардирование тут автоматическое — просто кладём объекты. А репликацию и цену хранения выбираем сами.", 6.8, 4.3, 5.8, 2.2)
    takeaway(s, "Объектное хранилище — пример, где шардирование инкапсулировано, а репликация — явный выбор инженера.")
    page_num(s, 18)


def c_k8s(s):
    add_header(s, "Kubernetes: те же приёмы для приложения", "копирует не данные, а сами запущенные процессы")
    rows = [
        ("ReplicaSet", "держит заданное число копий приложения. Упала — поднимется новая. Это «репликация процесса»."),
        ("StatefulSet", "стабильные имена pod-0, pod-1 и свои диски. Нужно базе, где каждая копия — конкретный шард."),
        ("PodDisruptionBudget", "не даёт выселить слишком много копий за раз — иначе упадёт и кворум, и сервис."),
        ("anti-affinity", "раскидывает копии по разным машинам и зонам, чтобы одна поломка не убила все копии сразу."),
    ]
    ry = 1.9
    for name, how in rows:
        shape(s, 0.7, ry, 2.7, 1.0, fill=LIGHT, line_color=LINE, rounded=True)
        label(s, 0.7, ry, 2.7, 1.0, name, 12, True, INK)
        shape(s, 3.5, ry, 9.2, 1.0, fill=LIGHT2, line_color=LINE, rounded=True)
        text(s, 3.7, ry+0.14, 8.8, 0.75, how, 11.5, False, INK)
        ry += 1.15
    takeaway(s, "Kubernetes не знает про «шарды данных», но даёт ровно те примитивы, из которых шардированная база собирается.")
    page_num(s, 19)


def c_citus_prac(s):
    add_header(s, "Практика: шардирование видно глазами", "поднимаем Citus в docker-compose")
    node(s, 0.9, 1.6, 3.0, 0.95, "coordinator\n(docker-compose)", BLUE_D, WHITE, 10.5)
    xs = [0.9, 4.1, 7.3]
    for i in range(3):
        node(s, xs[i], 2.8, 2.6, 1.1, "worker %d" % (i+1), CYAN, INK, 11)
        arrow(s, 2.4, 2.55, xs[i]+1.3, 2.75, BLUE, 1.5)
    text(s, 8.2, 1.6, 4.4, 0.5, "Что делаем по шагам:", 14, True, INK)
    steps = ["1. Включаем coordinator + workers", "2. create_distributed_table", "3. Смотрим строки на каждом воркере", "4. Меняем ключ — видим перекос"]
    for i, st in enumerate(steps):
        text(s, 8.2, 2.15 + i*0.5, 4.6, 0.45, st, 11, False, INK)
    text(s, 0.9, 4.3, 7.8, 1.8,
         "Что это даёт: план запроса показывает, когда данные с одного ключа лежат рядом — база не гоняет их между машинами. Перекос видно цифрами.",
         11.5, False, INK)
    takeaway(s, "Сколько строк уехало на каждый worker — это и есть шардирование, измеренное в числах.")
    page_num(s, 20)


def c_patroni_prac(s):
    add_header(s, "Практика: RTO превращаем в секундомер", "засекаем, за сколько Patroni поднимет нового главного")
    steps = ["1. отключаем\nprimary", "2. истекает\nаренда (ttl)", "3. выборы\nнового", "4. новый\nprimary готов"]
    x = 0.7
    for i, st in enumerate(steps):
        last = i == len(steps)-1
        shape(s, x, 2.2, 2.6, 1.6, fill=BLUE if last else LIGHT2, line_color=LINE, rounded=True)
        label(s, x+0.08, 2.2, 2.44, 1.6, st, 11, True, WHITE if last else INK)
        if i < len(steps)-1:
            arrow(s, x+2.6, 3.0, x+3.0, 3.0, PINK, 2)
        x += 3.0
    text(s, 0.7, 4.2, 11.5, 0.8,
         "Засекаем: от «docker stop primary» до момента, когда новый главный готов. Ожидаемо 30–60 секунд.", 13.5, True, INK)
    text(s, 0.7, 5.0, 11.5, 1.2,
         "Дальше pg_rewind возвращает упавший узел в строй как копию. При синхронной копии потерянных данных ноль — RPO = 0.",
         11.5, False, MUTE)
    page_num(s, 21)


CONTENT = [
    c_scale, c_cap, c_evol, c_shard, c_strategies, c_repl, c_consist,
    c_partition, c_citus, c_replpg, c_patroni, c_net, c_example,
    c_pg, c_pg_citus, c_greenplum, c_spark, c_s3, c_k8s,
    c_citus_prac, c_patroni_prac,
]


def main():
    prs = Presentation(TEMPLATE)
    _patch_next_slide_partname(prs)

    # из 8 слайдов шаблона (0..7) оставляем 0,1,2,3,4 (титул, связь, правила, рекап×2) и 7 (рефлексия).
    # выкидываем 5 (оглавление) и 6 (раздел 1).
    delete_slides(prs, [5, 6])
    # теперь порядок sldIdLst: [0,1,2,3,4, рефлексия]

    layout_content = get_layout(prs, "1_Заголовок")

    for fn in CONTENT:
        s = prs.slides.add_slide(layout_content)
        fn(s)

    # переносим рефлексию (индекс 5) в самый конец
    total = len(prs.slides._sldIdLst)   # 6 + len(CONTENT)
    k = len(CONTENT)
    order = list(range(0, 5)) + list(range(6, total)) + [5]
    reorder_slides(prs, order)

    out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "презентация-шардирование-репликация.pptx")
    prs.save(out)
    print("Saved:", out, "| slides:", len(prs.slides._sldIdLst))


if __name__ == "__main__":
    main()
