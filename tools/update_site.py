"""به‌روزرسانی روزانه داشبورد: اکسل‌های پوشه input را می‌خواند، دسته‌بندی می‌کند، رمزگذاری می‌کند و docs/data.enc را می‌نویسد."""
import os, sys, json, glob, base64, getpass, csv, datetime, warnings
from collections import defaultdict
import openpyxl
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from rules import *
warnings.filterwarnings('ignore')
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INP, DOCS = os.path.join(BASE, 'input'), os.path.join(BASE, 'docs')
CATS = [C1, C2, C3, C4, C5, C6, C7, C8, C9, C10]
SH = {C1: 'حقوق و پرسنلی', C2: 'بیمه و مالیات', C3: 'پس‌انداز و کسورات', C4: 'سنوات و ۴٪ و طبقه‌بندی', C5: 'پیمان',
      C6: 'قبوض و سایر', C7: 'بیمه تکمیلی', C8: 'بن کارگری', C9: 'خرید', C10: 'بودجه مدیریت'}
MN = ['فروردین', 'اردیبهشت', 'خرداد', 'تیر', 'مرداد', 'شهریور', 'مهر', 'آبان', 'آذر', 'دی', 'بهمن', 'اسفند']
ITER = 200000

def num(v):
    try: return float(v or 0)
    except Exception: return 0.0

files = [f for f in glob.glob(os.path.join(INP, '*.xlsx')) if not os.path.basename(f).startswith('~$')]
if not files: sys.exit('هیچ فایل اکسلی در پوشه input پیدا نشد.')
budget, seen, tx = {}, set(), []
for f in sorted(files):
    wb = openpyxl.load_workbook(f, data_only=True)
    if 'Fact_Budget' in wb.sheetnames:
        for r in wb['Fact_Budget'].iter_rows(min_row=2, values_only=True):
            if r[0] and r[1]: budget[(str(r[0]), r[1])] = (num(r[2]), num(r[3]), num(r[4]))
    for sn in wb.sheetnames:
        if not sn.startswith('ریز تراکنش'): continue
        rows = wb[sn].iter_rows(values_only=True); hd = {h: i for i, h in enumerate(next(rows)) if h}
        need = ['تاریخ قلم', 'طرف مقابل', 'بستانکار', 'شرح قلم']
        if any(n not in hd for n in need): print('ستون‌های لازم در', sn, 'نیست؛ رد شد'); continue
        g = lambda r, n: r[hd[n]] if n in hd and hd[n] < len(r) else None
        for r in rows:
            if not g(r, 'تاریخ قلم'): continue
            key = (g(r, 'تاریخ قلم'), g(r, 'طرف مقابل'), g(r, 'بستانکار'), g(r, 'شرح قلم'), g(r, 'مانده در خط'))
            if key in seen: continue
            seen.add(key); tx.append((str(g(r, 'تاریخ قلم')), g(r, 'طرف مقابل') or 'بدون نام', num(g(r, 'بستانکار')),
                                     g(r, 'شرح قلم') or '', g(r, 'دسته بندی'), g(r, 'سطح اطمینان')))
acts = defaultdict(lambda: defaultdict(float)); top = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
daily = defaultdict(lambda: defaultdict(lambda: defaultdict(float))); rev = defaultdict(list); cnt = defaultdict(int)
review_csv, last = [], ''
for date, cp, amt, desc, cat, lv in tx:
    if amt <= 0: continue
    if not cat: cat, auto_lv, _, _ = classify(desc, cp); lv = lv or auto_lv
    if cat == TR: continue
    mk = date[:7]; last = max(last, date)
    bad = cat not in CATS
    if bad or lv == Q:
        rev[mk].append((date, cp[:28], desc[:40], round(amt / 1e9, 3), SH.get(cat, 'نامشخص')))
        review_csv.append([date, cp, desc, amt, cat if not bad else 'نامشخص', lv])
    if bad: continue
    k = SH[cat]; cnt[mk] += 1
    acts[mk][cat] += amt; top[mk][cp][k] += amt; daily[mk][int(date[8:10])][k] += amt
months = sorted({m for m, _ in budget} | set(acts))
out = {'updated': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'), 'lastTx': last, 'order': months, 'months': {}}
for mk in months:
    cats = []
    for c in CATS:
        p, rq, al = budget.get((mk, c), (0, 0, 0))
        cats.append({'k': SH[c], 'a': al, 'q': rq, 'p': p, 'x': round(acts[mk][c] / 1e9, 2)})
    tp = sorted(top[mk].items(), key=lambda kv: -sum(kv[1].values()))[:8]
    rv = sorted(rev[mk], key=lambda x: -x[3])
    out['months'][mk] = {'label': f'{MN[int(mk[5:7]) - 1]} {mk[:4]}', 'cats': cats,
        'top': [[n[:30], {k: round(v / 1e9, 2) for k, v in m.items()}] for n, m in tp],
        'daily': {str(d): {k: round(v / 1e9, 2) for k, v in m.items()} for d, m in daily[mk].items()},
        'rev': len(rv), 'revamt': round(sum(x[3] for x in rv), 2), 'rvList': rv[:40], 'n': cnt[mk]}
with open(os.path.join(BASE, 'review.csv'), 'w', newline='', encoding='utf-8-sig') as fh:
    csv.writer(fh).writerows([['تاریخ', 'طرف مقابل', 'شرح', 'مبلغ', 'دسته', 'سطح']] + review_csv)
pw = os.environ.get('BUDGET_SITE_PASSWORD') or getpass.getpass('رمز سایت (حداقل ۱۰ نویسه؛ همان رمزی که مدیران وارد می‌کنند): ')
if len(pw) < 10: sys.exit('رمز کوتاه است؛ حداقل ۱۰ نویسه.')
salt, iv = os.urandom(16), os.urandom(12)
key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(pw.encode())
ct = AESGCM(key).encrypt(iv, json.dumps(out, ensure_ascii=False, separators=(',', ':')).encode(), None)
b = lambda x: base64.b64encode(x).decode()
os.makedirs(DOCS, exist_ok=True)
json.dump({'v': 1, 'it': ITER, 's': b(salt), 'i': b(iv), 'c': b(ct)}, open(os.path.join(DOCS, 'data.enc'), 'w'))
open(os.path.join(DOCS, '.nojekyll'), 'a').close()
print(f'انجام شد: {len(months)} ماه، {sum(cnt.values())} ردیف، آخرین تراکنش {last}، اقلام نیازمند بررسی: {len(review_csv)} (review.csv)')
