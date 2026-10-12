#!/usr/bin/env python3
"""
warmlink_sync.py — automatyczny zapis okresów pracy pompy z chmury Warmlink do pompa_dane.json.

Co robi (w pętli, co --interval sekund, domyślnie 120, min. 60):
  1. loguje się na DRUGIE konto Warmlink (konto członka domu) i pobiera odczyty getDataByCode,
  2. dopisuje surowe odczyty do warmlink_log.jsonl (do analiz, np. rozpoznania CO/CWU),
  3. wykrywa pracę sprężarki (T30 > 0 = Hz) i buduje „okres COP” (pole `meas` w aplikacji),
     z odczytami T01, T02, T04, T39 (flow), T30 (hz), T34 (V), T35 (A),
  4. wstawia/aktualizuje okres w pompa_dane.json (src: "api"). Aplikacja na komputerze sama
     scala plik i wysyła zmiany przez Gist na telefon — Gista tu nie używamy.

Bezpieczeństwo danych:
  - plik jest czytany, zmieniany tylko w `meas` i zapisywany atomowo; reszta (cykle, PV, faktury…) bez zmian,
  - kopia zapasowa <plik>.bak raz na dobę,
  - okres edytowany przez Ciebie w aplikacji (np. wpisane kWh) przestaje być aktualizowany przez skrypt,
  - okres usunięty w aplikacji nie jest przywracany.
Zużycie energii (kWh) zostaje puste — API go nie podaje; uzupełnisz z licznika przyciskiem „Edytuj”.
Tryb CO/CWU: domyślnie --mode CO (do potwierdzenia); surowy log pozwoli to potem rozpoznać automatycznie.

Hasło: zmienna WARMLINK_PASS, albo Pęk kluczy macOS:
  security add-generic-password -a TWOJ_EMAIL -s warmlink_bridge -w
albo pytanie w terminalu. Login: --user lub WARMLINK_USER. Pełny kod pompy: --device-code lub WARMLINK_DEVICE.
"""
import argparse, getpass, json, os, random, shutil, subprocess, sys, tempfile, time, datetime as dt
from types import SimpleNamespace

import warmlink_bridge as wb

CODES = ['Power', 'Mode', 'R01', 'R02', 'R03', 'R08', 'R09', 'T01', 'T02', 'T03', 'T04', 'T05', 'T06', 'T07',
         'T08', 'T09', 'T10', 'T11', 'T12', 'T30', 'T34', 'T35', 'T39']
B36 = '0123456789abcdefghijklmnopqrstuvwxyz'


def num(x):
    try:
        return None if x is None or str(x).strip() == '' else float(x)
    except ValueError:
        return None


def uid():
    n, s = int(time.time() * 1000), ''
    while n: s = B36[n % 36] + s; n //= 36
    return s + ''.join(random.choice(B36) for _ in range(4))


def atomic_write(path, text):
    d = os.path.dirname(os.path.abspath(path))
    fd, tmp = tempfile.mkstemp(dir=d, prefix='.tmp_', suffix='.json')
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f: f.write(text)
        os.replace(tmp, path)
    except Exception:
        if os.path.exists(tmp): os.unlink(tmp)
        raise


def keychain(user):
    try:
        r = subprocess.run(['security', 'find-generic-password', '-s', 'warmlink_bridge', '-a', user, '-w'],
                           capture_output=True, text=True, timeout=10)
        return r.stdout.strip() if r.returncode == 0 and r.stdout.strip() else None
    except Exception:
        return None


# ---------- logika okresów (czysta, testowalna bez sieci) ----------

def to_reading(ts, vals, a):
    g = lambda c: num(vals.get(c))
    r = {'ts': ts, 't01': g('T01'), 't02': g('T02'), 't04': g('T04'), 'flow': g('T39'), 'hz': g('T30'),
         'volt': g('T34'), 'amp': g('T35'), 'c1': True, 'c2': True, 'def': False, 'k': None}
    if r['t01'] is not None and r['t02'] is not None and (r['hz'] or 0) > 0 and r['t02'] - r['t01'] < a.defrost_dt:
        r['def'] = True      # heurystyka: woda wychodząca zimniejsza od wchodzącej = odszranianie
    return r


def build_period(op, settings, a):
    rd = op['rd']; start, end = rd[0]['ts'], rd[-1]['ts']
    return {'id': op['id'], 'start': start, 'end': end, 'date': start[:10], 'mode': a.mode,
            'defMin': round(op['defMin']), 'kwh': None, 'rd': rd, 'res': {'T': 0.1, 'F': 0.1, 'E': 0.01},
            'rho': settings.get('rho', 1000), 'cpw': settings.get('cpw', 4.186),
            'dtOff': settings.get('dtOff', 0), 'useOff': False, 'hzMax': settings.get('hzMax', 100),
            'note': 'auto z Warmlink API — tryb i zużycie (kWh) do uzupełnienia', 'src': 'api'}


def _strip(o):
    return {k: v for k, v in o.items() if k != 'updatedAt'}


def upsert_state(st, op, settings, a, now_ms):
    if len(op['rd']) < 2: return
    obj = build_period(op, settings, a)
    p = st['periods'].setdefault(op['id'], {'obj': None, 'upd': 0, 'written_upd': None, 'written': False, 'disowned': False})
    if p['obj'] is None or _strip(p['obj']) != _strip(obj):
        obj['updatedAt'] = now_ms; p['obj'] = obj; p['upd'] = now_ms


def step(st, vals, now, a, settings):
    """Jeden odczyt -> aktualizacja stanu okresów. Zwraca opis do wydruku."""
    ts, ep, now_ms = now.strftime('%Y-%m-%dT%H:%M'), now.timestamp(), int(now.timestamp() * 1000)
    hz = num(vals.get('T30')); running = hz is not None and hz > 0
    op = st.get('open')
    if op and ep - op['lastRun'] > a.stop_after * 60:        # postój/przerwa dłuższa niż próg -> zamknij okres
        upsert_state(st, op, settings, a, now_ms)
        st['open'] = op = None
    if running:
        if not op:
            op = st['open'] = {'id': uid(), 'rd': [], 'defMin': 0.0, 'lastRun': ep}
        rd = to_reading(ts, vals, a)
        if op['rd'] and op['rd'][-1]['ts'] == ts: op['rd'][-1] = rd
        else: op['rd'].append(rd)
        op['lastRun'] = ep
        if rd['def']: op['defMin'] += a.interval / 60
        upsert_state(st, op, settings, a, now_ms)
        return f'sprężarka pracuje (T30={hz:g} Hz), okres {op["id"]}, odczytów: {len(op["rd"])}'
    return 'sprężarka stoi' + (f' (okres {op["id"]} otwarty, czekam {a.stop_after} min)' if op else '')


def load_settings(path):
    try:
        return json.load(open(path, encoding='utf-8')).get('settings') or {}
    except Exception:
        return {}


def sync_file(path, st, a, now_ms):
    """Wstawia/aktualizuje okresy API w pompa_dane.json. Zwraca True, jeśli plik zapisano."""
    try:
        data = json.load(open(path, encoding='utf-8'))
    except (json.JSONDecodeError, OSError) as e:
        print('  plik danych chwilowo nieczytelny, ponowię przy następnym odczycie:', e, flush=True); return False
    meas = data.setdefault('meas', [])
    tomb = {(t.get('type'), t.get('id')) for t in data.get('tomb', [])}
    by_id = {m.get('id'): i for i, m in enumerate(meas) if isinstance(m, dict)}
    changed = False
    for pid, p in st['periods'].items():
        if p['disowned'] or p['obj'] is None: continue
        if ('meas', pid) in tomb:
            p['disowned'] = True; print(f'  okres {pid} usunięty w aplikacji — nie przywracam', flush=True); continue
        i = by_id.get(pid)
        if i is None:
            meas.append(p['obj']); changed = True
        else:
            cur = meas[i]
            if p['written'] and cur.get('updatedAt') != p['written_upd']:
                p['disowned'] = True; print(f'  okres {pid} edytowany w aplikacji — przestaję go aktualizować', flush=True); continue
            if _strip(cur) != _strip(p['obj']) or cur.get('updatedAt') != p['obj']['updatedAt']:
                meas[i] = p['obj']; changed = True
        p['written'] = True; p['written_upd'] = p['obj']['updatedAt']
    if changed:
        if a.dry_run:
            print('  [dry-run] zapisałbym zmiany w', path, flush=True); return False
        bak = path + '.bak'
        if not os.path.exists(bak) or time.time() - os.path.getmtime(bak) > 86400:
            shutil.copy2(path, bak)
        atomic_write(path, json.dumps(data, ensure_ascii=False, indent=1))
    return changed


def prune(st, now):
    cut = (now - dt.timedelta(days=60)).strftime('%Y-%m-%d')
    for pid in [k for k, p in st['periods'].items()
                if p['obj'] and p['obj']['date'] < cut and (st.get('open') or {}).get('id') != pid]:
        del st['periods'][pid]


# ---------- chmura ----------

def get_values(c, code):
    for attempt in (0, 1):
        res = c.data_by_code(code, CODES)
        ok = isinstance(res, dict) and not res.get('_http_error') and str(res.get('error_code')) == '0' \
            and isinstance(res.get('objectResult'), list)
        if ok:
            return {r.get('code'): r.get('value') for r in res['objectResult'] if isinstance(r, dict)}
        if attempt == 0: c.login()           # token wygasł / konto wylogowane
    raise RuntimeError('odczyt nieudany: ' + json.dumps(wb.redact(res), ensure_ascii=False)[:300])


def load_state(path):
    try:
        st = json.load(open(path, encoding='utf-8'))
        st.setdefault('periods', {}); st.setdefault('open', None); return st
    except Exception:
        return {'open': None, 'periods': {}}


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('--data', required=True, help='ścieżka do pompa_dane.json (ten sam plik, który wskazałeś w aplikacji)')
    p.add_argument('--user', default=os.environ.get('WARMLINK_USER'))
    p.add_argument('--device-code', default=os.environ.get('WARMLINK_DEVICE'), help='PEŁNY kod urządzenia')
    p.add_argument('--interval', type=int, default=120); p.add_argument('--stop-after', type=int, default=10,
                   help='po ilu minutach bez pracy sprężarki zamknąć okres (domyślnie 10)')
    p.add_argument('--mode', choices=['CO', 'CWU'], default='CO')
    p.add_argument('--defrost-dt', type=float, default=-0.5, help='T02−T01 poniżej tej wartości = odszranianie')
    p.add_argument('--state', default='warmlink_state.json'); p.add_argument('--log', default='warmlink_log.jsonl')
    p.add_argument('--no-raw-log', action='store_true'); p.add_argument('--once', action='store_true')
    p.add_argument('--dry-run', action='store_true', help='nic nie zapisuj w pliku danych')
    p.add_argument('--cafile'); p.add_argument('--pin-sha256'); p.add_argument('--app-id', default='16')
    a = p.parse_args()
    if a.interval < 60: a.interval = 60; print('Interwał podniesiony do 60 s (minimum).')
    if not a.user: sys.exit('Brak loginu: --user lub WARMLINK_USER.')
    if not a.device_code: sys.exit('Brak pełnego kodu urządzenia: --device-code lub WARMLINK_DEVICE.')
    if not os.path.isfile(a.data): sys.exit(f'Nie ma pliku danych: {a.data}')
    pw = os.environ.get('WARMLINK_PASS') or keychain(a.user) or (getpass.getpass('Hasło: ') if sys.stdin.isatty() else None)
    if not pw: sys.exit('Brak hasła: WARMLINK_PASS, Pęk kluczy (patrz opis) lub uruchom w terminalu.')

    tls = SimpleNamespace(insecure=False, pin_sha256=a.pin_sha256, cafile=a.cafile)
    c = wb.Cloud(wb.BASES['warmlink'], a.user, pw, 'try', a.app_id, opener=wb.build_opener(tls))
    try:
        c.login()
    except wb.CertError as e:
        sys.exit('\n' + str(e))
    print('Zalogowano. Start pętli, interwał', a.interval, 's.', flush=True)
    st, fails = load_state(a.state), 0
    while True:
        now = dt.datetime.now()
        try:
            vals = get_values(c, a.device_code)
            if not a.no_raw_log:
                with open(a.log, 'a', encoding='utf-8') as f:
                    f.write(json.dumps({'ts': now.isoformat(timespec='seconds'), 'v': {k: vals.get(k) for k in CODES}}, ensure_ascii=False) + '\n')
            msg = step(st, vals, now, a, load_settings(a.data))
            wrote = sync_file(a.data, st, a, int(now.timestamp() * 1000))
            prune(st, now); atomic_write(a.state, json.dumps(st, ensure_ascii=False))
            print(f'[{now:%H:%M:%S}] {msg}' + (' -> zapisano w pliku danych' if wrote else ''), flush=True)
            fails = 0
        except wb.CertError as e:
            sys.exit('\n' + str(e))
        except Exception as e:
            fails += 1; print(f'[{now:%H:%M:%S}] Błąd ({fails}): {e}', flush=True)
        if a.once: break
        time.sleep(min(a.interval * (1 + min(fails, 5)), 900) if fails else a.interval)


if __name__ == '__main__':
    main()
