#!/usr/bin/env python3
"""
warmlink_bridge.py — nieoficjalny pomost do chmury Warmlink / Linked-Go (tylko biblioteka standardowa).

STATUS: napisane na podstawie opisów społeczności (m.in. issue #99 w radical-squared/aquatemp,
ha-warmlink). NIE było testowane na prawdziwym koncie ani na aplikacji Warmlink 3.0.8 — dlatego
pierwszym krokiem jest tryb `probe`, który tylko CZYTA i zapisuje surowe odpowiedzi chmury.

Tryby:
  probe   logowanie -> lista urządzeń -> status -> zapis surowych odpowiedzi do probe_output.json
          (token, login i e-mail są zamazane; plik można bezpiecznie pokazać)
  certinfo  diagnostyka TLS: pokazuje łańcuch certyfikatów serwera (bez wysyłania hasła)
  log     co N sekund (min. 60) dopisuje surowy status do warmlink_log.jsonl (1 linia = 1 odczyt)

Dane logowania: zmienne środowiskowe WARMLINK_USER / WARMLINK_PASS albo pytanie w terminalu.
Hasło nigdzie nie jest zapisywane.

WAŻNE: Warmlink dopuszcza jedno aktywne logowanie na konto. Załóż drugie konto, udostępnij mu
"dom" w aplikacji i używaj go tutaj (konto członka dostaje pustą listę urządzeń — wtedy podaj
--device-code z ekranu informacji o urządzeniu w aplikacji).
Korzystanie z nieoficjalnego API może naruszać regulamin producenta — używaj na własną odpowiedzialność
i nie odpytuj częściej niż co 60 s.
"""
import argparse, getpass, hashlib, http.client, json, os, socket, ssl, sys, tempfile, time
import urllib.request, urllib.error, urllib.parse, datetime as dt

BASES = {
    # Warmlink (aplikacja 2.x/3.x wg opisu w issue #99): port 449, ścieżka /crmservice
    'warmlink': 'https://cloud.linked-go.com:449/crmservice/api/app',
    # starsze konta Aquatemp (api level 1)
    'aquatemp': 'https://cloud.linked-go.com/cloudservice/api/app',
}
SUSPECT = ('kaspersky', 'eset', 'avast', 'avg', 'bitdefender', 'zscaler', 'fortinet', 'cisco', 'sophos',
           'norton', 'mcafee', 'mitmproxy', 'charles', 'fiddler', 'netskope', 'palo alto', 'checkpoint', 'proxy',
           'firewall', 'vpn', 'adguard', 'trend micro', 'symantec', 'webroot', 'g data', 'panda')

CERT_HELP = """Weryfikacja certyfikatu TLS nie powiodła się:
  {reason}

Połączenie przerwano PRZED wysłaniem loginu i hasła — nic nie wyciekło.
Możliwe przyczyny:
  1) coś po drodze podmienia certyfikat: antywirus skanujący HTTPS, VPN, proxy firmowe, router,
  2) serwer producenta używa własnego CA / certyfikatu z własnym podpisem,
  3) (Mac, Python z python.org) brak certyfikatów systemowych — uruchom plik
     "/Applications/Python 3.x/Install Certificates.command" (przy błędzie 'self-signed' mniej prawdopodobne).

Jak to sprawdzić (certinfo nie wysyła żadnych danych logowania):
  python3 warmlink_bridge.py certinfo
  - jeśli wystawca wygląda na antywirus/VPN/proxy -> wyłącz skanowanie HTTPS lub spróbuj z hotspotu telefonu,
  - jeśli wystawca to ktoś nieznany/producent i odcisk jest taki sam z różnych sieci -> serwer ma własny cert.

Bezpieczne obejścia (od najlepszego):
  --cafile PLIK.pem        ufaj tylko temu CA (np. root antywirusa/firmy wyeksportowany do PEM)
  --pin-sha256 ODCISK      ufaj tylko certyfikatowi serwera o tym odcisku (odcisk poda `certinfo`)
  --insecure               BEZ weryfikacji — każdy w sieci może przechwycić hasło. Tylko ostateczność,
                           najlepiej na haśle zmienionym/jednorazowym."""


class CertError(RuntimeError):
    pass


SECRET_KEYS = ('token', 'password', 'passwd', 'email', 'user', 'phone', 'mobile', 'nick', 'address', 'lat', 'lon',
               'accesskey', 'iccid', 'devicename', 'device_name', 'houseid', 'house_id', 'deviceid', 'device_id')
SECRET_EXACT = ('sn',)


def redact(o):
    if isinstance(o, dict):
        out = {}
        for k, v in o.items():
            kl = str(k).lower()
            if (kl in SECRET_EXACT or any(s in kl for s in SECRET_KEYS)) and not isinstance(v, (dict, list)):
                out[k] = '***'
            elif 'device_code' in kl or 'devicecode' in kl:
                s = str(v); out[k] = (s[:4] + '…' + s[-4:]) if len(s) > 10 else '***'
            else:
                out[k] = redact(v)
        return out
    if isinstance(o, list):
        return [redact(x) for x in o]
    return o


def norm_fp(x):
    return x.replace(':', '').replace(' ', '').lower()


class PinnedHandler(urllib.request.HTTPSHandler):
    """HTTPS bez weryfikacji łańcucha, ale tylko dla certyfikatu serwera o zadanym SHA-256."""
    def __init__(self, pin):
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT); ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE
        super().__init__(context=ctx); self.pin = norm_fp(pin)

    def https_open(self, req):
        pin = self.pin
        class Conn(http.client.HTTPSConnection):
            def connect(self):
                super().connect()
                fp = hashlib.sha256(self.sock.getpeercert(True)).hexdigest()
                if fp != pin:
                    self.sock.close()
                    raise ssl.SSLCertVerificationError(f'odcisk certyfikatu serwera ({fp}) != podany --pin-sha256')
        return self.do_open(Conn, req, context=self._context)


def build_opener(a):
    if a.insecure:
        print('!' * 70, '\nUWAGA: --insecure wyłącza weryfikację certyfikatu. Login i hasło mogą zostać\n'
              'przechwycone przez każdego, kto jest „po drodze”. Użyj tylko, jeśli rozumiesz ryzyko.\n' + '!' * 70)
        if not sys.stdin.isatty() or input('Wpisz TAK, aby kontynuować: ').strip() != 'TAK':
            sys.exit('Przerwano (brak potwierdzenia).')
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT); ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE
        return urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))
    if a.pin_sha256:
        return urllib.request.build_opener(PinnedHandler(a.pin_sha256))
    ctx = ssl.create_default_context(cafile=a.cafile) if a.cafile else ssl.create_default_context()
    return urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))


class Cloud:
    def __init__(self, base, user, password, hash_mode='try', app_id='16', timeout=20, opener=None):
        self.base, self.user, self.password, self.hash_mode = base.rstrip('/'), user, password, hash_mode
        self.app_id, self.timeout, self.token = app_id, timeout, None
        self.opener = opener or urllib.request.build_opener()

    def _post(self, path, body, token=None, lang='en'):
        req = urllib.request.Request(f'{self.base}/{path}?lang={lang}', data=json.dumps(body).encode(),
                                     method='POST', headers={'Content-Type': 'application/json; charset=utf-8'})
        if token: req.add_header('x-token', token)
        try:
            with self.opener.open(req, timeout=self.timeout) as r:
                return json.loads(r.read().decode('utf-8', 'replace'))
        except urllib.error.HTTPError as e:
            return {'_http_error': e.code, '_body': e.read().decode('utf-8', 'replace')[:500]}
        except urllib.error.URLError as e:
            if isinstance(e.reason, ssl.SSLError):
                r = e.reason
                msg = getattr(r, 'verify_message', None) or (r.args[0] if r.args else r)
                raise CertError(CERT_HELP.format(reason=msg)) from None
            raise

    def login(self):
        md5 = hashlib.md5(self.password.encode()).hexdigest()
        variants = {'md5': [md5], 'plain': [self.password], 'try': [md5, self.password]}[self.hash_mode]
        last = None
        for pw in variants:
            body = {'userName': self.user, 'password': pw, 'loginSource': 'IOS', 'areaCode': 'en',
                    'appId': self.app_id, 'type': '2'}
            res = self._post('user/login', body); last = res
            tok = (res.get('objectResult') or {}).get('x-token') if isinstance(res.get('objectResult'), dict) else None
            if tok:
                self.token = tok; return res
        raise RuntimeError('Logowanie nie powiodło się: ' + json.dumps(redact(last), ensure_ascii=False)[:300])

    def device_list(self):
        return self._post('device/deviceList', {'appId': self.app_id}, self.token)

    def device_status(self, code):
        return self._post('device/getDeviceStatus', {'appId': self.app_id, 'deviceCode': code}, self.token)

    def data_by_code(self, code, codes):
        return self._post('device/getDataByCode', {'appId': self.app_id, 'deviceCode': code, 'protocalCodes': codes}, self.token)


def first_device_code(lst):
    rows = lst.get('objectResult') if isinstance(lst, dict) else None
    if isinstance(rows, list) and rows:
        r = rows[0]
        return r.get('device_code') or r.get('deviceCode')
    return None


def make_cloud(a):
    user = a.user or os.environ.get('WARMLINK_USER') or input('Login (e-mail): ')
    pw = os.environ.get('WARMLINK_PASS') or getpass.getpass('Hasło: ')
    return Cloud(a.base or BASES[a.api], user, pw, a.hash, a.app_id, opener=build_opener(a))


def _fmt_name(t):
    return ', '.join(f'{k}={v}' for rdn in t for k, v in rdn) if t else '?'


def _decode_pem(pem):
    try:
        with tempfile.NamedTemporaryFile('w', suffix='.pem', delete=False) as f:
            f.write(pem); path = f.name
        try:
            return ssl._ssl._test_decode_cert(path)
        finally:
            os.unlink(path)
    except Exception:
        return {}


def cmd_certinfo(a):
    """Diagnostyka: pobiera łańcuch certyfikatów serwera BEZ wysyłania jakichkolwiek danych logowania."""
    u = urllib.parse.urlparse(a.base or BASES[a.api]); host, port = u.hostname, u.port or 443
    print(f'Serwer: {host}:{port}\n')
    try:
        ctx = ssl.create_default_context(cafile=a.cafile) if a.cafile else ssl.create_default_context()
        with socket.create_connection((host, port), timeout=15) as s, ctx.wrap_socket(s, server_hostname=host):
            print('Weryfikacja standardowa: OK — certyfikat jest zaufany, problem nie występuje z tej sieci.\n')
    except ssl.SSLCertVerificationError as e:
        print('Weryfikacja standardowa: BŁĄD —', getattr(e, 'verify_message', e), '\n')
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT); ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE
    with socket.create_connection((host, port), timeout=15) as s, ctx.wrap_socket(s, server_hostname=host) as ss:
        chain = ss.get_unverified_chain() if hasattr(ss, 'get_unverified_chain') else None
        ders = [c if isinstance(c, bytes) else c.public_bytes(ssl._ssl.ENCODING_DER) for c in (chain or [])]
        if not ders: ders = [ss.getpeercert(True)]
    pems = []
    for i, der in enumerate(ders):
        pem = ssl.DER_cert_to_PEM_cert(der); pems.append(pem); d = _decode_pem(pem)
        subj, iss = _fmt_name(d.get('subject')), _fmt_name(d.get('issuer'))
        print(f'[{i}] {"serwer (leaf)" if i == 0 else "pośredni/root"}')
        print(f'    podmiot : {subj}\n    wystawca: {iss}')
        if d: print(f'    ważny   : {d.get("notBefore")} -> {d.get("notAfter")}')
        print(f'    SHA-256 : {hashlib.sha256(der).hexdigest()}')
        if d and d.get('subject') == d.get('issuer'): print('    ! certyfikat z własnym podpisem (root)')
        hit = [k for k in SUSPECT if k in (subj + ' ' + iss).lower()]
        if hit: print(f'    ! wystawca wygląda na oprogramowanie pośredniczące ({", ".join(hit)})')
        print()
    with open(a.chain_out, 'w') as f: f.write(''.join(pems))
    leaf = hashlib.sha256(ders[0]).hexdigest()
    print(f'Zapisano łańcuch do {a.chain_out}.\n\nJak interpretować:\n'
          '  - wystawca = antywirus/VPN/proxy/router -> ruch jest przechwytywany lokalnie; porównaj z hotspotem telefonu\n'
          '    (inny odcisk w innej sieci = potwierdzony „man in the middle”),\n'
          '  - ten sam odcisk ze wszystkich sieci -> serwer sam używa własnego CA.\n'
          'Jeśli ufasz temu certyfikatowi, uruchom probe z przypięciem:\n'
          f'  python3 warmlink_bridge.py probe --pin-sha256 {leaf}')


def cmd_probe(a):
    c = make_cloud(a); out = {'when': dt.datetime.now().isoformat(timespec='seconds'), 'base': c.base, 'steps': {}}
    out['steps']['login'] = redact(c.login())
    lst = c.device_list(); out['steps']['deviceList'] = redact(lst)
    code = a.device_code or first_device_code(lst)
    if not code:
        out['note'] = 'Brak urządzenia na liście (konto członka?). Podaj --device-code z aplikacji.'
    else:
        out['steps']['getDeviceStatus'] = redact(c.device_status(code))
        if a.codes:
            out['steps']['getDataByCode'] = redact(c.data_by_code(code, a.codes.split(',')))
    with open(a.out, 'w', encoding='utf-8') as f: json.dump(out, f, ensure_ascii=False, indent=1)
    print('Zapisano', a.out, '— pokaż ten plik (jest zamazany), a dopasuję pola do aplikacji.')


def cmd_log(a):
    if a.interval < 60: a.interval = 60; print('Interwał podniesiony do 60 s (minimum).')
    c = make_cloud(a); c.login(); code = a.device_code or first_device_code(c.device_list())
    if not code: sys.exit('Brak kodu urządzenia — użyj --device-code.')
    n = 0
    while True:
        try:
            st = c.device_status(code)
            if isinstance(st, dict) and (st.get('error_code') not in (None, '0', 0) or st.get('_http_error')):
                c.login(); st = c.device_status(code)           # token wygasł / wylogowany przez aplikację
            with open(a.out, 'a', encoding='utf-8') as f:
                f.write(json.dumps({'ts': dt.datetime.now().isoformat(timespec='seconds'), 'raw': st}, ensure_ascii=False) + '\n')
            n += 1; print(f'[{n}] zapisano odczyt', flush=True)
        except CertError:
            raise
        except Exception as e:
            print('Błąd:', e, flush=True)
        if a.once: break
        time.sleep(a.interval)


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('mode', choices=['probe', 'log', 'certinfo'])
    p.add_argument('--api', choices=list(BASES), default='warmlink'); p.add_argument('--base', help='własny adres bazowy (testy)')
    p.add_argument('--user'); p.add_argument('--hash', choices=['try', 'md5', 'plain'], default='try')
    p.add_argument('--app-id', default='16'); p.add_argument('--device-code')
    p.add_argument('--codes', help='lista kodów po przecinku do getDataByCode, np. T01,T02,T04,T30,T39')
    p.add_argument('--interval', type=int, default=300); p.add_argument('--once', action='store_true')
    tls = p.add_mutually_exclusive_group()
    tls.add_argument('--cafile', help='plik PEM z zaufanym CA (np. root antywirusa/firmy)')
    tls.add_argument('--pin-sha256', help='ufaj tylko certyfikatowi serwera o tym odcisku SHA-256 (z `certinfo`)')
    tls.add_argument('--insecure', action='store_true', help='WYŁĄCZ weryfikację certyfikatu (niebezpieczne)')
    p.add_argument('--chain-out', default='cert_chain.pem', help='gdzie zapisać łańcuch w trybie certinfo')
    p.add_argument('--out'); a = p.parse_args()
    a.out = a.out or ('probe_output.json' if a.mode == 'probe' else 'warmlink_log.jsonl')
    try:
        {'probe': cmd_probe, 'log': cmd_log, 'certinfo': cmd_certinfo}[a.mode](a)
    except CertError as e:
        sys.exit('\n' + str(e))
    except (urllib.error.URLError, OSError) as e:
        sys.exit(f'\nBłąd połączenia: {e}\nSkrypt łączy się z portem 449 — spróbuj z innej sieci (hotspot telefonu).')


if __name__ == '__main__':
    main()
