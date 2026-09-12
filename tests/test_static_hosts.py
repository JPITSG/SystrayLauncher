"""Compile production static-host logic on Linux with socket outcomes mocked.

Run: python3 -m unittest discover -s tests -p test_static_hosts.py -v
Requires cc. Windows/WebView2 are not run by this test.
"""
import pathlib
import re
import subprocess
import tempfile
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[1]
SOURCE = (ROOT / 'SystrayLauncher.c').read_text()


def function(name):
    match = re.search(r'static [^;{}]+\b' + name + r'\([^;{}]*\)\s*\{', SOURCE)
    end, depth = match.end(), 1
    while depth:
        depth += (SOURCE[end] == '{') - (SOURCE[end] == '}')
        end += 1
    return SOURCE[match.start():end]


class StaticHostTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.directory = tempfile.TemporaryDirectory()
        cls.addClassCleanup(cls.directory.cleanup)
        path = pathlib.Path(cls.directory.name)
        prelude = r'''
#include <assert.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>
#include <wctype.h>
#include <arpa/inet.h>
#include <netdb.h>
typedef int BOOL, LONG, SOCKET;
typedef uint32_t DWORD;
typedef uint64_t ULONGLONG;
typedef void *HANDLE, *LPVOID;
typedef struct in_addr IN_ADDR;
typedef struct in6_addr IN6_ADDR;
#define TRUE 1
#define FALSE 0
#define INVALID_SOCKET (-1)
#define SD_BOTH SHUT_RDWR
#define WINAPI
#define WM_APP_HOST_ROUTE_CHANGED 1
#define InetPtonA inet_pton
#define InetNtopA inet_ntop
#define DebugPrint(...) ((void)0)
#define swprintf_s swprintf
#define wcscpy_s(p,n,s) wcscpy(p,s)
#define wcscat_s(p,n,s) wcscat(p,s)
static int InetPtonW(int family, const wchar_t *text, void *out) {
    char ascii[64]; size_t n = wcslen(text); assert(n < sizeof(ascii));
    for (size_t i = 0; i <= n; ++i) ascii[i] = (char)text[i];
    return inet_pton(family, ascii, out);
}
static int g_hostProxyLock, lockDepth, g_hwnd = 1, titleChanges, closed, shut;
static ULONGLONG ticks = 1000;
static ULONGLONG GetTickCount64(void) { return ticks; }
static void EnterCriticalSection(int *lock) { (void)lock; ++lockDepth; }
static void LeaveCriticalSection(int *lock) { (void)lock; assert(--lockDepth == 0); }
static LONG InterlockedCompareExchange(volatile LONG *p, LONG value, LONG expected) {
    LONG old = *p; if (old == expected) *p = value; return old;
}
static LONG InterlockedExchange(volatile LONG *p, LONG value) {
    LONG old = *p; *p = value; return old;
}
static void InterlockedIncrement(volatile LONG *p) { ++*p; }
static void InterlockedDecrement(volatile LONG *p) { --*p; }
static void PostMessageW(int hwnd, int msg, int w, int l) {
    (void)hwnd; (void)msg; (void)w; (void)l; ++titleChanges;
}
static int closesocket(SOCKET socket) { assert(socket != INVALID_SOCKET); ++closed; return 0; }
static int mockShutdown(SOCKET socket, int how) { (void)socket; (void)how; ++shut; return 0; }
#define shutdown mockShutdown
static DWORD (*pendingThread)(LPVOID);
static LPVOID pendingTask;
static BOOL threadFails;
static HANDLE CreateThread(void *sa, size_t stack, DWORD (*fn)(LPVOID),
    LPVOID param, DWORD flags, void *id) {
    (void)sa; (void)stack; (void)flags; (void)id; assert(!lockDepth);
    if (threadFails) return NULL;
    assert(!pendingTask); pendingTask = param; pendingThread = fn; return (HANDLE)1;
}
static void CloseHandle(HANDLE h) { (void)h; }
static struct { wchar_t staticHostMappings[2048]; BOOL useStaticHostMappings; } g_config;
static BOOL g_hostProxyDnsFallback;
static unsigned short g_hostProxyPort = 12345;
static volatile LONG g_hostProxyStopping, g_hostProxyWorkerCount;
static char attempts[64];
static BOOL reachable[3], dnsWorks = TRUE;
static void attempted(char c) {
    assert(!lockDepth); size_t n = strlen(attempts); assert(n + 1 < sizeof(attempts));
    attempts[n] = c; attempts[n + 1] = 0;
}
static SOCKET HostProxyConnectWithTimeout(const struct addrinfo *address, DWORD timeout) {
    assert(timeout == 800);  // Real numeric getaddrinfo, mocked TCP result.
    char numeric[64];
    assert(getnameinfo(address->ai_addr, address->ai_addrlen, numeric,
        sizeof(numeric), NULL, 0, NI_NUMERICHOST) == 0);
    const char *known[] = { "1.2.3.4", "3.4.5.6", "2001:db8::1" };
    for (int i = 0; i < 3; ++i) if (strcmp(numeric, known[i]) == 0) {
        attempted('A' + i); return reachable[i] ? 100 + i : INVALID_SOCKET;
    }
    assert(0); return INVALID_SOCKET;
}
static SOCKET HostProxyConnectViaDns(const char *host, unsigned short port,
    char *used, size_t size) {
    assert(strcmp(host, "domain.com") == 0 && port == 443);
    assert(g_hostProxyDnsFallback); attempted('D');
    if (dnsWorks) snprintf(used, size, "9.9.9.9");
    return dnsWorks ? 200 : INVALID_SOCKET;
}
'''
        types = SOURCE[SOURCE.index('typedef enum {\n    HOST_PROXY_UNTESTED'):
                       SOURCE.index('static BOOL g_winsockInitialized')]
        types += re.search(r'typedef struct \{\n    int mappingIndex;[^}]+\} HostProxyProbeTask;', SOURCE)[0]
        constants = '\n'.join(re.findall(r'^#define HOST_PROXY_.*$', SOURCE, re.M))
        globals_ = '''
static HostProxyMapping *g_hostProxyMappings;
static size_t g_hostProxyMappingCount;
static HostProxyTunnel *g_hostProxyTunnelList;
'''
        names = [
            'IsOriginListSeparator', 'IsValidStaticHostName', 'IsValidStaticIpAddress',
            'IsValidStaticHostMapping', 'FreeHostProxyMappings', 'ParseStaticHostProxyMappings',
            'HostProxyRequired', 'BuildHostProxyPacScript', 'BuildStaticHostBrowserArguments',
            'HostProxyConnectMapped', 'HostProxyConnectMappedRange',
            'CloseHostTunnelsForMapping', 'HostProxyUpdateRoute', 'HostProxyProbeThread',
            'HostProxyStartProbeIfDue', 'HostProxyEstablishUpstream',
            'HostProxyExpireFallbackCooldowns', 'FindHostProxyMapping',
        ]
        main = r'''
static void clearMappings(void) {
    FreeHostProxyMappings(g_hostProxyMappings, g_hostProxyMappingCount);
    g_hostProxyMappings = NULL; g_hostProxyMappingCount = 0;
}
static void configure(const wchar_t *text) {
    clearMappings(); wcscpy(g_config.staticHostMappings, text);
    assert(ParseStaticHostProxyMappings());
}
static void resetRoute(void) {
    assert(!pendingTask && !g_hostProxyWorkerCount);
    HostProxyMapping *m = &g_hostProxyMappings[0];
    m->state = HOST_PROXY_UNTESTED; m->activeAddress = 0;
    m->routeGeneration = 0; m->lastProbeTick = 0; m->probeInFlight = FALSE;
    strcpy(m->currentAddress, m->addresses[0].address);
    memset(reachable, 0, sizeof(reachable)); attempts[0] = 0;
    g_hostProxyDnsFallback = TRUE; dnsWorks = TRUE; ticks = 1000;
    g_hostProxyStopping = FALSE; titleChanges = shut = closed = 0;
}
static SOCKET connectRequest(const char *expected) {
    attempts[0] = 0;
    HostProxyTunnel tunnel = { .client = 1, .upstream = INVALID_SOCKET, .mappingIndex = -1 };
    SOCKET result = HostProxyEstablishUpstream(&tunnel, 0, 443);
    assert(strcmp(attempts, expected) == 0);
    assert(tunnel.mappingIndex == 0 && tunnel.upstream == result);
    return result;
}
static void finishProbe(const char *expected) {
    assert(pendingTask && g_hostProxyWorkerCount == 1);
    LPVOID task = pendingTask; pendingTask = NULL; attempts[0] = 0;
    pendingThread(task);
    assert(strcmp(attempts, expected) == 0);
    assert(!g_hostProxyWorkerCount && !g_hostProxyMappings[0].probeInFlight);
}
static void testParser(void) {
    configure(L"DOMAIN.com:1.2.3.4\nother.test:127.0.0.1,domain.com:3.4.5.6;"
              L"domain.com:1.2.3.4 domain.com:[2001:0DB8:0:0::1],domain.com:[2001:db8::1]");
    assert(g_hostProxyMappingCount == 2);
    HostProxyMapping *m = &g_hostProxyMappings[0];
    assert(strcmp(m->host, "domain.com") == 0 && m->addressCount == 3);
    assert(strcmp(m->addresses[0].address, "1.2.3.4") == 0);
    assert(strcmp(m->addresses[1].address, "3.4.5.6") == 0);
    assert(strcmp(m->addresses[2].address, "2001:db8::1") == 0);
    assert(m->addresses[2].addressFamily == AF_INET6);
    assert(strcmp(m->currentAddress, "1.2.3.4") == 0);
    assert(FindHostProxyMapping("domain.com") == 0);
    assert(FindHostProxyMapping("other.test") == 1);
    assert(FindHostProxyMapping("unlisted.test") == -1);
    const wchar_t *invalid[] = {L"", L"domain.com:1.2.3.4,domain.com:999.1.1.1",
        L"domain.com:1.2.3.4,domain.com:2001:db8::1", L"bad/host:1.2.3.4",
        L"domain.com:[not-an-ip]", L"domain.com:1.2.3.4 --bad-switch"};
    for (size_t i = 0; i < sizeof(invalid) / sizeof(invalid[0]); ++i) {
        clearMappings(); wcscpy(g_config.staticHostMappings, invalid[i]);
        assert(!ParseStaticHostProxyMappings()); assert(!g_hostProxyMappings);
    }
}
static void testRouting(void) {
    configure(L"domain.com:1.2.3.4,domain.com:3.4.5.6,domain.com:[2001:db8::1]");
    HostProxyMapping *m = &g_hostProxyMappings[0];
    resetRoute(); reachable[0] = reachable[1] = reachable[2] = TRUE;
    assert(connectRequest("A") == 100); assert(m->activeAddress == 0);
    resetRoute(); reachable[1] = reachable[2] = TRUE;
    assert(connectRequest("AB") == 101); assert(m->activeAddress == 1);
    assert(strcmp(m->currentAddress, "3.4.5.6") == 0 && titleChanges == 1);
    assert(connectRequest("B") == 101); // Reuse the working route during cooldown.
    reachable[1] = FALSE;
    assert(connectRequest("BC") == 102); assert(m->activeAddress == 2);
    reachable[2] = FALSE;
    assert(connectRequest("CD") == 200); assert(m->state == HOST_PROXY_FALLBACK);
    assert(strcmp(m->currentAddress, "9.9.9.9") == 0);
    assert(connectRequest("D") == 200);
    resetRoute(); assert(connectRequest("ABCD") == 200);
    resetRoute(); dnsWorks = FALSE; assert(connectRequest("ABCD") == INVALID_SOCKET);
    resetRoute(); g_hostProxyDnsFallback = FALSE; reachable[1] = TRUE;
    assert(connectRequest("AB") == 101);
    resetRoute(); g_hostProxyDnsFallback = FALSE;
    assert(connectRequest("ABC") == INVALID_SOCKET);
    reachable[0] = TRUE; assert(connectRequest("A") == 100); // Retry without DNS.
    resetRoute(); g_hostProxyStopping = TRUE;
    assert(connectRequest("") == INVALID_SOCKET); assert(m->state == HOST_PROXY_UNTESTED);
}
static void testProbes(void) {
    configure(L"domain.com:1.2.3.4,domain.com:3.4.5.6,domain.com:[2001:db8::1]");
    HostProxyMapping *m = &g_hostProxyMappings[0];
    resetRoute(); reachable[2] = TRUE; assert(connectRequest("ABC") == 102);
    ticks += HOST_PROXY_PROBE_INTERVAL_MS - 1;
    assert(connectRequest("C") == 102); assert(!pendingTask);
    ticks++; reachable[1] = TRUE;
    assert(connectRequest("C") == 102); assert(pendingTask);
    assert(connectRequest("C") == 102); // Single-flight while probe is pending.
    finishProbe("AB"); assert(m->activeAddress == 1);
    assert(connectRequest("B") == 101);
    ticks += HOST_PROXY_PROBE_INTERVAL_MS; reachable[0] = TRUE;
    assert(connectRequest("B") == 101); finishProbe("A");
    assert(m->activeAddress == 0 && strcmp(m->currentAddress, "1.2.3.4") == 0);
    // DNS recovery checks the entire mapped list in order.
    resetRoute(); assert(connectRequest("ABCD") == 200);
    ticks += HOST_PROXY_PROBE_INTERVAL_MS;
    assert(connectRequest("D") == 200); finishProbe("ABC");
    assert(!pendingTask && m->state == HOST_PROXY_FALLBACK);
    assert(connectRequest("D") == 200); assert(!pendingTask);
    ticks += HOST_PROXY_PROBE_INTERVAL_MS; reachable[1] = TRUE;
    assert(connectRequest("D") == 200); finishProbe("AB");
    assert(m->activeAddress == 1);
    // Recovery still works with DNS disabled; resume expires its cooldown.
    g_hostProxyDnsFallback = FALSE; reachable[0] = TRUE;
    HostProxyExpireFallbackCooldowns(); assert(connectRequest("B") == 101);
    finishProbe("A"); assert(m->activeAddress == 0);
    // Failed thread creation releases the single-flight claim.
    resetRoute(); assert(connectRequest("ABCD") == 200);
    ticks += HOST_PROXY_PROBE_INTERVAL_MS; threadFails = TRUE;
    assert(connectRequest("D") == 200); assert(!m->probeInFlight && !pendingTask);
    assert(!g_hostProxyWorkerCount); threadFails = FALSE;
}
static void testStaleAndTunnels(void) {
    configure(L"domain.com:1.2.3.4,domain.com:3.4.5.6,domain.com:[2001:db8::1]");
    HostProxyMapping *m = &g_hostProxyMappings[0];
    resetRoute(); reachable[1] = TRUE; assert(connectRequest("AB") == 101);
    ticks += HOST_PROXY_PROBE_INTERVAL_MS;
    assert(connectRequest("B") == 101); assert(pendingTask);
    reachable[1] = FALSE; reachable[2] = TRUE;
    assert(connectRequest("BC") == 102); // Newer route supersedes pending probe.
    ULONGLONG generation = m->routeGeneration; reachable[0] = TRUE;
    finishProbe("A"); assert(m->activeAddress == 2 && m->routeGeneration == generation);
    assert(!HostProxyUpdateRoute(m, generation - 1, 0, "1.2.3.4"));
    assert(strcmp(m->currentAddress, "2001:db8::1") == 0);
    HostProxyTunnel other = { .client = 5, .upstream = 6, .mappingIndex = 1 };
    HostProxyTunnel except = { .client = 3, .upstream = 4, .mappingIndex = 0, .next = &other };
    HostProxyTunnel old = { .client = 1, .upstream = 2, .mappingIndex = 0, .next = &except };
    g_hostProxyTunnelList = &old;
    CloseHostTunnelsForMapping(0, &except);
    assert(old.abortRequested && !except.abortRequested && !other.abortRequested && shut == 2);
    g_hostProxyTunnelList = NULL;
}
static void testBrowserRouting(void) {
    configure(L"domain.com:1.2.3.4");
    g_hostProxyDnsFallback = FALSE; assert(!HostProxyRequired());
    g_hostProxyDnsFallback = TRUE; assert(HostProxyRequired());
    configure(L"domain.com:1.2.3.4,domain.com:3.4.5.6");
    g_hostProxyDnsFallback = FALSE; assert(HostProxyRequired());
    char *pac = BuildHostProxyPacScript(); assert(pac);
    assert(strstr(pac, "PROXY 127.0.0.1:12345\""));
    assert(!strstr(pac, "; DIRECT")); assert(strstr(pac, "return \"DIRECT\"")); free(pac);
    g_hostProxyDnsFallback = TRUE; pac = BuildHostProxyPacScript();
    assert(strstr(pac, "PROXY 127.0.0.1:12345; DIRECT")); free(pac);
    g_config.useStaticHostMappings = TRUE; size_t count;
    for (int fallback = 0; fallback <= 1; ++fallback) {
        g_hostProxyDnsFallback = fallback;
        wchar_t *args = BuildStaticHostBrowserArguments(&count);
        assert(count == 1 && wcsstr(args, L"--proxy-pac-url=")); free(args);
    }
    g_hostProxyPort = 0; wchar_t *args = BuildStaticHostBrowserArguments(&count);
    assert(count == 2 && wcsstr(args, L"--host-resolver-rules=")); free(args);
    g_config.useStaticHostMappings = FALSE;
    assert(!BuildStaticHostBrowserArguments(&count) && count == 0);
}
int main(int argc, char **argv) {
    assert(argc == 2);
    if (!strcmp(argv[1], "parser")) testParser();
    else if (!strcmp(argv[1], "routing")) testRouting();
    else if (!strcmp(argv[1], "probes")) testProbes();
    else if (!strcmp(argv[1], "stale")) testStaleAndTunnels();
    else if (!strcmp(argv[1], "browser")) testBrowserRouting();
    else assert(0);
    clearMappings(); return 0;
}
'''
        (path / 'static_hosts.c').write_text(
            prelude + constants + '\n' + types + globals_ + '\n' +
            '\n\n'.join(function(name) for name in names) + main)
        cls.binary = path / 'static_hosts'
        subprocess.run(['cc', '-std=gnu11', '-Wall', '-Wextra', '-Werror',
                        '-fsanitize=address,undefined', '-g', str(path / 'static_hosts.c'),
                        '-o', str(cls.binary)], check=True)

    def test_parser_preserves_order_and_validates(self):
        subprocess.run([self.binary, 'parser'], check=True)

    def test_ordered_connections_and_optional_dns(self):
        subprocess.run([self.binary, 'routing'], check=True)

    def test_recovery_probes_cooldown_and_resume(self):
        subprocess.run([self.binary, 'probes'], check=True)

    def test_stale_results_and_tunnel_eviction(self):
        subprocess.run([self.binary, 'stale'], check=True)

    def test_proxy_selection_and_dns_gating(self):
        subprocess.run([self.binary, 'browser'], check=True)


if __name__ == '__main__':
    unittest.main()
