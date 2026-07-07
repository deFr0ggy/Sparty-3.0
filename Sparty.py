#!/usr/bin/env python3

import argparse
import csv
import getpass
import re
import sys
from urllib.parse import urljoin

import requests
import colorama
from urllib3.exceptions import InsecureRequestWarning

requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)

colorama.init(autoreset=True)

GREEN = colorama.Fore.GREEN
GRAY = colorama.Fore.LIGHTBLACK_EX
RESET = colorama.Fore.RESET
RED = colorama.Fore.RED
BLUE = colorama.Fore.BLUE
CYAN = colorama.Fore.CYAN
YELLOW = colorama.Fore.YELLOW


FRONT_BIN = [
    "_vti_inf.html",
    "_vti_bin/shtml.dll/_vti_rpc",
    "_vti_bin/owssvr.dll",
    "_vti_bin/_vti_adm/admin.dll",
    "_vti_bin/_vti_adm/admin.exe",
    "_vti_bin/_vti_aut/author.exe",
    "_vti_bin/_vti_aut/WS_FTP.log",
    "_vti_bin/_vti_aut/ws_ftp.log",
    "_vti_bin/shtml.exe/_vti_rpc",
    "_vti_bin/_vti_aut/author.dll",
]

FRONT_SERVICES = [
    "_vti_bin/Admin.asmx",
    "_vti_bin/alerts.asmx",
    "_vti_bin/dspsts.asmx",
    "_vti_bin/forms.asmx",
    "_vti_bin/Lists.asmx",
    "_vti_bin/people.asmx",
    "_vti_bin/Permissions.asmx",
    "_vti_bin/search.asmx",
    "_vti_bin/UserGroup.asmx",
    "_vti_bin/versions.asmx",
    "_vti_bin/Views.asmx",
    "_vti_bin/webpartpages.asmx",
    "_vti_bin/webs.asmx",
    "_vti_bin/spsdisco.aspx",
    "_vti_bin/AreaService.asmx",
    "_vti_bin/BusinessDataCatalog.asmx",
    "_vti_bin/ExcelService.asmx",
    "_vti_bin/SharepointEmailWS.asmx",
    "_vti_bin/spscrawl.asmx",
    "_vti_bin/spsearch.asmx",
    "_vti_bin/UserProfileService.asmx",
    "_vti_bin/WebPartPages.asmx",
]

FRONT_PVT = [
    "_vti_pvt/authors.pwd",
    "_vti_pvt/administrators.pwd",
    "_vti_pvt/users.pwd",
    "_vti_pvt/service.pwd",
    "_vti_pvt/service.grp",
    "_vti_pvt/bots.cnf",
    "_vti_pvt/service.cnf",
    "_vti_pvt/access.cnf",
    "_vti_pvt/writeto.cnf",
    "_vti_pvt/botsinf.cnf",
    "_vti_pvt/doctodep.btr",
    "_vti_pvt/deptodoc.btr",
    "_vti_pvt/linkinfo.cnf",
    "_vti_pvt/services.org",
    "_vti_pvt/structure.cnf",
    "_vti_pvt/svcacl.cnf",
    "_vti_pvt/uniqperm.cnf",
    "_vti_pvt/service/lck",
    "_vti_pvt/frontpg.lck",
]

DIRECTORY_CHECK = [
    "_vti_pvt/",
    "_vti_bin/",
    "_vti_log/",
    "_vti_cnf/",
    "_vti_bot",
    "_vti_bin/_vti_adm",
    "_vti_bin/_vti_aut",
    "_vti_txt/",
]

SHAREPOINT_LAYOUT = [
    "_layouts/aclinv.aspx",
    "_layouts/addrole.aspx",
    "_layouts/AdminRecycleBin.aspx",
    "_layouts/AreaNavigationSettings.aspx",
    "_Layouts/AreaTemplateSettings.aspx",
    "_Layouts/AreaWelcomePage.aspx",
    "_layouts/associatedgroups.aspx",
    "_layouts/bpcf.aspx",
    "_Layouts/ChangeSiteMasterPage.aspx",
    "_layouts/create.aspx",
    "_layouts/editgrp.aspx",
    "_layouts/editprms.aspx",
    "_layouts/groups.aspx",
    "_layouts/help.aspx",
    "_layouts/images/",
    "_layouts/listedit.aspx",
    "_layouts/ManageFeatures.aspx",
    "_layouts/mcontent.aspx",
    "_layouts/mngctype.aspx",
    "_layouts/mngfield.aspx",
    "_layouts/mngsiteadmin.aspx",
    "_layouts/mngsubwebs.aspx",
    "_layouts/mngsubwebs.aspx?view=sites",
    "_layouts/mobile/mbllists.aspx",
    "_layouts/MyInfo.aspx",
    "_layouts/MyPage.aspx",
    "_layouts/MyTasks.aspx",
    "_layouts/navoptions.aspx",
    "_layouts/NewDwp.aspx",
    "_layouts/newgrp.aspx",
    "_layouts/newsbweb.aspx",
    "_layouts/PageSettings.aspx",
    "_layouts/people.aspx",
    "_layouts/people.aspx?MembershipGroupId=0",
    "_layouts/permsetup.aspx",
    "_layouts/picker.aspx",
    "_layouts/policy.aspx",
    "_layouts/policyconfig.aspx",
    "_layouts/policycts.aspx",
    "_layouts/Policylist.aspx",
    "_layouts/prjsetng.aspx",
    "_layouts/quiklnch.aspx",
    "_layouts/recyclebin.aspx",
    "_Layouts/RedirectPage.aspx",
    "_layouts/role.aspx",
    "_layouts/settings.aspx",
    "_layouts/SiteDirectorySettings.aspx",
    "_layouts/sitemanager.aspx",
    "_layouts/SiteManager.aspx?lro=all",
    "_layouts/spcf.aspx",
    "_layouts/storman.aspx",
    "_layouts/themeweb.aspx",
    "_layouts/topnav.aspx",
    "_layouts/user.aspx",
    "_layouts/userdisp.aspx",
    "_layouts/userdisp.aspx?ID=1",
    "_layouts/useredit.aspx",
    "_layouts/useredit.aspx?ID=1",
    "_layouts/viewgrouppermissions.aspx",
    "_layouts/viewlsts.aspx",
    "_layouts/vsubwebs.aspx",
    "_layouts/WPPrevw.aspx?ID=247",
    "_layouts/wrkmng.aspx",
]

SHAREPOINT_LAYOUT_15 = [
    "_layouts/15/aclinv.aspx",
    "_layouts/15/addrole.aspx",
    "_layouts/15/AdminRecycleBin.aspx",
    "_layouts/15/AreaNavigationSettings.aspx",
    "_layouts/15/AreaTemplateSettings.aspx",
    "_layouts/15/AreaWelcomePage.aspx",
    "_layouts/15/associatedgroups.aspx",
    "_layouts/15/bpcf.aspx",
    "_layouts/15/ChangeSiteMasterPage.aspx",
    "_layouts/15/create.aspx",
    "_layouts/15/editgrp.aspx",
    "_layouts/15/editprms.aspx",
    "_layouts/15/groups.aspx",
    "_layouts/15/help.aspx",
    "_layouts/15/images/",
    "_layouts/15/listedit.aspx",
    "_layouts/15/ManageFeatures.aspx",
    "_layouts/15/mcontent.aspx",
    "_layouts/15/mngctype.aspx",
    "_layouts/15/mngfield.aspx",
    "_layouts/15/mngsiteadmin.aspx",
    "_layouts/15/mngsubwebs.aspx",
    "_layouts/15/mngsubwebs.aspx?view=sites",
    "_layouts/15/mobile/mbllists.aspx",
    "_layouts/15/MyInfo.aspx",
    "_layouts/15/MyPage.aspx",
    "_layouts/15/MyTasks.aspx",
    "_layouts/15/navoptions.aspx",
    "_layouts/15/NewDwp.aspx",
    "_layouts/15/newgrp.aspx",
    "_layouts/15/newsbweb.aspx",
    "_layouts/15/PageSettings.aspx",
    "_layouts/15/people.aspx",
    "_layouts/15/people.aspx?MembershipGroupId=0",
    "_layouts/15/permsetup.aspx",
    "_layouts/15/picker.aspx",
    "_layouts/15/policy.aspx",
    "_layouts/15/policyconfig.aspx",
    "_layouts/15/policycts.aspx",
    "_layouts/15/Policylist.aspx",
    "_layouts/15/prjsetng.aspx",
    "_layouts/15/quiklnch.aspx",
    "_layouts/15/recyclebin.aspx",
    "_layouts/15/RedirectPage.aspx",
    "_layouts/15/role.aspx",
    "_layouts/15/settings.aspx",
    "_layouts/15/SiteDirectorySettings.aspx",
    "_layouts/15/sitemanager.aspx",
    "_layouts/15/SiteManager.aspx?lro=all",
    "_layouts/15/spcf.aspx",
    "_layouts/15/storman.aspx",
    "_layouts/15/themeweb.aspx",
    "_layouts/15/topnav.aspx",
    "_layouts/15/user.aspx",
    "_layouts/15/userdisp.aspx",
    "_layouts/15/userdisp.aspx?ID=1",
    "_layouts/15/useredit.aspx",
    "_layouts/15/useredit.aspx?ID=1",
    "_layouts/15/viewgrouppermissions.aspx",
    "_layouts/15/viewlsts.aspx",
    "_layouts/15/vsubwebs.aspx",
    "_layouts/15/WPPrevw.aspx?ID=247",
    "_layouts/15/wrkmng.aspx",
    "_layouts/15/sitepermissions.aspx",
    "_layouts/15/ChangeGroupSettings.aspx",
    "_layouts/15/AppPrincipals.aspx",
    "_layouts/15/appinv.aspx",
    "_layouts/15/appregnew.aspx",
]

SHAREPOINT_FORMS = [
    "Forms/DispForm.aspx",
    "Forms/DispForm.aspx?ID=1",
    "Forms/EditForm.aspx",
    "Forms/EditForm.aspx?ID=1",
    "Forms/Forms/AllItems.aspx",
    "Forms/MyItems.aspx",
    "Forms/NewForm.aspx",
    "Pages/default.aspx",
    "Pages/Forms/AllItems.aspx",
]

SHAREPOINT_CATALOG = [
    "_catalogs/masterpage/Forms/AllItems.aspx",
    "_catalogs/wp/Forms/AllItems.aspx",
    "_catalogs/wt/Forms/Common.aspx",
]


def banner():
    print(f"""{RED}
 __    ___  _      __  _____        _____  ___  
/ _\  / _ \/_\    /__\/__   \/\_/\ |___ / / _ \ 
\ \  / /_)//_\\  / \//  / /\/\_ _/   |_ \| | | |
_\ \/ ___/  _  \/ _  \ / /    / \   ___) | |_| |
\__/\/   \_/ \_/\/ \_/ \/     \_/  |____(_)___/ 
- Ported from SPARTY 2.0 by Kamran Saif Ullah Khan       

{RESET}""")
    print(f"{CYAN}SPARTY Safe NTLM Edition: SharePoint / FrontPage Non-Destructive Auditing Tool{RESET}")
    print(f"{YELLOW}Supports unauthenticated, cookie-based, and NTLM-authenticated checks.{RESET}")
    print(f"{YELLOW}No credential dumping, upload, deletion, or destructive action is performed.{RESET}\n")


def normalize_base_url(url):
    return url.rstrip("/")


def build_url(base_url, path):
    return urljoin(base_url + "/", path)


def parse_headers(header_items):
    headers = {}

    if not header_items:
        return headers

    for item in header_items:
        if "=" not in item:
            print(f"{RED}Invalid header format: {item}. Expected key=value{RESET}")
            sys.exit(1)

        key, value = item.split("=", 1)
        headers[key.strip()] = value.strip()

    return headers


def make_proxy(proxy_url):
    if not proxy_url:
        return {}

    return {
        "http": proxy_url,
        "https": proxy_url,
    }


def setup_auth(args):
    if not args.ntlm:
        return None

    try:
        from requests_ntlm import HttpNtlmAuth
    except ImportError:
        print(f"{RED}[-] Missing dependency: requests-ntlm{RESET}")
        print("Install it using:")
        print("    pip3 install requests-ntlm")
        sys.exit(1)

    if not args.username:
        print(f"{RED}[-] NTLM authentication requires --username{RESET}")
        sys.exit(1)

    username = args.username

    if args.domain and "\\" not in username and "@" not in username:
        username = f"{args.domain}\\{username}"

    password = args.password

    if not password:
        password = getpass.getpass("NTLM Password: ")

    return HttpNtlmAuth(username, password)


def detect_login_page(response):
    body_sample = response.text[:5000].lower()
    final_url = response.url.lower()

    login_indicators = [
        "sign in",
        "signin",
        "log in",
        "login",
        "adfs",
        "saml",
        "fedauth",
        "rtfa",
        "microsoftonline",
        "office 365",
        "forms-based authentication",
        "wa=wsignin",
        "wreply",
        "oauth2",
        "authorize",
        "idpinitiatedsignon",
        "onmicrosoft",
        "enter your password",
        "enter password",
        "pick an account",
    ]

    if any(indicator in body_sample for indicator in login_indicators):
        return True

    if any(indicator in final_url for indicator in login_indicators):
        return True

    return False


def create_session(headers, proxies, auth):
    session = requests.Session()
    session.headers.update(headers)
    session.proxies.update(proxies)

    if auth:
        session.auth = auth

    return session


def request_get(session, url, timeout):
    try:
        response = session.get(
            url,
            verify=False,
            timeout=timeout,
            allow_redirects=True,
        )

        appears_login_page = detect_login_page(response)

        return {
            "url": url,
            "final_url": response.url,
            "status_code": response.status_code,
            "content_length": len(response.content),
            "content_type": response.headers.get("Content-Type", ""),
            "server": response.headers.get("Server", ""),
            "sharepoint_version": response.headers.get("MicrosoftSharePointTeamServices", ""),
            "sp_request_guid": response.headers.get("SPRequestGuid", ""),
            "x_sharepoint_health_score": response.headers.get("X-SharePointHealthScore", ""),
            "www_authenticate": response.headers.get("WWW-Authenticate", ""),
            "appears_login_page": appears_login_page,
            "error": "",
        }

    except requests.exceptions.RequestException as exc:
        return {
            "url": url,
            "final_url": "",
            "status_code": "",
            "content_length": "",
            "content_type": "",
            "server": "",
            "sharepoint_version": "",
            "sp_request_guid": "",
            "x_sharepoint_health_score": "",
            "www_authenticate": "",
            "appears_login_page": False,
            "error": str(exc),
        }


def target_information(base_url, session, timeout):
    print(f"\n{GREEN}[*] Target Information{RESET}\n")

    result = request_get(session, base_url, timeout)

    if result["error"]:
        print(f"{RED}[-] Error connecting to target: {result['error']}{RESET}")
        return result

    print(f"[+] Target URL: {result['final_url']}")
    print(f"[+] HTTP Status: {result['status_code']}")

    if result["server"]:
        print(f"[+] Server Header: {result['server']}")

    if result["sharepoint_version"]:
        print(f"[+] SharePoint Version Header: {result['sharepoint_version']}")

    if result["x_sharepoint_health_score"]:
        print(f"[+] SharePoint Health Score: {result['x_sharepoint_health_score']}")

    if result["sp_request_guid"]:
        print(f"[+] SPRequestGuid Present: {result['sp_request_guid']}")

    if result["www_authenticate"]:
        print(f"[+] WWW-Authenticate: {result['www_authenticate']}")

    if result["appears_login_page"]:
        print(f"{BLUE}[*] Response appears to be a login page.{RESET}")

    print("")

    return result


def fingerprint_frontpage(base_url, session, timeout):
    print(f"\n{GREEN}[*] FrontPage Fingerprinting{RESET}\n")

    paths_nix = [
        "_vti_bin/_vti_aut/author.exe",
        "_vti_bin/_vti_adm/admin.exe",
        "_vti_bin/shtml.exe",
    ]

    paths_win = [
        "_vti_bin/_vti_aut/author.dll",
        "_vti_bin/_vti_aut/dvwssr.dll",
        "_vti_bin/_vti_adm/admin.dll",
        "_vti_bin/shtml.dll",
    ]

    findings = []

    for path in paths_nix:
        url = build_url(base_url, path)
        result = request_get(session, url, timeout)
        result["category"] = "FrontPage Fingerprinting"
        result["path"] = path
        findings.append(result)

        if result["status_code"] == 200 and not result["appears_login_page"]:
            print(f"{YELLOW}[!] Possible FrontPage UNIX endpoint exposed: {url}{RESET}")

    for path in paths_win:
        url = build_url(base_url, path)
        result = request_get(session, url, timeout)
        result["category"] = "FrontPage Fingerprinting"
        result["path"] = path
        findings.append(result)

        if result["status_code"] == 200 and not result["appears_login_page"]:
            print(f"{YELLOW}[!] Possible FrontPage Windows endpoint exposed: {url}{RESET}")

    version_url = build_url(base_url, "_vti_inf.html")
    result = request_get(session, version_url, timeout)
    result["category"] = "FrontPage Fingerprinting"
    result["path"] = "_vti_inf.html"
    findings.append(result)

    if result["status_code"] == 200 and not result["appears_login_page"]:
        try:
            response = session.get(
                version_url,
                verify=False,
                timeout=timeout,
                allow_redirects=True,
            )
            version_matches = re.findall(r"FPVersion=(.*)", response.text)

            if version_matches:
                print(f"{YELLOW}[!] FrontPage version indicator found: {version_matches}{RESET}")
            else:
                print("[*] _vti_inf.html found, but FPVersion was not identified.")
        except requests.exceptions.RequestException:
            pass

    print("")

    return findings


def print_result(url, result):
    status = result["status_code"]

    if result["error"]:
        print(f"{RED}[-] {url} -> ERROR: {result['error']}{RESET}")
        return

    if status in [200, 201, 202, 203]:
        if result.get("appears_login_page"):
            print(f"{BLUE}[*] {url} -> {status} Login page returned, not confirmed accessible{RESET}")
        else:
            print(f"{YELLOW}[!] {url} -> {status} Accessible with current session/request{RESET}")

    elif status in [301, 302, 303, 307, 308]:
        print(f"{BLUE}[*] {url} -> {status} Redirect to: {result.get('final_url', '')}{RESET}")

    elif status == 401:
        print(f"{GREEN}[+] {url} -> 401 Authentication required or NTLM negotiation failed{RESET}")
        if result.get("www_authenticate"):
            print(f"    Auth method: {result.get('www_authenticate')}")

    elif status == 403:
        print(f"{GREEN}[+] {url} -> 403 Forbidden / access denied{RESET}")

    elif status == 404:
        print(f"{GRAY}[-] {url} -> 404 Not found{RESET}")

    elif status == 500:
        print(f"{YELLOW}[!] {url} -> 500 Server error; path may exist but caused an error{RESET}")

    else:
        print(f"[*] {url} -> {status}")


def audit_paths(base_url, paths, category, session, timeout):
    print(f"\n{GREEN}[*] Auditing {category}{RESET}\n")

    results = []

    for path in paths:
        url = build_url(base_url, path)
        result = request_get(session, url, timeout)
        result["category"] = category
        result["path"] = path
        results.append(result)

        print_result(url, result)

    print("")

    return results


def save_csv(results, output_file):
    fieldnames = [
        "category",
        "path",
        "url",
        "final_url",
        "status_code",
        "content_length",
        "content_type",
        "server",
        "sharepoint_version",
        "sp_request_guid",
        "x_sharepoint_health_score",
        "www_authenticate",
        "appears_login_page",
        "error",
    ]

    with open(output_file, "w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()

        for result in results:
            row = {key: result.get(key, "") for key in fieldnames}
            writer.writerow(row)

    print(f"{GREEN}[+] Results saved to: {output_file}{RESET}")


def summarize(results):
    print(f"\n{GREEN}[*] Summary{RESET}\n")

    total = len(results)

    accessible = [
        r for r in results
        if str(r.get("status_code", "")) in ["200", "201", "202", "203"]
        and not r.get("appears_login_page")
    ]

    login_page = [
        r for r in results
        if r.get("appears_login_page")
    ]

    auth_required = [
        r for r in results
        if str(r.get("status_code", "")) == "401"
    ]

    forbidden = [
        r for r in results
        if str(r.get("status_code", "")) == "403"
    ]

    not_found = [
        r for r in results
        if str(r.get("status_code", "")) == "404"
    ]

    errors = [
        r for r in results
        if r.get("error")
    ]

    print(f"Total checks: {total}")
    print(f"Accessible with current session/request: {len(accessible)}")
    print(f"Login page returned: {len(login_page)}")
    print(f"Authentication required / failed 401: {len(auth_required)}")
    print(f"Forbidden 403: {len(forbidden)}")
    print(f"Not found 404: {len(not_found)}")
    print(f"Errors: {len(errors)}")

    if accessible:
        print(f"\n{YELLOW}[!] Accessible paths:{RESET}")
        for item in accessible:
            print(f"    - {item.get('url')} [{item.get('status_code')}]")

    print("")


def main():
    banner()

    parser = argparse.ArgumentParser(
        description="SPARTY Safe NTLM Edition: SharePoint / FrontPage Non-Destructive Auditing Tool"
    )

    parser.add_argument(
        "-u",
        "--url",
        required=True,
        help="Target URL. Example: https://sharepoint.example.com/sites/test"
    )

    parser.add_argument(
        "-enum",
        "--enumeration",
        action="store_true",
        help="Run non-destructive enumeration checks"
    )

    parser.add_argument(
        "-p",
        "--proxy",
        help="Proxy URL. Example: http://127.0.0.1:8080"
    )

    parser.add_argument(
        "-hds",
        "--headers",
        nargs="+",
        help="Custom headers as key=value pairs. Example: 'User-Agent=Mozilla/5.0'"
    )

    parser.add_argument(
        "--cookie",
        help="Raw Cookie header copied from browser. Example: 'FedAuth=xxx; rtFa=yyy'"
    )

    parser.add_argument(
        "--ntlm",
        action="store_true",
        help="Enable NTLM authentication"
    )

    parser.add_argument(
        "--domain",
        help="Windows domain. Example: CONTOSO"
    )

    parser.add_argument(
        "--username",
        help="Username. Examples: CONTOSO\\user, user@domain.local, or user with --domain"
    )

    parser.add_argument(
        "--password",
        help="Password. If not provided with --ntlm, the script will prompt securely."
    )

    parser.add_argument(
        "-t",
        "--timeout",
        type=int,
        default=10,
        help="HTTP timeout in seconds. Default: 10"
    )

    parser.add_argument(
        "-o",
        "--output",
        default="sparty_safe_ntlm_results.csv",
        help="CSV output file. Default: sparty_safe_ntlm_results.csv"
    )

    args = parser.parse_args()

    base_url = normalize_base_url(args.url)
    headers = parse_headers(args.headers)
    proxies = make_proxy(args.proxy)
    timeout = args.timeout

    if args.cookie:
        headers["Cookie"] = args.cookie

    if not any(key.lower() == "user-agent" for key in headers.keys()):
        headers["User-Agent"] = (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/126.0.0.0 Safari/537.36"
        )

    auth = setup_auth(args)
    session = create_session(headers, proxies, auth)

    all_results = []

    target_result = target_information(base_url, session, timeout)
    target_result["category"] = "Target Information"
    target_result["path"] = "/"
    all_results.append(target_result)

    if args.enumeration:
        print(f"{BLUE}[!!] Running Non-Destructive Enumeration Module{RESET}")

        all_results.extend(
            fingerprint_frontpage(
                base_url,
                session,
                timeout
            )
        )

        all_results.extend(
            audit_paths(
                base_url,
                DIRECTORY_CHECK,
                "FrontPage / SharePoint Directories",
                session,
                timeout
            )
        )

        all_results.extend(
            audit_paths(
                base_url,
                FRONT_BIN,
                "FrontPage Binary / RPC Paths",
                session,
                timeout
            )
        )

        all_results.extend(
            audit_paths(
                base_url,
                FRONT_PVT,
                "FrontPage Private Metadata Paths",
                session,
                timeout
            )
        )

        all_results.extend(
            audit_paths(
                base_url,
                FRONT_SERVICES,
                "SharePoint / FrontPage Web Services",
                session,
                timeout
            )
        )

        all_results.extend(
            audit_paths(
                base_url,
                SHAREPOINT_LAYOUT,
                "Classic SharePoint Layout Paths",
                session,
                timeout
            )
        )

        all_results.extend(
            audit_paths(
                base_url,
                SHAREPOINT_LAYOUT_15,
                "SharePoint Layout 15 Paths",
                session,
                timeout
            )
        )

        all_results.extend(
            audit_paths(
                base_url,
                SHAREPOINT_FORMS,
                "SharePoint Forms Paths",
                session,
                timeout
            )
        )

        all_results.extend(
            audit_paths(
                base_url,
                SHAREPOINT_CATALOG,
                "SharePoint Catalog Paths",
                session,
                timeout
            )
        )

    else:
        print(f"{RED}[!!] No module selected. Use -enum to run enumeration.{RESET}")

    save_csv(all_results, args.output)
    summarize(all_results)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n{RED}Keyboard interrupt detected. Exiting.{RESET}")
        sys.exit(0)
