# SPARTY Safe NTLM Edition

SPARTY Safe NTLM Edition is a Python-based auditing tool for SharePoint and FrontPage-style environments. It performs non-destructive enumeration of common SharePoint and FrontPage endpoints, checks for exposed services and metadata paths, detects likely login pages, and can optionally use NTLM authentication.

It is designed for reconnaissance and assessment rather than exploitation. The current version does not upload files, dump credentials, or remove content.

```
 __    ___  _      __  _____        _____  ___  
/ _\  / _ \/_\    /__\/__   \/\_/\ |___ / / _ \ 
\ \  / /_)//_\\  / \//  / /\/\_ _/   |_ \| | | |
_\ \/ ___/  _  \/ _  \ / /    / \   ___) | |_| |
\__/\/   \_/ \_/\/ \_/ \/     \_/  |____(_)___/ 
- Ported from SPARTY 2.0 by Kamran Saif Ullah Khan       
```

## Features

- Collects basic target information such as HTTP status, server headers, and SharePoint headers
- Fingerprints FrontPage-related endpoints such as `_vti_inf.html` and common `_vti_bin` paths
- Checks for common FrontPage and SharePoint directories, services, metadata files, and layout paths
- Detects whether a response appears to be a login page or authentication challenge
- Supports custom headers, cookies, proxies, and optional NTLM authentication
- Saves results to a CSV report for later review

## Installation

1. Clone the repository.
2. Install dependencies.
3. For NTLM support, the required package is already listed in `requirements.txt`.

## Recommended Python Version

- Python 3.8 or newer is recommended.

## Usage

Run the tool with the target URL and the `-enum` flag:

```bash
python Sparty-3.0.py -u https://sharepoint.example.com/sites/test -enum
```

### Main Options

- `-u`, `--url` - Target URL to audit
- `-enum`, `--enumeration` - Run the non-destructive enumeration checks
- `-p`, `--proxy` - Proxy URL, for example `http://127.0.0.1:8080`
- `-hds`, `--headers` - Custom headers as `key=value` pairs
- `--cookie` - Raw Cookie header value
- `--ntlm` - Enable NTLM authentication
- `--domain` - Windows domain for NTLM auth
- `--username` - Username for NTLM auth
- `--password` - Password for NTLM auth
- `-t`, `--timeout` - HTTP timeout in seconds (default: 10)
- `-o`, `--output` - CSV output file name
- `-h` - Show help

### Example with NTLM

```bash
python Sparty-3.0.py -u https://sharepoint.example.com/sites/test -enum --ntlm --domain CONTOSO --username user --password 'Secret123'
```

## What the Enumeration Checks

The script audits a broad set of SharePoint and FrontPage-related paths, including:

- FrontPage binary and RPC paths
- FrontPage private metadata paths
- Common SharePoint / FrontPage web service endpoints
- Classic SharePoint layout paths
- SharePoint 15 layout paths
- SharePoint forms and catalog paths

## Notes and Safety Guidance

- Always provide the full target URL, including `http://` or `https://`.
- Use the correct site path if the target is hosted under a subpath.
- Double-check any results manually before acting on them.
- This tool is intended for authorized assessment and should be used responsibly.

## References

- Original SPARTY project: https://github.com/adityaks/sparty
- Sparty 2.0 Project: https://github.com/MayankPandey01/Sparty-2.0

## Happy hacking!
