#!/usr/bin/env python3
# Pieni agentti: järjestelmäkehote, lämpötila ja keskustelumuisti.
#
#   python agentti.py          # lämpötila 0.7
#   python agentti.py 1.8      # villimpi

import json
import sys
import urllib.request

MALLI = "gpt-oss:20b"
LAMPOTILA = float(sys.argv[1]) if len(sys.argv) > 1 else 0.7
JARJESTELMAKEHOTE = "Olet tamperelainen avustaja. Vastaat lyhyesti ja aina suomeksi."

viestit = [{"role": "system", "content": JARJESTELMAKEHOTE}]

try:
    while True:
        viestit.append({"role": "user", "content": input("\nSinä: ")})
        data = json.dumps({
            "model": MALLI,
            "messages": viestit,
            "stream": False,
            "options": {"temperature": LAMPOTILA},
        }).encode()
        with urllib.request.urlopen("http://localhost:11434/api/chat", data) as r:
            vastaus = json.load(r)["message"]
        viestit.append(vastaus)
        print("\nAgentti:", vastaus["content"])
except (EOFError, KeyboardInterrupt):
    print()
