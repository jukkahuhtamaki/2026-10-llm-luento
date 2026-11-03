#!/usr/bin/env python3
# Agentti, jolla on yksi työkalu: säähavainnot Ilmatieteen laitoksen avoimesta datasta.
#
#   python agentti_saa.py          # lämpötila 0.7
#   python agentti_saa.py 0.0

import json
import sys
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

MALLI = "gpt-oss:20b"
LAMPOTILA = float(sys.argv[1]) if len(sys.argv) > 1 else 0.7
JARJESTELMAKEHOTE = "Olet tamperelainen avustaja. Vastaat lyhyesti ja aina suomeksi."

SUUREET = {"t2m": "lampotila_c", "ws_10min": "tuuli_ms", "r_1h": "sade_mm"}


def hae_saa(paikka, tunnit=24):
    alku = datetime.now(timezone.utc) - timedelta(hours=tunnit)
    osoite = "https://opendata.fmi.fi/wfs?" + urllib.parse.urlencode({
        "service": "WFS", "version": "2.0.0", "request": "getFeature",
        "storedquery_id": "fmi::observations::weather::simple",
        "place": paikka, "parameters": ",".join(SUUREET), "timestep": 60,
        "starttime": alku.strftime("%Y-%m-%dT%H:00:00Z"),
    })
    with urllib.request.urlopen(osoite) as r:
        puu = ET.parse(r)

    havainnot = {}
    ns = {"BsWfs": "http://xml.fmi.fi/schema/wfs/2.0"}
    for el in puu.iterfind(".//BsWfs:BsWfsElement", ns):
        aika = datetime.fromisoformat(el.findtext("BsWfs:Time", namespaces=ns))
        aika = aika.astimezone(ZoneInfo("Europe/Helsinki")).strftime("%Y-%m-%d %H:%M")
        suure = SUUREET[el.findtext("BsWfs:ParameterName", namespaces=ns)]
        arvo = float(el.findtext("BsWfs:ParameterValue", namespaces=ns))
        havainnot.setdefault(aika, {"aika": aika})[suure] = None if arvo != arvo else arvo
    return list(havainnot.values())


TYOKALUT = [{
    "type": "function",
    "function": {
        "name": "hae_saa",
        "description": "Hakee säähavainnot tunneittain Ilmatieteen laitokselta: "
                       "lämpötila, tuulen nopeus ja sademäärä.",
        "parameters": {
            "type": "object",
            "properties": {
                "paikka": {"type": "string", "description": "Paikkakunta tai kaupunginosa, esim. Härmälä"},
                "tunnit": {"type": "integer", "description": "Kuinka monta tuntia taaksepäin"},
            },
            "required": ["paikka"],
        },
    },
}]


def kysy_mallilta(viestit):
    data = json.dumps({
        "model": MALLI,
        "messages": viestit,
        "tools": TYOKALUT,
        "stream": False,
        "options": {"temperature": LAMPOTILA},
    }).encode()
    with urllib.request.urlopen("http://localhost:11434/api/chat", data) as r:
        return json.load(r)["message"]


viestit = [{"role": "system", "content": JARJESTELMAKEHOTE}]

try:
    while True:
        viestit.append({"role": "user", "content": input("\nSinä: ")})
        vastaus = kysy_mallilta(viestit)

        # Malli voi pyytää työkalua, ohjelma ajaa sen ja palauttaa tuloksen
        while vastaus.get("tool_calls"):
            viestit.append(vastaus)
            for kutsu in vastaus["tool_calls"]:
                argumentit = kutsu["function"]["arguments"]
                print(f"  [työkalu: hae_saa({argumentit})]")
                tulos = hae_saa(**argumentit)
                viestit.append({"role": "tool", "tool_name": "hae_saa",
                                "content": json.dumps(tulos, ensure_ascii=False)})
            vastaus = kysy_mallilta(viestit)

        viestit.append(vastaus)
        print("\nAgentti:", vastaus["content"])
except (EOFError, KeyboardInterrupt):
    print()
