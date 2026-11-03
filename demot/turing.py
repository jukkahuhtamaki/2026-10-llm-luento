#!/usr/bin/env python3
# Sokkotesti: kielimallit vastaavat samaan kysymykseen, yleisö arvaa kuka on kuka.
#
#   python turing.py               # oletuskysymykset
#   python turing.py "Oma kysymys"
#   python turing.py --raw         # pelkkä tekstin jatkaminen, ei chattia

import json
import random
import sys
import time
import urllib.request

KYSYMYKSET = [
    "Täydennä lause: Moro ...",
    "Sää Härmälässä viimeisen 24 tunnin aikana JSON-muodossa",
    "Määrittele GenAI:n ja kielimallin suhde.",
]

MALLIT = {
    "TinyLlama (1,1 mrd parametria)": "tinyllama",
    "gpt-oss (20 mrd parametria)": "gpt-oss:20b",
}


def ollama(polku, data):
    req = urllib.request.Request("http://localhost:11434" + polku,
                                 json.dumps(data).encode(),
                                 {"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=600) as r:
        return json.load(r)


def kysy(malli, kysymys):
    return ollama("/api/chat", {
        "model": malli,
        "messages": [{"role": "user", "content": kysymys}],
        "stream": False,
    })["message"]["content"]


# Ilman keskustelupohjaa: malli vain jatkaa tekstiä sanake kerrallaan.
def jatka(malli, teksti):
    return teksti + ollama("/api/generate", {
        "model": malli,
        "prompt": teksti,
        "raw": True,
        "stream": False,
        "options": {"num_predict": 40},
    })["response"]


def main():
    argit = sys.argv[1:]
    vastaa = kysy
    if "--raw" in argit:
        argit.remove("--raw")
        vastaa = jatka
    kysymykset = argit or KYSYMYKSET

    for nro, kysymys in enumerate(kysymykset, 1):
        print(f"\n{'=' * 70}\nKysymys {nro}: {kysymys}\n{'=' * 70}")

        nimet = list(MALLIT)
        random.shuffle(nimet)
        ajat = {}

        for kirjain, nimi in zip("ABC", nimet):
            alku = time.time()
            try:
                vastaus = vastaa(MALLIT[nimi], kysymys)
            except Exception as e:
                vastaus = f"(virhe: {e})"
            ajat[kirjain] = time.time() - alku
            print(f"\n--- Vastaaja {kirjain} ---\n{vastaus.strip()}")

        input("\n[Enter paljastaa vastaajat]")
        for kirjain, nimi in zip("ABC", nimet):
            print(f"  {kirjain} = {nimi}  ({ajat[kirjain]:.1f} s)")

        if nro < len(kysymykset):
            input("\n[Enter seuraavaan kysymykseen]")


if __name__ == "__main__":
    main()
