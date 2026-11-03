# Kielimallista agentiksi

Luennon demot suurista kielimalleista: sokkotesti kolmella kielimallilla ja pieni agentti Pythonilla.

## Demot

| Tiedosto | Mitä tekee |
|---|---|
| `demot/turing.py` | Sokkotesti: paikalliset kielimallit vastaavat samoihin kysymyksiin, yleisö arvaa kuka on kuka. `--raw` näyttää pelkän tekstin jatkamisen ilman keskustelupohjaa. |
| `demot/agentti.py` | Pieni agentti: järjestelmäkehote, lämpötila ja muisti noin 30 rivissä. |
| `demot/agentti_saa.py` | Sama agentti ja yksi työkalu, joka hakee säähavainnot Ilmatieteen laitoksen avoimesta datasta. |

Demot käyttävät vain Pythonin standardikirjastoa.

### Tarvitaan

- Python 3.9 tai uudempi
- [Ollama](https://ollama.com) ja mallit:

  ```sh
  ollama pull tinyllama
  ollama pull gpt-oss:20b
  ```

### Käyttö

```sh
python3 demot/turing.py                 # oletuskysymykset
python3 demot/turing.py --raw           # pelkkä kielimalli
python3 demot/turing.py "Oma kysymys"

python3 demot/agentti.py 0.7            # lämpötila komentoriviltä
python3 demot/agentti_saa.py 0.2        # kokeile: "Sää Härmälässä viimeisen 24 tunnin aikana JSON-muodossa"
```

## Diat

`diat/demot.pptx` on demo-osion diasarja. Sen voi rakentaa uudelleen skriptillä:

```sh
pip install python-pptx lxml "qrcode[pil]"
cd diat && python3 rakenna_diat.py
```

## Sanasto

Luennon termit ovat tiedostossa [`sanasto.md`](sanasto.md), esimerkiksi token = sanake.

## Lisenssi

MIT
