# Cheatsheet Listini

## Nuovo Adapter

```text
Usa $listini-normalization.
Aggiungi un nuovo adapter per questo file: [path file]
Prima di scrivere codice:
1. individua il dispatcher e il detector giusto
2. trova l'adapter più simile
3. separa parsing, business rules, TMS mapping e validation
4. indica i rischi di regressione
Poi implementa solo cambi locali e valida il caso target.
```

## Bug Fix

```text
Usa $listini-normalization.
Indaga questo bug: [descrizione]
Prima:
1. trova il code path esatto
2. spiega il comportamento attuale
3. mostra perché fallisce
4. proponi la fix più piccola sicura
Poi applica la fix e verifica i casi correlati.
```

## Review Senza Modifiche

```text
Usa $listini-normalization.
Non modificare codice.
Rivedi:
- validity
- direction inference
- surcharge filtering
- arbitrary handling
- TEU/container evaluation
- transshipment assignment
- detection order
- TMS output mapping
Restituisci: comportamento attuale, rischi, ambiguità, test consigliati.
```

## Aree Ad Alto Rischio

- validity dates
- direction inference
- surcharge scope
- arbitrary origin/destination
- transshipment assignment
- UNLOCODE resolution
- TEU vs container evaluation
- output column mapping

## Frase Chiave

Questo progetto non fa semplice estrazione dati.
Applica regole business rigide per produrre un output TMS controllato.
In caso di dubbio: privilegia correttezza, tracciabilità e cambi minimi.

## Mini Checklist Finale

- carrier corretto?
- direction corretta?
- validity corretta?
- surcharge ammesse soltanto?
- inclusive escluse?
- arbitrary separato o unito volutamente?
- transshipment giustificato?
- regressioni evidenti controllate?
