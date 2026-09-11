"""aga_lib - gjenbrukbare byggeklosser for AGA-analysen av Ført arbeid-eksport.

Pakken er bevisst delt opp i små, testbare moduler slik at samme logikk
(kolonnemapping, kommuneidentifikasjon, sonebestemmelse, linjeklassifisering
osv.) kan gjenbrukes fra `aga_analyse.py`, fra `tests/`, og fra fremtidige
kjøringer uten at noe må skrives om.
"""

__version__ = "1.0.0"
