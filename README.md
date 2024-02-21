# Mailscript

Skriptet lar brukeren sende én mail (emne, kropp) til flere brukere med forskjellige vedlegg. Vedlegg matches til mottaker gjennom recipients.csv
som ligger i en undermappe av der skriptet kjøres fra. For at skriptet skal kjøres riktig, skal mappestruktur se slik ut:

- Mailscript-hovedmappe
  - Mailscript.py
  - Templates/
    - .docx mal-filer som inneholder emne (første paragraf) og kropp (resterende paragrafer)
  - Recipients/
    - recipients.csv kobler kommunenummer og mail opp mot kommunenummer ivedlegg-filene.
  - FilesToSend/
    - filer du skal sende ut. Må inneholde kommunenummer i navn på fil eks: 
      1234_atkomstpunkt.sos
  - Signatures/
    - .docx filer med signatur som velges og legges til på slutten av mailen.

![Mappestruktur](https://github.com/jesperfjellin-kv/Mailscript/assets/145996132/39df15da-6012-4d13-810f-19b376d399ff)



**Nødvendige moduler med installasjon**

Os

Pandas (conda install pandas)

win32com (conda install pywin32)

python-docx (conda install conda-forge::python-docx)

tkinter (conda install anaconda::tk)

**Eksempelfiler**

Eksempelfiler er lagt med i repoet. Filene inneholder ingen identifiserende informasjon.
