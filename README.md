# prawo-jazdy

Skrypt do importowania bazy pytań egzaminacyjnych na prawo jazdy publikowanych przez Ministerstwo Infrastruktury do programu Anki.

## Pliki wejściowe

Źródłem danych są pliki ze strony Ministerstwa Infrastruktury: https://www.gov.pl/web/infrastruktura/prawo-jazdy

- [Baza pytań na prawo jazdy](https://www.gov.pl/documents/905843/1047987/pytania_plik_pa%C5%BAdziernik_2018.xlsx)
- [Multimedia cz. 1](https://www.gov.pl/documents/905843/1047987/cz%C4%99%C5%9B%C4%87_1.zip)
- [Multimedia cz. 2](https://www.gov.pl/documents/905843/1047987/cz%C4%99%C5%9B%C4%87_2.zip)
- [Multimedia cz. 3](https://www.gov.pl/documents/905843/1047987/cze%C5%9B%C4%87_3.zip)
- [Multimedia cz. 4](https://www.gov.pl/documents/905843/1047987/cz%C4%99%C5%9B%C4%87_4.zip)
- [Multimedia cz. 5](https://www.gov.pl/documents/905843/1047987/cz%C4%99%C5%9B%C4%87_5.zip)

## Przygotowanie plików do importu

- Należy ściągnąć wszystkie potrzebne pliki ze strony Ministerstwa Infrastruktury
- Zip-y z multimediami należy wypakować do jednego katalogu (bez podkatalogów)
- Zainstalować pythona 3 oraz dwie biblioteki: argparse i pyexcel_xlsx
- Ściągnąć plik `prawko2anki.py`

Skrypt uruchamiamy tak: `python prawko2anki.py -i pytania_plik_październik_2018.xlsx -m media -o out -c A`

Pomoc wyświetlana przez skrypt:

```
> python prawko2anki.py -h
usage: prawko2anki.py [-h] -i INPUT -m MEDIA -o OUTPUT -c CATEGORY

optional arguments:
  -h, --help            show this help message and exit
  -i INPUT, --input INPUT
                        Input XLSX file name
  -m MEDIA, --media MEDIA
                        Media directory
  -o OUTPUT, --output OUTPUT
                        Output directory
  -c CATEGORY, --category CATEGORY
                        License category
```

W katalogu wyjściowym powstanie plik `prawo-jazdy-pytania.txt`
oraz podkatalog `media` z filmikami i obrazkami do pytań.