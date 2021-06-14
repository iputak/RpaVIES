# RpaVIES
VIES Skripta je računalni program koji se izvršava samostalno u svrhu višestruke provjere valjanosti PDV identifikacijskih brojeva. Program je kreiran 2021. godine.

## RPA
Robotska automatizacija procesa (engl. Robotic process automation) je oblik tehnologije automatizacije poslovnih procesa koja se temelji na računalnim programima koji se izvršavaju samostalno (engl. bot) ili na umjetnoj inteligenciji (engl. Artificial intelligence - skraćeno AI). Ponekad se naziva softverskom robotikom koju ne treba miješati sa softverom robota.

Kratko predavanje o RPA: [What is RPA (Robotic Process Automation)?](https://www.youtube.com/watch?v=aZDaNVh3l0k)
## VIES
VIES je sustav za elektronički prijenos informacija o valjanosti PDV identifikacijskih brojeva te o isporukama dobara i usluga na zajedničkom tržištu Europske unije.

Referenca: [Porezna uprava](https://www.porezna-uprava.hr/PdviEu/Stranice/Naj%C4%8De%C5%A1%C4%87e-postavljena-pitanja.aspx#p1)

## KORACI PROGRAMA
1.	Dohvaćanje Excel datoteke
2.	Pronalazak listova
3.	Zbrajanje zapisanih redova u Request Member State listu
4.	Spremanje URL, Member State i VAT number iz Excela u pomoćne varijable
5.	Ulazak u petlju koja se ponavlja koliko ima podataka, tj. redova u tablici
  *	Spremanje Requester Member State, Requester VAT number iz Excela u pomoćne varijable
  *	Otvaranje internet stranice u Chrome pregledniku
    *	Dohvaćanje elemenata otvorene Internet stranice
    *	Upis podataka iz pomoćnih varijabli u elemente otvorene Internet stranice
    *	Prijava
    *	Dohvaćanje Consultation Numbers
      *	Ako Consultation Numbers ne postoji, u pomoćnu varijablu se zapisuje '//'
    *	Spremanje stranice kao pdf dokument pod prilagođenim imenom
  *	Zatvaranje preglednika
  *	Zapis podataka u Excel dokument na list Results
  *	Spremanje dokumenta
6.	Kraj programa


## EXCEL DOKUMENT
Svi potrebni podatci za rad skripte se nalaze u Excel dokumentu naziva vatNumbers.xlsx koji se nalazi u mapi naziva excel.
Dokument se sastoji od tri lista:
1.	Member State
2.	Request Member State
3.	Results
