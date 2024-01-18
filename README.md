# **Projekts**
Iztrādāja Alise Paula Lazdiņa (231RDB252) un Reinis Junga (231RDB211)
## Uzdevuma skaidrojums
Kā mūsu projektu nolēmām izstrādāt Python programmu, kura nolasa Excel faila "Klasu grafiks 2023_2024.xlsx" datus, ļauj apskatīt katras nedēļas dienas atlikušos brīvos laikus kādā no trijām filiālēm - Rīga, Mārupe, Ādaži, kā arī pierakstīties uz telpas izmantošanu, ievadot Excel failā datus.

Programma strādā sekojoši:
1. Lietotājs terminālī ievada sev vēlamo filiāli un dienu, kuras brīvos laikus vēlas apskatīt.
2. Terminālī tiek izvadīta attiecīgā, lietotāja izvēlētā informācija un tiek uzskaitīti konkrētās filiāles un dienas brīvie laiki.
3. Ja tiek ievadīta informācija, kas nav atrodama vai neatbilst Excel failā, tiek izvadīts kļūdas paziņojums.
4. Pēc informācijas apskates lietotājs var veikt izvēli - pieteikt nodarbību vai pabeigt darbu.
5. Izvēloties opciju "pieteikt nodarbību", lietotājs var izvēlēties laiku, kurā ievietot jaunu nodarbību (nepareiza laika formāta gadījumā tiek izvadīts kļūdas paziņojums "Ievadīts nederīgs laiks"). Lietotājam ir jāievada pasniedzēja vārds un uzvārds, jāizvēlas nodarbības veids - individuālā vai grupu. (Nederīga veida ievades gadījumā tiek izvadīts kļūdas paziņojums).
6. Pēc nodarbības veida izvēles, seko audzēkņa/audzēkņu vārdu, uzvārdu vai skaita ievade, (ievades kļūdas gadījumā, tiek izvadīts kļūdas paziņojums).
7. Pēc nepieciešamās informācijas ievadīšanas, informācija tiek saglabāta esošajā Excel failā un darbs tiek pabeigts.

Būtiski pieminēt, ka visas ievades vērtības ir standartizētas pēc noteiktiem principiem, kas nepieļauj neatbilstošu datu ievadi. 

## Izmantotās bibliotēkas
Programmas izstrādei izmantojām bibliotēku "openpyxl".
Šī bibliotēka ir paredzēta darbam ar Excel failiem, tā sniedz iespēju lasīt un rakstīt Excel failus. "Openpyxl" bibliotēkai ir dažādas klases un funkcijas. Mūsu projekta izstrādei izmantojām klasi "Workbook" un funkciju "load_workbook", ko izmantojām, lai ielādētu Excel failu, apstrādātu un darbotos ar failā esošo informāciju.

Bibliotēka "sys" sniedz iespēju izmantot dažādas funkcijas un mainīgos, lai varētu darboties ar dažādām izpilddarbībām. Mūsu projektā šī bibliotēka tiek izmantota, lai apstādinātu programmu noteiktos punktos, kad tas ir nepieciešams.

## Izmantošanas metodes
Mūsu izstrādātais projekts un programma būtiski atvieglina informācijas meklēšanu, apskati un automatizē Excel faila darbību. Programmas lietotājam nav nepieciešams pašam meklēt failu, manuāli to atvērt un atrast sev interesējošo informāciju, jo visa nepieciešamā informācija tiek izvadīta ekrānā un turpat, automātiski ievadot datus, var pieteikt nākamo nodarbību. 

Uzņēmumos parasti ir telpas, kuru izmantojamība ir jāpiesaka iepriekš. Šī programma var lieliski noderēt, jo ļauj ātri un viegli pārskatīt informāciju, noskaidrot telpu noslogojamību, kā arī pieteikt telpas izmantošanu.

Projektā izmantotais Excel fails ir noderīgs, jo to ir iespējams pielāgot savām vajadzībām, kā arī izstrādāt citas programmas, kas sniedz iespēju ātri, ērti un vienuviet veikt visas nepieciešamās darbības, kā arī, papildinot kodu, to var izmantot datu analīzē, lai noskaidrotu jebkāda priekšmeta, objekta, telpas amortizāciju jeb nolietojumu.
