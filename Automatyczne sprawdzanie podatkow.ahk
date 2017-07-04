#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#SingleInstance Force

PoliczLiniePliku(plik) ;Liczy liczbę lini w pliku
{
	ostatniaLinia = 0
	
	Loop, Read, %plik%
		ostatniaLinia++
	
	return ostatniaLinia
}

Potega(liczba, wykladnik) ;Potęguję daną liczbę do naturalnej potęgi
{
	wynik := 1
	
	Loop %wykladnik%
	{
		wynik *= liczba
	}
	
	return wynik
}

IELoad(wb)    ;You need to send the IE handle to the function unless you define it as global.
{
    If !wb    ;If wb is not a valid pointer then quit
        Return False
    Loop    ;Otherwise sleep for .1 seconds untill the page starts loading
        Sleep,100
    Until (wb.busy)
    Loop    ;Once it starts loading wait until completes
        Sleep,100
    Until (!wb.busy)
    Loop    ;optional check to wait for the page to completely load
        Sleep,100
    Until (wb.Document.Readystate = "Complete")
Return True
}

IEGet(Name="")        ;Retrieve pointer to existing IE window/tab with a name specified as a parametr
{
    IfEqual, Name,, WinGetTitle, Name, ahk_class IEFrame
        Name := ( Name="New Tab - Windows Internet Explorer" ) ? "about:Tabs"
        : RegExReplace( Name, " - (Windows|Microsoft) Internet Explorer" )
    For wb in ComObjCreate( "Shell.Application" ).Windows
        If ( wb.LocationName = Name ) && InStr( wb.FullName, "iexplore.exe" )
            Return wb
} ;written by Jethrow

Koniec()		;Funkcja wywoływana gdy użytkownik naciśnie jakiś klawisz na klawiaturze w trakcie dziłania programu
{
	blokada = 0
	Progress, Off
	MsgBox, 262208, Koniec, Program został zatrzyany
	ExitApp
	return
}

SprawdzNIP(nip) ;Oblicza sumę kontrolną NIP. Zwraca 1 gdy poprawna i 0 gdy nie poprawna
{	
	if (StrLen(nip) != 10) ;Gdyby NIP nie miał 10 cyfr to na pewno jest nie prawidłowy
		return 0
	
	cyfryNIP := Object() ;Tworzy tablicę asocajacyjną
	Loop, 10 ;Dzieli NIP na cyfry
	{
		cyfryNIP.Insert(Mod(Floor(nip / Potega(10, (10 - a_index))),10))
	}
	
	suma := 0
	wagi := [6, 5, 7, 2, 3, 4, 5, 6, 7]
	Loop, 9 ;Oblicza sumę kontroną mnożąc cyfry przzez odpowiedznie wagi
	{
		suma += (cyfryNIP[a_index] * wagi[a_index])
	}
		
	if (Mod(suma, 11) = cyfryNIP[10]) ;Sprawdzenei czy (suma cyfr * wagi) mod 11 = ostatnia cyfra
		return 1
	else
		return 0
}

global versjaOficjalna = 0 ;Czy ma pracować jako wersja oficjalna - 1, czy jako wersja do testów - 0
if A_IsCompiled ;Gdy skrypt jest skompilowany to zawsze jest wersją oficjalną
	versjaOficjalna = 1

;Następne zmienne są nadpisywane gdy wersja oficjalna = 1!
global blokada = 1 ;Czy wciśnięcie klawisza na klawiaturze ma przerywać program
global pasekPostepu = 1 ;Czy ma być wyświetlany pasek postępu
global przerobWszystkieReordy = 0 ;Czy ma pracować na wszystkich rekordach. Gdy 0 używa liczby rekordów z liczbaRekorkowDoZrobienia
global sprawdzeniePlikow = 1 ;Czy poprawność danych ma być sprawdzona
global restartPrzegladarki = 0 ;Czy ma wymuszać otwarcie nowej instancji explorera
global wylapujPowtorki = 1 ;Czy ma pomijać rekordy, jeżeli odpowiednie wyniki już istnieją
global NIPZamiastVendor = 0 ;Czy zastąpić numer Vendor numerem NIP podczas nazywania plików

liczbaRekorkowDoZrobienia = 20 ;Limit rekordów do wykonania - TESTY


;POLE TESTÓW

;MsgBox, % ComObjCreate("Scripting.FileSystemObject").GetFolder((A_ScriptDir . "\wyniki\poprawne")).Files.Count ;!!! Liczy ile jest plików w folderze. Może się przydać do kontroli czy prpgram działa prawidłowo

;~ MsgBox, % Potega(2, 6)


;~ ExitApp


;KONIEC POLA TESTÓW

if versjaOficjalna ;Dopasownie zmiennych i początkowe czynności gdy versja oficjalna
{
	wylapujPowtorki = 1 ;Automatycznie pomija dane jeśli już wcześniej zostały przetworzone
	blokada = 1 ;Wciśnięcie dowolnego klawisza zatrzymuje działanie programu
	pasekPostepu = 1 ;Pastek postępu będzie widoczny
	sprawdzeniePlikow = 1 ;Pliki źródłowe zostaną sprawdzone
	przerobWszystkieReordy = 1 ;Pracuje na całości danych
	restartPrzegladarki = 1 ;Zamyka istniejącą kartę ze stroną do sprawdzania NIP
	NIPZamiastVendor = 0 ;Nie zastępuje numeru Vendor numerem NIP chyba, że jest to konieczne
}

if restartPrzegladarki ;Wymusza otwarcie IE na nowo. Zapezpieczenie na wypadek wygaśnięcia sesji
{
	WinClose, Portal Podatkowy - Internet Explorer
}

if pasekPostepu ;Zainicjuj pasek postępu
{
	Progress, b w250, 0`%, Postęp, Pasek postępu ;Rysuje pasek postępu z 0%
	Progress, 0 ; Set the position of the bar to x%.
}

if sprawdzeniePlikow ;Sprawdza czy istnieją pliki na dane
{ 
dane := [1,1,1] ;Tablica na które pliki są poprawne

IfNotExist %A_ScriptDir%\vendor.txt ;Vendor - 1. plik
	dane[1] := 0
IfNotExist %A_ScriptDir%\name.txt ;Name - 2. plik
	dane[2] := 0
IfNotExist %A_ScriptDir%\vat_number.txt ;Vat number - 3. plik
	dane[3] := 0

if((dane[1] = 0) or (dane[2] = 0) or (dane[3] = 0)) ;Jeśli nie istnieją to je tworzy
{
	blokada = 0 
	Progress, Off
	
	MsgBox, 262180, Dane nie gotowe, Brakuje plików na dane. Czy mam je utworzyć? ;Prompt czy utworzyć brakujące pliki
	IfMsgBox, Yes ;Jeśli użytkownik kliknie tak to tworzy te pliki, których brakuje
	{
		if (dane[1] = 0)
			FileAppend, ,vendor.txt
		if (dane[1] = 0)
			FileAppend, ,name.txt
		if (dane[1] = 0)
			FileAppend, ,vat_number.txt
		
		MsgBox, 64, Gotowe, Pliki na dane zostały utwożone. Należy je wypełnić i uruchomić ponownie program
		ExitApp
	}
	IfMsgBox, No 
	{
		ExitApp
	}
}
}

if sprawdzeniePlikow ;Liczenie liczby linii w plikach
{ 
	nameLiczbaLini := PoliczLiniePliku("name.txt")
	vendorLiczbaLini := PoliczLiniePliku("vendor.txt")
	vat_numberLiczbaLini := PoliczLiniePliku("vat_number.txt")
	
	if (nameLiczbaLini = vat_numberLiczbaLini) ;Porównanie liczby linii w różnych plikach
	{
		if ((vendorLiczbaLini = 0) AND (nameLiczbaLini != 0))
		{
			blokada = 0
			Progress, Off
			MsgBox, 262180, Niepełne dane, Brakuje numerów vendor w plikach z danymi.`nCzy zamiast nich użyć numerów NIP do nazywania plików?
			IfMsgBox, Yes
			{
				blokada = 1
				NIPZamiastVendor = 1
			}
			IfMsgBox, No
			{
				MsgBox, 16, Błąd, Liczba lini w plikach z danymi jest różna! Należy sprawdzić dane`n`nMoże to być spowodowane pustą linią na końcu któregoś z plików
			ExitApp
			}
			
		}
		else if (nameLiczbaLini != vendorLiczbaLini)
		{
			blokada = 0
			Progress, Off
			MsgBox, 16, Błąd, Liczba lini w plikach z danymi jest różna! Należy sprawdzić dane`n`nMoże to być spowodowane pustą linią na końcu któregoś z plików
			ExitApp
		}
	}
	else
	{
		blokada = 0
		Progress, Off
		MsgBox, 16, Błąd, Liczba lini w plikach z danymi jest różna! Należy sprawdzić dane
		ExitApp
	}
}

if przerobWszystkieReordy ;Pracuj na całości danych
{
	liczbaRekorkowDoZrobienia := vat_numberLiczbaLini ;Przerabia wszystkie rekordy
}

{ ;Tworzenie folderów na wyniki
IfNotExist, `"%A_ScriptDir%\wyniki`" ;Tworzy folder na wyniki jeśli nie istnieje
{
	try
	{
		FileCreateDir, wyniki
	}
	catch ;Gdyby wystąpił błąd
	{
		blokada = 0 ;Program nie będzie narzekał na kliknięcie klawiszy
		Progress, Off ;Znika pasek postępu
		MsgBox, 16, Błąd, Nie udało się utworzyć folderu na wyniki
		ExitApp
	}
}

IfNotExist, `"%A_ScriptDir%\wyniki\poprawne`" ;Tworzy folder na poprawne wyniki jeśli nie istnieje
{
	try
	{
		FileCreateDir, %A_ScriptDir%\wyniki\poprawne
	}
	catch ;Gdyby wystąpił błąd
	{
		blokada = 0 ;Program nie będzie narzekał na kliknięcie klawiszy
		Progress, Off ;Znika pasek postępu
		MsgBox, 16, Błąd, Nie udało się utworzyć folderu na wyniki
		ExitApp
	}
}

IfNotExist, `"%A_ScriptDir%\wyniki\nie poprawne`" ;Tworzy folder na nie poprawne wyniki jeśli nie istnieje
{
	try
	{
		FileCreateDir, %A_ScriptDir%\wyniki\nie poprawne
	}
	catch ;Gdyby wystąpił błąd
	{
		blokada = 0 ;Program nie będzie narzekał na kliknięcie klawiszy
		Progress, Off ;Znika pasek postępu
		MsgBox, 16, Błąd, Nie udało się utworzyć folderu na wyniki
		ExitApp
	}
}
}

IfNotExist, %A_ScriptDir%\nircmd\nircmd.exe ;Gdy nie ma programu do zrzutów ekranu to o tym powiadomi
{
	Progress, Off
	blokada = 0
	
	MsgBox, 262160, Błąd, Brakuje programu do robienia zrzutów ekranu!`n`nProgram nircmd.exe powinien być w folderze nircmd w tym samym miejscu co główny program. Jeśli go braku należy ponownie wypakować paczkę z głównym programem lub pobrać nircmd.exe ze strony: http://www.nirsoft.net/utils/nircmd.zip i wypakować w folderze z głównym programem.
	ExitApp
}


;Zmienne liczące dane
poprawneRaporty := 0 ;Gdy wynik jest standardowy
niePoprawneRaporty := 0 ;Gdy odpowiedź strony jest różna niż standardowa
powrorki := 0 ;Gdy danyc wynik był już utworzony wcześniej
bledneNIP := 0 ;Gdy numer NIP ma błędną sumę kontrolną
procenty := 0 ;Ile % rekordów zostało już przerobionych

czytanaLinia := 1 ;Ktróra linia plików jest aktualnie przerabiana
Loop
{	
	;Porusza myszką w tę i z powrotem by komputer nie poszedł spać
	LastX := CurrentX
    LastY := CurrentY
    MouseGetPos, CurrentX, CurrentY
    If (CurrentX = LastX and CurrentY = LastY) {
        MouseMove, 1, 1, , R
        Sleep, 100
        MouseMove, -1, -1, , R
    }
	
	
	NIPNieWBazie := 0
	
	
	if pasekPostepu ;Wyświetlenie paska postępu
	{
		procenty := Floor(((czytanaLinia -1) / liczbaRekorkowDoZrobienia) * 100)
		Progress, b w250, % procenty "`%`n" czytanaLinia - 1 "/"  liczbaRekorkowDoZrobienia, Postęp, Pasek postępu ;Tekst na pasku
		Progress, %procenty% ;Długość zielonego paska
	}
	
	if(czytanaLinia > liczbaRekorkowDoZrobienia) ;Zakończenie pracy gdy przerobiona zostanie całość/dana liczba rekordów
	{
		temp := czytanaLinia - 1 ;Zmienna tymczasowa
		blokada = 0 ;Program nie będzie narzekał na kliknięcie klawiszy
		Progress, Off ;Znika pasek postępu
		MsgBox, 0, Gotowe, Praca skończona`nLiczba wykonancyh działań: %temp%`n`nLiczba poprawnych: %poprawneRaporty%`nLiczba nie będących w bazie: %niePoprawneRaporty%`nLiczba błędnych numerów NIP: %bledneNIP%`nLiczba już istniejących: %powrorki%
		ExitApp
	}
	
	try ;Czytanie plików
	{
		if !NIPZamiastVendor
			FileReadLine, vendor, vendor.txt, %czytanaLinia%
		FileReadLine, firmName, name.txt, %czytanaLinia%
		FileReadLine, vatNo, vat_number.txt, %czytanaLinia%
	}
	catch ;Gdyby był błąd w trakcie czytania
	{
		blokada = 0 ;Program nie będzie narzekał na kliknięcie klawiszy
		Progress, Off ;Znika pasek postępu
		MsgBox, 16, Błąd, Błąd w trakcie wczytywania danych
		ExitApp
	}

	;Zmienna do nazywania plików
	if NIPZamiastVendor
		nazwa := vatNo . " " . firmName ;Utowrzenie nazwy do podpisywania wyników w przypaku gdy zostanie wybrana opcja pracy bez numeru Vendor
	else
		nazwa := vendor . " " . firmName ;Utowrzenie nazwy do podpisywania wyników
	
	if wylapujPowtorki ;Wyłapuje powtórki. Pzechodzi wtedy do następnej wartości.
	{
		IfExist, %A_ScriptDir%\wyniki\poprawne\%nazwa%.png ;Dla poprawncyh
		{
			czytanaLinia := czytanaLinia + 1
			powrorki := powrorki + 1
			continue
		}
		IfExist, %A_ScriptDir%\wyniki\nie poprawne\NIE POPRAWNY %nazwa%.png ;Dla nie poprawnych
		{
			czytanaLinia := czytanaLinia + 1
			powrorki := powrorki + 1
			continue
		}
		IfExist, %A_ScriptDir%\wyniki\nie poprawne\BŁĘDNY NIP! %nazwa%.txt ;Dla NIP-ów ze złą sumą kntrolną
		{
			czytanaLinia := czytanaLinia + 1
			powrorki := powrorki + 1
			continue
		}
	}
	
	if (!SprawdzNIP(vatNo)) ;Sprawdza czy dany numer NIP jest poprawny
	{
		nazwa := "BŁĘDNY NIP! " . nazwa
		
		FileOpen(A_ScriptDir . "\wyniki\nie poprawne\" . nazwa ".txt", "w").Write(vatNo . "`r`n").Close ;Jeśli nie jest poprawny to tworzy plik z informacją
		
		bledneNIP := bledneNIP + 1
		czytanaLinia := czytanaLinia + 1
		continue
	}
	
	
	IfWinExist, Portal Podatkowy - Internet Explorer ;Gdy przeglądarka nie jest włączona włącza ją
	{
		wb := IEGet("Portal Podatkowy") 
	}
	IfWinNotExist, Portal Podatkowy - Internet Explorer ;Gdy jest już odpalona to używa gotowego okna
	{
		wb := ComObjCreate("InternetExplorer.Application")
		wb.Visible := True
	}

	WinActivate, Portal Podatkowy - Internet Explorer ;Wysuwa okna na przód i aktuwuje je
	WinShow, Portal Podatkowy - Internet Explorer

	wb.Navigate("http://www.finanse.mf.gov.pl/web/wp/pp/sprawdzanie-statusu-podmiotu-w-vat") ;Odpala stronę od sprawdzania podatków
	IELoad(wb) ;Czeka na stronę aż się załaduje
	
	WinWait, Portal Podatkowy - Internet Explorer
	WinMaximize, Portal Podatkowy - Internet Explorer ;Otwiera okno na pełny ekran. Program prawdopodobnie działałby bez tego, ale tak ładniej wygląda. Może poprawić niezawodność w przypadku implementacji klikania współżędnych na ekranie

	

	;Wersja z poleceniami bezpośrednio do przeglądarki. Działa wyśmienicie :D
	Loop ;Wpisuje NIP w odpowiednie pole dopuki nie zostanie wpisany
	{
		wb.Document.getElementById("b-7").value := vatNo
		
		if wb.Document.getElementById("b-7").value = vatNo
			break
	}
	wb.Document.getElementById("b-8").Click() ;Klika "Sprawdź"
	
	Loop ;Czeka aż pojawi sie jakieś powiadomienie
	{
		if((wb.Document.getElementById("caption2_b-3").innertext != "") && (wb.Document.getElementById("caption2_b-3").innertext != "False")) ;Pole to przyjmuje wartości "" i "False" zanim przyjmnie wartość wyniku przed podaniem odpowiedzi 
		{
			break
		}
	}
	
	if(StrLen(wb.Document.getElementById("caption2_b-3").innertext) != 339) ;Gdyby komunikat był inny niż prawidłowy - Prawidłowy komunikat ma 339 znaków :D Są w nim nowe linie i nie za barzo wiem jak go wpisiać w kod. Rozwiązanie na liczbę znaków działa bardzo dobrze
	{
		NIPNieWBazie := 1
		nazwa := "NIE POPRAWNY " . nazwa
		
	}
	
	;Zapis jako zrzut ekranu
	Progress, Off
	if NIPNieWBazie ;Wzależności od wyniku strony jest zapisywany w odpowiednim oflderze
	{
		Run, `"%A_ScriptDir%\nircmd\nircmd.exe`" savescreenshotwin `"%A_ScriptDir%\wyniki\nie poprawne\%nazwa%.png`"
	}
	else
	{
		Run, `"%A_ScriptDir%\nircmd\nircmd.exe`" savescreenshotwin `"%A_ScriptDir%\wyniki\poprawne\%nazwa%.png`"
	}
	
	Sleep, 500 ;Potrzebne by na zrzucie ekranu nie było paska postępu. Diała już przy 200
	
	if NIPNieWBazie
		niePoprawneRaporty := niePoprawneRaporty + 1
	else
		poprawneRaporty := poprawneRaporty + 1
	
	if pasekPostepu ;Aktualizacja paska postępu
	{
		procenty := Floor((czytanaLinia / liczbaRekorkowDoZrobienia) * 100)
		Progress, b w250, %procenty%`%`n%czytanaLinia%`/%liczbaRekorkowDoZrobienia%, Postęp, Pasek postępu
		Progress, %procenty%
	}
	
	czytanaLinia := czytanaLinia + 1
}

#If blokada ;Kontrola user input
{ ;Dowolny klawisz konczy program
$a::
$b::
$c::
$d::
$e::
$f::
$g::
$h::
$i::
$j::
$k::
$l::
$m::
$n::
$o::
$p::
$q::
$r::
$s::
$t::
$u::
$v::
$w::
$x::
$y::
$z::
$+A::
$+B::
$+C::
$+D::
$+E::
$+F::
$+G::
$+H::
$+I::
$+J::
$+K::
$+L::
$+M::
$+N::
$+O::
$+P::
$+Q::
$+R::
$+S::
$+T::
$+U::
$+V::
$+W::
$+X::
$+Y::
$+Z::
$`::
$!::
$@::
$#::
$$::
$^::
$&::
$*::
$(::
$)::
$-::
$_::
$=::
$+::
$[::
${::
$]::
$}::
$\::
$|::
$;::
$'::
$<::
$.::
$>::
$/::
$?::
$enter::
$space::
$tab::
$CapsLock::
$backspace::
$1::
$2::
$3::
$4::
$5::
$6::
$7::
$8::
$9::
$0::

Koniec()
return
}
#If
