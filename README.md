# Klient poczty email

## Instrukcja uruchomienia

1. Instalacja wymaganych modułów, poprzez uruchomienie komendy `sh install.sh`.
2. Uruchomienie rozwiązania komendą `perl mail.pl`.
3. Po uruchomieniu skryptu, w konsoli pojawi się tekst :

```
To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code CRZMQTV8X to authenticate.
```

Należy przejść na stronę `https://microsoft.com/devicelogin` i wpisać kod wyświetlony w konsoli, a następnie zalogować się danymi użytkownika usługi Office 365.

4. Aplikacja zostaje uruchomiona

## Opis działania

1. Aplikacja wykorzystuje do uwierzytelnienia protokół OAuth2.0.
1. Integracja ze skrzynką pocztową opiera się o Microsoft Graph API.
1. W widoku głównym w pasku nawigacyjnym widoczny są następujące przyciski:
    * File
        * New - utworzenie nowego emaila
        * Exit - zamknięcie skryptu
    * Folders - lista folderów email dostępnych w skrzynce pocztowej użytkownika. Po wybraniu jednego z folderów, wyświetlona zostaje lista maili przechowywanych w danym folderze
1. Ekran główny wyświetla listę wiadomości email przechowywanych w wybranym w poprzednim kroku folderze. Wyświetla nadawcę, tytuł oraz czas wysyłki maila.
1. Po naciśnięciu na wierz z interesującym nas mailem, w nowym oknie zostaje otwarty ten email. Wyświetlone zostają między innymi tytuł maila (w pasku tytułowym okna), nadawca, odbiorca i zawartość maila (bez załączników).
    * Naciśnięciu przycisku `Reply` powoduje otworznie edytora maila.
    * Naciśnięcie przycisku `Delete` powoduje usunięcie maila z serwera.
