read_exel(ścieżka do pliku, nazwa arkusza=, nagłowek=) - padna odczytuje plik exela jako dataFrame
fillna('czym zatąpić puste komórki') - wypełnia puste komórki
applymap(funkcja którą stosuje do wszystkich komórek po kolei)
strip() - usuwa spacje przed i po wartościach w komórce
iloc[wiersze,kolumny] - wybiera określlone wiersze i kolumny z DF
iterrows() - iteruje po wierszach DF zwracając listy
enumerate(kolekcja po której iterujesz, start-od którego indexu zaczynasz)
lower() - zmienia wszystkie stringi na małe litery
DataFrame(kolekcja, columns=[nazwy kolumn])
.to_exel(ścieżka, index=)