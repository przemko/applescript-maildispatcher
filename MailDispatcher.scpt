-- MailDispatcher.scpt

set the WorkbookFile to choose file with prompt "Proszę wybrać plik Excela zawierający oceny:"
tell application "Microsoft Excel"
  open theWorkbookFile
  tell worksheet 1 of active workbook
    set courseNameCell to text returned of (display dialog "Podaj adres komórki zawierającej nazwę kursu:" default answer "B8")
    set courseName to value of cell courseNameCell
    set startRowNumber to text returned of (display dialog "Podaj numer wiersza, w którym rozpoczynają się dane o studentach:" default answer 12) as integer
    set indexNumberColumn to text returned of (display dialog "Podaj numer kolumny, w której znajdują się numery indeksów:" default answer 2) as integer
    set rowNumber to startRowNumber
    set cellValue to value of cell rowNumber of column indexNumberColumn
    repeat while cellValue is not ""
      set indexNumber to ((characters 5 thru -1 of cellValue) as string)
      set mark to value of cell rowNumber of column markColumn as number
      set emailAddress to indexNumber & "@student.pwr.wroc.pl"
      set message to "Informuję, że z kursu " & courseName & "otrzymał(a) Pan(i) ocenę " & mark & "
Pozdrawiam,
skrypt MailDispatcher"
      display dialog "Wysyłam na adres " & emailAddress & " list o treści:
" & message

      set rowNumber to rowNumber + 1
      set cellValue to value of cell rowNumber of column indexNumberColumn
    end repeat
  end tell
  close active workbook
end tell
