'WScript.Echo("Create the FwPolicy object.") виводить вікно повідомлення     з текстом написаним в дужках 
'set - команда для створення змінних
'оператор Dim обявляє змінну

'інструкція Set Объявляет Set процедуру, которая используется для присвоения значения свойству
Set FSO = CreateObject("Scripting.FileSystemObject")    'обєкт для доступа до файлової системи
Set Folder = FSO.GetFolder("D:\replenishment of the warehouse")          'отримуємо папку в змінну

Dim mon 'змінна для місяця
mon = Month(Date) 'виділяємо номери місяця із дати
Dim days 'змінна для дня
days = Day(Date) 'виділяємо номер дня із дати

'Set TextStream = FSO.CreateTextFile("D:\TestFolder\Test.txt", True)      'створюємо новий файл
Set TextStream = FSO.CreateTextFile("D:\replenishment of the warehouse\Log\ListOfGroups" & "_" & Date & ".txt", True)

TextStream.Write("Today's date: ")
TextStream.Write(Date & vbCrLf) 'записуємо у файл дату запуска файла

TextStream.Write("List of groups of goods you need to refill: " & vbCrLf) 'запис в файл
TextStream.Write(vbCrLf)
'& - логічний оператор і
'vbCrLf - возврат каретки і перевод нового рядка
Str = vbNullString

For Each SubFolder In Folder.SubFolders                 'цикл який перебирає всі вкладені папки в корневій папкі
'тут пишиться умова if і якщо вона true то назва папки додається в список
    If DateDiff("d", SubFolder.DateLastModified, Date) > 60 Then
        TextStream.Write("- " & SubFolder.Name & " " & vbCrLf)
    End If
Next
Set File = FSO.GetFile("D:\replenishment of the warehouse\Log\ListOfGroups" & "_" & Date & ".txt")
Set TextFile = File.OpenAsTextStream(1) 'відкрити файл для зчитування
WScript.Echo TextFile.ReadAll()
TextFile.Close  'закриває відкритий текстовий файл