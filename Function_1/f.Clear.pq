/*
 fClear - функция очистки текста

Удаляет из всех колонок таблцы:
- любые символы в т.ч. непечатные (ПОСИМВОЛЬНО в текстовой строке "абвг0123" или "#(lf)#(00A0)" или двойные пробелы "  ")
- пустые колонки
- пустые строки

  Пример
  Аргументы: таблица + удалить непечатные + удалить пустые столбцы + удалить пустые строки                                                                                                                        
  fxClear = fxClear(Source, "#(lf)#(00A0)", 1, 1),  - удалит только 2 непечатных символа
  fxClear = fxClear(Source, " ", 1, 1),             - удалит только двойные пробелы. Можно так "" - тоже удаляет
  fxClear = fxClear(Source, "123", 1, 1),           - удалит во всей таблице символы 1, 2 и 3
  fxClear = fxClear(Source, null, 1, 1),            - ничего не удалит
*/
  
let 
fxClear = (
    tbl, 						// таблица
    optional clearData, 		// удалить пробел + непечатные символы: " #(lf)#(00A0)" или "" удалит только двойные пробелы ли 
    optional clearColumn, 		// удалить пустые столбцы
    optional clearRows			// удалить пустые строки
  ) =>

    [ 
      // делаем все колонки {} текстовыми
      tr = Table.TransformColumns(tbl, {}, Text.From),
      
      // удаляем непечатные символы (активируется, если аргумент неравен null, т.е. любой символ)
      p =   if   clearData is null 
			then tr
			else [ k  = (x) => 
						if   x = null 
						then x 
						else Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x), clearData), {""}), " "),
				   to = Table.TransformColumns(tr, {}, k)
				 ][to],`

      // удаляем пустые колонки
      cl =  if clearColumn is null 
			then p 
			else [ a = List.Sort(Table.ColumnNames(p)),
				   b = Table.Profile(p, {{"tmp", (x) => true, List.NonNullCount}})[tmp],
				   c = List.PositionOf(b, 0, Occurrence.All),
				   d = List.Transform(c, (x) => a{x}),
				   e = Table.RemoveColumns(p, d)
				 ][e],

      // удаляем пустые строки
      to =  if   clearRows is null 
			then cl 
			else Table.SelectRows(cl, each not List.IsEmpty( List.RemoveMatchingItems( Record.FieldValues(_), {"", null})))
  		
    ][to]
in 
	fxClear


/*
-- ПРИМЕРЫ еще

-- убираем нули перед цифрами. Текст не трогает. Если текст то ошибка, которая возвращает текст (otherwise _)
= Table.TransformColumns(fxClear,{"Column2", each try Number.From(_) otherwise _})

-- убираем двойные пробелы во всей таблице
	  w = [ f = (this) => Text.Combine(List.Select(Text.Split(this, " "), each _ <> ""), " "),
	        m = Table.TransformColumns(to, {}, f)
		  ][m]

-- убираем  нули в ячейках
 Replace0ToNull = Table.ReplaceValue(
        RenameColumns, 
        0, 
        null, 
        Replacer.ReplaceValue, 
        {"Текущий период", "Прошлый период"}
    ),
	
-- работа внутри таблицы	
	let
    Источник = ВсеФормыВпапке2[[Name], [Content]],
    f=(x)=>
        [ a1 = x{4}[Data],
          a2 = Table.Skip(a1, 16),
          a3 = Table.RemoveLastN(a2,11),
          a4 = Table.SelectColumns(a3,{"Column1", "Column6", "Column7", "Column11"}),
          a5 = Table.AddIndexColumn(a4, "Index", 1, 1, Int64.Type)  // чтобы строки потом сортировалось правильно от номера 100 и выше НЕПОЛУЧИЛОСЬ
        ][a4],
    
    tr = Table.TransformColumns(Источник,{"Content",f})
in
    tr
	
*/