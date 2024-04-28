// Статотчет ОИСК

let
	/* ШАГ 0 - Функция Load */
	
	Load = (fnName as text, optional BasePath as text) as function =>
		let
			DefaultPath = "G:\=EXCEL PQ\PQ-function-load\",
			GitHubPath  = "https://raw.githubusercontent.com/hohlick/pquery/master/",
				
			BasePath = if (BasePath <> null) then BasePath else DefaultPath,
			Path = BasePath & (if Text.End(BasePath, 1) <> "\" then "\" else ""),
			File = Path & fnName & ".pq",
			
			Function = try Expression.Evaluate(Text.Replace(fnName, ".", "_"), #shared)  //если путь уже импортирован в рабочую книгу
			otherwise try  Expression.Evaluate(Text.FromBinary(Binary.Buffer(File.Contents(File))), #shared) //если еще не импортирован, попробуйте загрузить его из текстового файла в папке
			otherwise      Expression.Evaluate(Text.FromBinary(Binary.Buffer(Web.Contents(GitHubPath & fnName & ".pq"))), #shared) // если папка не найдена - взять из GitHubPath
		in	Function,
	
    
	/* ШАГ 1 - ПОДКЛЮЧЕНИЕ к источнику данных */
    
    // подключаемся к папке
    Path  = Value.Metadata(PathFolder)[List]{1},
    
    LoadFile = 
        let from  = Folder.Files(Path),
            filtr = Table.SelectRows(from, each [Extension] = ".xls" or [Extension] = ".xlsx" and [Attributes][Hidden] = false),
            tbl   = Table.SelectColumns(filtr,{"Name","Content"})
        in  tbl,

    // выбираем нужные листы в файле 
    tr = Table.TransformColumns(LoadFile,
        {"Content", (x)=>
            let a = Excel.Workbook(x),                                  // соджержимое файла, если лист в файле один, то (x)=>Excel.Workbook(x){1}[Data]}
                b = Table.SelectRows(a, each Text.Length([Name]) = 15), // выбираем нужный лист (нам нужен с кодом организации в названии листа = 15 знаков)
                c = b{0}[Data]
            in  c
        }),

    // добавили колонки с доинформацией о файле/отчете
    AddColumns = 
        let f = (tbl as table, addcol as list) as table => List.Accumulate(addcol, tbl, (t, d) => Table.AddColumn(t, d{0}, d{1}, d{2}?)),
            // т.к в файле xls дата = 01.16, в xlsx = 01.12.2016, то возникает ошибка => обрабатываем ее через try, если верно и otherwise, если ошибка
            AddedColumns = f(tr, {
                { "Date",           each try       Date.From("01.01.20" & Text.End([Content][Column1]{0},2))             // если файл xls
                                         otherwise Date.From([Content][Column1]{0}+ #duration(30, 0, 0, 0)), type date}, // если файл xlsx, то добаваляем 30 дней
                { "ID_Предприятие", each Text.From([Content][Column1]{1}), type text },
                { "ID_Форма",       each Text.Start([Content][Column2]{0},4), type text}
            })
        in  AddedColumns,
    
    // справочник форм отчетности
    Спр_ФормыОтчетности = 
        let Источник = #table(type table [ID_Форма = text, Название = text, КраткоеНазвание = text, Примечание = text],{     
                    {"0112","Баланс","ББ",""}, 
                    {"0212","Отчет и прибылях и убытках", "ОПУ",""},
                    {"0312","Отчет о доходах, расходах и финансовых результатах","ОДРФР",""},
                    {"0412","Отчет о движении денежных средств","ОДДС",""},
                    {"0512","Приложение к бухгалтерскому балансу","ПББ",""},
                    {"3012","Отчет об изменении капитала","ОИК",""}
                })
        in Источник,

    // Добавляем колонки через объединение таблиц
    JoinExpand_1 =
        let a = Table.NestedJoin(AddColumns,"ID_Форма",Спр_ФормыОтчетности,"ID_Форма","tmp"),
            b = Table.ExpandTableColumn(a,"tmp", {"Название"})
        in  b,
        
    Filter = Table.SelectRows(JoinExpand_1, each ([ID_Форма] = "3012")) [[Date],[Content]],
    
    to = Table.Buffer(Filter),
          
    // ЧАСТЬ 3 
  
    // Функция чистки таблицы. fClear использовать набо бы
    f = (x)=>
            let a = Table.Skip(x,2), // пропускаем 2 строки (из за кода организации выдает ошибку, обходим)
                b = Table.TransformColumnTypes(a,{{"Column2", Int64.Type}}), // делаем тип колонки КодСтроки ЦелоеЧисло
                c = Table.RemoveRowsWithErrors(b, {"Column2"}), // удаляем из колонки КодСтроки ошибки, т.к. в ней есть текстовые данные
                d = Table.SelectRows(c, each [Column2] <> null and [Column2] > 2), // удаляем все строки кроме строк, с номерами статей
                e = Table.AddIndexColumn(d, "Index", 1, 1),  // добавляем колонку индекс
                // ClearData = Load("fxClear")(e, " #(lf)#(00A0)", true, true)
                k=(x)=> try Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " ") otherwise x,
                ClearData = Table.TransformColumns(e, {}, k)
            in  ClearData,
    
    // Добавляем новую колонку в подтаблицами
    AddColumn2 = Table.AddColumn(to, "Content2", each f([Content])),

    // Определяем список заголовков - 10 
        lstHeader = 
            {   
                "Статья", "ID_Строка",
                "Уставный капитал", "Неоплаченная часть уставного капитала",
                "Собственные акции (доли в уставном капитале)", "Резервный капитал", "Добавочный капитал",
                "Нераспределенная прибыль (непокрытый убыток)", "Чистая прибыль (убыток)",
                "Итого"
            },

    // Разворачиваем таблицу
        Expand = Table.ExpandTableColumn(AddColumn2[[Date],[Content2]], 
                "Content2", 
                Table.ColumnNames(AddColumn2[Content2]{0}), // берем список имен колонок с первой таблицы, 10 штук
                lstHeader & {"Index"}), // + 1 колонка

    // Делаем заголовки (кроме первого "Статья") числовыми (всего 10+1, последние 10)
        TransformColumnTypes = Table.TransformColumnTypes(
            Expand, 
            List.Transform(
                List.LastN(lstHeader & {"Index"}, 10), // считая по номеру с конца списка
                each {_, type number} 
            )
        ),
    
    // Сортируем колонку "ID_Строка", вторая в списке
        Sort = Table.Sort(TransformColumnTypes, { {lstHeader{1}, Order.Ascending} } ),

     // Делаем множественную замену в подстроке столбца
        МультиЗамена =  
            let func = (x)=> List.Accumulate(dict, x, (state, current) => Text.Replace(state, current{0}, current{1} )),
                dict = List.Buffer(Table.FromValue({ 
                        {"В том числе: ", ""},
                        {"в том числе: ", ""}
                    }) [Value]),
                trans = Table.TransformColumns(Sort, {
                        {"ID_Строка", (x)=>x, Int64.Type}, 
                        {"Статья",    func, type text}
                    })
            in trans,

    // ЧАСТЬ 2 - СПРАВОЧНИКИ

        Replace = [
            с1 = Table.ReplaceValue(МультиЗамена, each [Статья], each if [ID_Строка] = 100 then "Остаток на начало 1" else [Статья],Replacer.ReplaceValue,{"Статья"}),
            с11 = Table.ReplaceValue(с1, each [Статья], each if [ID_Строка] = 110 then "Остаток на начало" else [Статья],Replacer.ReplaceValue,{"Статья"}),
            c2 = Table.ReplaceValue(с11, each [Статья], each if [ID_Строка] = 140 then "Остаток на начало скорректированный" else [Статья],Replacer.ReplaceValue,{"Статья"}),
            с3 = Table.ReplaceValue(c2, each [Статья], each if [ID_Строка] = 150 then "Увеличение собственного капиатала - всего" else [Статья],Replacer.ReplaceValue,{"Статья"}),
            c4 = Table.ReplaceValue(с3, each [Статья], each if [ID_Строка] = 151 or [Строка] = 161 then "чистая прибыль" else [Статья],Replacer.ReplaceValue,{"Статья"}),
            c5 = Table.ReplaceValue(c4, each [Статья], each if [ID_Строка] = 200 then "Остаток на конец " else [Статья],Replacer.ReplaceValue,{"Статья"})
        ][c5],

	// Разделы (свой вариант) отчета
    Спр_Разделы = 
		let	from = #table(type table 
			[	
				ID_Раздел 	  = number, 
				РазделПрефикс = text, 
				РазделИмя 	  = text, 
				СтрокаНач     = number, 
				СтрокаКон     = number
			],
			{    
				{1, "Раздел 1", "Позапрошлый год", 10, 40}, 
				{2, "Раздел 2", "Прошлый год", 50, 130},
                {3, "Раздел 3", "Текущий год", 140, 200}
			}),
			to = Table.AddColumn(from, "РазделПолноеИмя", each [РазделПрефикс] & " " & [РазделИмя])
		in	to,

    Select  = Table.SelectRows(Replace, each not List.Contains({10..100}, [ID_Строка] )),
    Remove  = Table.RemoveColumns(Select,{"Итого", "Index"}),
    Unpivot = Table.UnpivotOtherColumns(Remove, {"Date", "Статья", "ID_Строка"}, "Показатель", "Значение"),
    AddDataSet = Table.AddColumn(Unpivot , "DataSet", each 2,Int64.Type)
in
    AddDataSet