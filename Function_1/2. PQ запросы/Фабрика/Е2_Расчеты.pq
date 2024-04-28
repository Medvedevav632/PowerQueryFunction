// Статотчет Расчеты

let

/* ШАГ 0 - Модуль загрузки функций Load ---------------------------------------------------------------------*/

	Load = (fnName as text, optional BasePath as text) as function =>
		let DefaultPath = "G:\=EXCEL PQ\PQ-function-load\",
			GitHubPath  = "https://raw.githubusercontent.com/hohlick/pquery/master/",
				
			BasePath = if (BasePath <> null) then BasePath else DefaultPath,
			Path = BasePath & (if Text.End(BasePath, 1) <> "\" then "\" else ""),
			File = Path & fnName & ".pq",
			
			Function = try Expression.Evaluate(Text.Replace(fnName, ".", "_"), #shared)
			otherwise try  Expression.Evaluate(Text.FromBinary(Binary.Buffer(File.Contents(File))), #shared)
			otherwise      Expression.Evaluate(Text.FromBinary(Binary.Buffer(Web.Contents(GitHubPath & fnName & ".pq"))), #shared)
		in	Function,
	
    
/* ШАГ 1 - ПОДКЛЮЧЕНИЕ к источнику данных -------------------------------------------------------------------*/
    
    // подключаемся к папке
    path = Value.Metadata(PathFolder)[List]{1},
    load = 
		let from = Folder.Files(path),
            fltr = Table.SelectRows(from, each [Extension] = ".xls" or [Extension] = ".xlsx" and [Attributes][Hidden] = false),
            tbl  = Table.SelectColumns(fltr,{"Name","Content"})
        in  tbl,

    // выбираем нужные листы в файле 
    selectSheet = Table.TransformColumns(load,
        {"Content", (x)=>
            let a = Excel.Workbook(x), // соджержимое файла, если лист в файле один, то (x)=>Excel.Workbook(x){1}[Data]}
                b = Table.SelectRows(a, each Text.Length([Name]) = 15), // лист с кодом организации в названии листа = 15 знаков
                c = b{0}[Data]
            in  c
        }),

    // добавили колонки с допинформацией о файле/отчете
    addColumns1 = 
        let f = (tbl as table, addcol as list) as table => List.Accumulate(addcol, tbl, (t, d) => Table.AddColumn(t, d{0}, d{1}, d{2}?)),
            // т.к в файле xls дата = 01.16, в xlsx = 01.12.2016, то возникает ошибка => обрабатываем ее через try/otherwise
            AddedColumns = f(selectSheet, {
                { "Date",           each try       Date.From("01.01.20" & Text.End([Content][Column1]{0},2))             // если файл xls
                                         otherwise Date.From([Content][Column1]{0} + #duration(30, 0, 0, 0)), type date}, // если файл xlsx, то добаваляем 30 дней
                { "ID_Предприятие", each Text.From([Content][Column1]{1}), type text },
                { "ID_Форма",       each Text.Start([Content][Column2]{0},4), type text}
            })
        in  AddedColumns,
    #"Удаленные ошибки" = Table.RemoveRowsWithErrors(addColumns1, {"Date"}),
    
    // добавляем колонки через объединение таблиц по ID_Форма (нет необходимости, убрать надо будет)
    addColumns2 =
        let Спр_ФормыОтчетности = #table(type table 
				[ID_Форма = text, Название = text, КраткоеНазвание = text, Примечание = text],{
					{"0112","Баланс","ББ",""}, 
					{"0212","Отчет и прибылях и убытках", "ОПУ",""},
					{"0312","Отчет о доходах, расходах и финансовых результатах","ОДРФР",""},
					{"0412","Отчет о движении денежных средств","ОДДС",""},
					{"0512","Приложение к бухгалтерскому балансу","ПББ",""},
					{"3012","Отчет об изменении капитала","ОИК",""} }),
			Join   = Table.NestedJoin(#"Удаленные ошибки","ID_Форма",Спр_ФормыОтчетности,"ID_Форма","tmp"),
			Expand = Table.ExpandTableColumn(Join,"tmp", {"Название"})
		in  Expand,
    filterReport = Table.SelectRows(addColumns2, each [ID_Форма] = "0612") [[Date],[Content]],
    

/* ШАГ 2 - ЧИСТКА ДАННЫХ в подтаблицах -------------------------------------------------------*/
    from = filterReport,
   
    // список колонок итоговой таблицы
    lstNewColNаmes = {"Date", "Статья", "ID_Строка", "Всего", "Просроченная"},

	/* ШАГ 2.1 - Очистка подтаблицы */
    
	// общая чистка в подтаблице
    f=(x)=>
        let 
			// фильтруем колонки
			Select = Table.SelectColumns(x, {"Column1", "Column2", "Column3", "Column4"} ), 

			// чистим текст во всей таблице (" " двойные пробелы, табуляцию,)
			// ClearData = Load("fxClear")(Select, " #(lf)#(00A0)", true, true),
			k=(x)=> try Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " ") otherwise x,
            ClearData = Table.TransformColumns(Select, {}, k),

			// список имен колонок подтаблицы
			ColNames = Table.ColumnNames(ClearData),

			// меняем тип данных второй колонки - "НомерСтроки" для дальнейшей фильтрации
			Transform1 = Table.TransformColumns(ClearData,{ 
					{ ColNames{1}, (x)=> try Number.From(x) otherwise null, Int64.Type},
                    { ColNames{3}, (x)=> try Number.From(x) otherwise null, Int64.Type}
				} ),
		
			// фильтруем строки в колонке ПО ИМЕНИ колонки, не по [полю]
			// Table.SelectRows(RenameColumns, each {ColNames{1}} <> null), не работает!
			f=(x)=> (x <> null) and (x > 0) and (x < 1000),
			SelectRow = Table.SelectRows (Transform1, 
					(row) => f (Record.Field (row, ColNames{1} ) ) 
				),

			// меняем названия колонок кроме первой (первая Date)
			RenameColumns = Table.RenameColumns(SelectRow, List.Zip({
					List.Buffer(ColNames), 
					List.Skip(lstNewColNаmes, 1)
				}) )
    
		in RenameColumns,

    /* ШАГ 2.2 - Добавление очищенной подтаблицы */
	
	// добавляем подтаблицу
    addColumn3 = Table.AddColumn(from, "Content2", each f([Content])),

    // конец шага 2.1 
    End_Step1 = addColumn3 [[Date],[Content2]],


/* ШАГ 3 - ТРАНСФОРМАЦИЯ --------------------------------------------------------------------------------------*/

    // Разворачиваем подтаблицы. Список столбцов без первой колонки "Date", которой внутри нет
    // надежней будет колонки определять сразу до разворота  ColNames = Table.ColumnNames(from2[Step2]{0})
    EXPAND = Table.ExpandTableColumn(End_Step1, 
        "Content2", 
		List.Skip(lstNewColNаmes, 1), 
        List.Skip(lstNewColNаmes, 1)
    ),
    
    // меняем тип колонок с трансформацией. Можно делать это и в подтаблицах, до разворота
    Transform = Table.TransformColumns(EXPAND,
		{ 
			{ lstNewColNаmes{0}, (x)=> Date.From(x), Date.Type},
            { lstNewColNаmes{1}, (x)=> x, type text},
            { lstNewColNаmes{2}, (x)=> try Number.From(x) otherwise null, Int64.Type},
            { lstNewColNаmes{3}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
            { lstNewColNаmes{4}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type}
		} 
	),

	
// ставим метки на шагах запроса
    // Value.Metadata(А_Баланс_Загрузка)[Expand_]
    Result = Transform,
    #"Добавлен пользовательский объект" = Table.AddColumn(Result, "СтатьяСтрока", each Text.From([ID_Строка]) & " " & [Статья]),
    #"Строки с примененным фильтром" = Table.SelectRows(#"Добавлен пользовательский объект", each [ID_Строка] < 200) // AddDataSet meta [Expand_ = EXPAND]
in
    #"Строки с примененным фильтром"