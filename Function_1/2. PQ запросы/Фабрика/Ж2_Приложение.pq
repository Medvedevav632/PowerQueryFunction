// ПРИЛОЖЕНИЕ к балансу

let

/* ШАГ 0 - Модуль загрузки функций Load ---------------------------------------------------------------------*/

//  СПРАВОЧНИКИ
    // Названия разделов отчета Приложение в Балансу
    Спр_Разделы = 
        let Источник = #table(type table 
            [   ID_Раздел     = number, 
                РазделПрефикс = text, 
                РазделИмя     = text, 
                Str1          = number,
                Str2          = number],
            {   {1, "Раздел 1", "Раздел I. Основные средства и нематериальные активы", 1, 36}, 
                // {2, "Раздел 1", "Справочно. ", 1, 5}, 
                // {3, "Раздел 1", "Источники финансирования сложений ", 30, 36}, 
                {2, "Раздел 2", "Раздел II. Незавершенные вложения во внеоборотные активы", 50, 65},
                {3, "Раздел 3", "Раздел III. Резервы по сомнительным долгам", 70, 70},
                {4, "Раздел 4", "Раздел IV. Финансовые вложения", 80, 95},
                {5, "Раздел 5", "Раздел V. Полученные кредиты и займы", 140, 171},
                {6, "Раздел 6", "Раздел VI. Справка о платежах в бюджет и ФСЗН", 1010, 1320},
                {7, "Раздел 7", "Раздел VII. Забалансовые счета", 1401, 1411}
            }),

            AddCol  = Table.AddColumn(Источник, "РазделИмяПрефикс", each [РазделПрефикс] & " " & [РазделИмя]),
            Reorder = Table.ReorderColumns(AddCol,
					let nms = Table.ColumnNames(AddCol), lst = {"РазделИмяПрефикс"}, col = 3  // 0 - если в начало
					in  List.InsertRange(List.Difference(nms, lst ), col, lst) ),
            SelCol = Table.SelectColumns(Reorder, List.FirstN(Table.ColumnNames(Reorder), 4) )
    in  SelCol,

    // заголовки разделов
    Спр_Заголовки =
        let Источник = #table(type table 
            [   
                IndexIn = text,      // порядковый номер в заголовка пределах раздела отчета
                ЗаголовокИмя = text, // имя заголовка 
                НомерРаздела = text  // номера разделов, в которых есть этот заголовок
            ],
            {   
                {"1", "На начало года","1"}, 
                {"2", "Поступило","1"},
                {"3", "Выбыло","1"},
                {"4", "На конец отчетного периода всего","1"},
                {"5", "На конец отчетного периода всего, и.т.ч. полностью самортизированных","1"},
                
                {"1", "На начало года","2"},
                {"2", "Поступило","2"},
                {"3", "Выбыло","2"},
                {"4", "На конец периода","2"}, 
                
                {"1", "На начало года","3"},
                {"2", "Создано","3"},
                {"3", "Списано","3"},
                {"4", "На конец периода","3"}, 

                {"1", "На начало года","4,7"},
                {"2", "На конец периода","4,7"}, 

                {"1", "На начало года","5"},
                {"2", "Получено","5"},
                {"3", "Погашено","5"},
                {"4", "На конец периода","5"}, 

                {"1", "Пpичитается по расчету","6"},
                {"2", "Изpасходовано","6"},
                {"3", "Фактически уплачено","6"}
            }),
           
            // разворачиваем по строкам через запятую
            ExpandList = Table.ExpandListColumn(Table.TransformColumns(Источник, 
                            {{"НомерРаздела", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), 
                            let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}}), 
                            "НомерРаздела"),
            
            // объединяем два столбца в новый используя функцию
            CombineCol = Table.CombineColumns(ExpandList, 
                            {"НомерРаздела", "IndexIn"},
                            Combiner.CombineTextByDelimiter("", QuoteStyle.None),
                            "ID_Заголовок"),
            TransCol   = Table.TransformColumnTypes(CombineCol,{{"ID_Заголовок", Int64.Type}})
        in  TransCol,

   
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
    RemoveRowsWithErrors = Table.RemoveRowsWithErrors(addColumns1, {"Date"}),
    
    // добавляем колонки через объединение таблиц по ID_Форма (нет необходимости, убрать надо будет)
    addColumns2 =
		let	Join   = Table.NestedJoin(RemoveRowsWithErrors,"ID_Форма",Спр_ФормыОтчетности,"ID_Форма","tmp"),
			Expand = Table.ExpandTableColumn(Join,"tmp", {"Название"})
		in  Expand,
    filterReport = Table.SelectRows(addColumns2, each [ID_Форма] = "0512") [[Date],[Content]],
    

/* ШАГ 2 - ЧИСТКА ДАННЫХ в подтаблицах -------------------------------------------------------*/
    from = filterReport,
   
    // список колонок итоговой таблицы
    lstNewColNаmes = {"Date", "Статья", "ID_Строка", "1", "2", "3", "4", "5"},

	/* ШАГ 2.1 - Очистка подтаблицы */
    
	// общая чистка в подтаблице
    f=(x)=>
        let 
			// фильтруем колонки
			Select = Table.SelectColumns(x, {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7"} ), 

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
			f=(x)=> (x <> null) and (x > 0) and (x < 2000),
            
			SelectRow = Table.SelectRows (Transform1, 
					(row) => f (Record.Field (row, ColNames{1} ) ) 
				),

			// меняем названия колонок кроме первой (первая Date)
			RenameColumns = Table.RenameColumns(SelectRow, List.Zip({
					List.Buffer(ColNames), 
					List.Skip(lstNewColNаmes,1)
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
            { lstNewColNаmes{2}, (x)=> x, Int64.Type},
            { lstNewColNаmes{3}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
            { lstNewColNаmes{4}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
            { lstNewColNаmes{5}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
            { lstNewColNаmes{6}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
            { lstNewColNаmes{7}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type}
		} 
	),
    Unpivot = Table.UnpivotOtherColumns(Transform, {"Date", "Статья", "ID_Строка"}, "Заголовок", "Значение"),

  
    // добавляем колонки
    AddedColumns = fxAddMultiColumns(Unpivot, 
        {
            { "ID_Раздел",  each if List.Contains ({1..36},      [ID_Строка]) = true then 1
                            else if List.Contains ({50..65},     [ID_Строка]) = true then 2
                            else if List.Contains ({70..70},     [ID_Строка]) = true then 3
                            else if List.Contains ({80..95},     [ID_Строка]) = true then 4
                            else if List.Contains ({140..171},   [ID_Строка]) = true then 5
                            else if List.Contains ({1010..1320}, [ID_Строка]) = true then 6
                            else if List.Contains ({1401..1411}, [ID_Строка]) = true then 7
                            else 9, Int64.Type},
            
            { "СтатьяСтрока", each Text.From([ID_Строка]) & " " & [Статья], type text},
            
            { "ID_Заголовок", each Number.From(Text.From([ID_Раздел]) & Text.From([Заголовок])), Int64.Type}
        }),

	// добавляем значения колонок с ID, если переключатель IsReportViewGroup = True
    Join = 
        let Join_Заголо  = Table.Join(AddedColumns, "ID_Заголовок", Спр_Заголовки,   "ID_Заголовок"),
            Join_Раздел  = Table.Join(Join_Заголо,  "ID_Раздел",    Спр_Разделы,     "ID_Раздел")
        in  Join_Заголо,
            
// ставим метки на шагах запроса
    // Value.Metadata(А_Баланс_Загрузка)[Expand_]
    Result = Join
in
    Result