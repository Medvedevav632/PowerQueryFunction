// Надо разделить на части. 
// Переменная часть. Загрузка - зависит от источника разный код может быть
// Постоянная часть. Независит от источника. Но на входе д.быть одинаковый перечень и название столбов + типы
// Шаг 0 - функция Load (если не используется, можно убрать. Пока только для fxClear нужен)
// Шаг 1 - Подключенее к источнику данных + создание допколонок (зависит от источника)
// Шаг 2 - Чистка в подтаблицах (почти универсальный шаг)
// Шаг 3 - Обогащение в основной таблице (универсальный шаг)
// ДОБАВИТЬ: 1. Обработку ошибок при добавлении в папку с отчетами других файлов
//  		 2. 
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
    path = "G:\=EXCEL BI - мои проекты\2023.01 - Анализ финансовой отчетности (Гом ОПС)\= Отчетность",
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
    
    // добавляем колонки через объединение таблиц по ID_Форма (нет необходимости, убрать надо будет)
    addColumns2 =
        let Спр_ФормыОтчетности = #table(type table 
				[ID_Форма = text, Название = text, КраткоеНазвание = text, Примечание = text],{
					{"0112","Баланс","ББ",""}, 
					{"0212","Отчет и прибылях и убытках", "ОПУ",""},
					{"0312","Отчет о доходах, расходах и финансовых результатах","ОДРФР",""},
					{"0412","Отчет о движении денежных средств","ОДДС",""},
					{"0512","Приложение к бухгалтерскому балансу","ПББ",""},
                    {"0612","Отчет о состоянии расчетов","ОСР",""},
                    {"06R1","Отчет о состоянии расчетов","ОСР",""},
                    {"12f1","Отчет о финансовых результатах","ОФР_12ф",""},
                    {"1212","Отчет о финансовых результатах ежемесячная","ОФР_мес",""},
					{"3012","Отчет об изменении капитала","ОИК",""} }),
			Join   = Table.NestedJoin(addColumns1,"ID_Форма",Спр_ФормыОтчетности,"ID_Форма","tmp"),
			Expand = Table.ExpandTableColumn(Join,"tmp", {"Название"})
		in  Expand
		
    filterReport = Table.SelectRows(addColumns2, each ([ID_Форма] = "0112")) [[Date],[Content]],
    

/* ШАГ 2 - ЧИСТКА ДАННЫХ в подтаблицах -------------------------------------------------------*/
    from = filterReport,
   
    // список колонок итоговой таблицы
    lstNewColNаmes = {"Date", "Статья", "ID_Строка", "На начало", "На конец"},

	/* ШАГ 2.1 - Очистка подтаблицы */
    
	// общая чистка в подтаблице
    f=(x)=>
        let 
			// фильтруем колонки
			Select = Table.SelectColumns(x, {"Column1", "Column2", "Column3", "Column4"} ), 

			// чистим текст во всей таблице (" " двойные пробелы, табуляцию,)
			ClearData = Load("fxClear")(Select, " #(lf)#(00A0)", true, true),
			
			// список имен колонок подтаблицы
			ColNames = Table.ColumnNames(ClearData),

			// меняем тип данных второй колонки - "НомерСтроки" для дальнейшей фильтрации
			Transform1 = Table.TransformColumns(ClearData,{ 
					{ ColNames{1}, (x)=> try Number.From(x) otherwise null, Int64.Type}
				} ),
		
			// фильтруем строки в колонке ПО ИМЕНИ колонки, не по [полю]
			// Table.SelectRows(RenameColumns, each {ColNames{1}} <> null), не работает!
			f=(x)=> (x <> null) and (x > 100) and (x < 1000),
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
            { lstNewColNаmes{3}, (x)=> if Number.From(x) = 0 then null else Number.From(x), Int64.Type},
            { lstNewColNаmes{4}, (x)=> if Number.From(x) = 0 then null else Number.From(x), Int64.Type}
		} 
	),

	
/* ШАГ 4 - ОБОГАЩЕНИЕ ----------------------------------------------------------------------------------------*/

    // функция мультидобавления колонок
    fxAddMultiColumns = (tbl as table, addCol as list) as table =>
        List.Accumulate(addCol, tbl, (t, d) => Table.AddColumn(t, d{0}, d{1}, d{2}?)),
    
    // добавляем колонки с цифровыми ID атрибутами
    AddedColumns1 = fxAddMultiColumns(Transform, {
        // статьи с номерами строк от 100 до 399 соответствуют Части 1 (Активы), другие Части 2 (Пассивы)
        { "ID_Часть",  each 
                if   List.Contains ({110..399}, [ID_Строка]) = true 
                then 1 
                else 2, Int64.Type},
        
        // номер раздела соответствует первому символу номера строки. 0 - не вошедшие в разделы
        { "ID_Раздел", each 
                let fxRazdel = (x as text) => Record.FieldOrDefault([1 = 1, 2 = 2, 4 = 3, 5 = 4, 6 = 5, 0 = 0], x, null),
                    Номер = Text.Start(Text.From([ID_Строка]), 1)    
                in  fxRazdel(Номер), Int64.Type},
        
        // если первая буква большая, то группа (1), если маленькая, то подгруппа (2)
        { "ID_Группа", each 
				if   List.Contains ({"А".."Я", "A".."Z"}, Text.Start([Статья], 1)) = true 
				then 1 
				else 2, Int64.Type},
        
        // уровень иерархии: 1 - Часть, 2 - Раздел, 3 - Статья (равно 0 в конце № строки), 4 - Подстрока (не равно 0 в конце № строки)
        { "ID_Уровень", each 
				if      List.Contains({300, 700}, [ID_Строка]) = true then 1 
                else if List.Contains({190, 290, 390, 490, 590, 690}, [ID_Строка]) = true then 2 
                else if Number.Mod([ID_Строка], 10) <> 0  then 4
                else 3, Int64.Type}
        }),

    
/* АНАЛИЗ ЛИКВИДНОСТИ ---------------------------------------------------------------------------------------*/

    // добавляем колонки с ID атрибутами для анализа ликвидности
    AddColumnLikvid = fxAddMultiColumns(AddedColumns1, {
        // статус строк для включения в таблицу ликвидности
        { "КатегорияЛиквOnOf", each 
            if   List.Contains ({190,290,390,490,590,690,300,700, 630, 211,212,213,214,215,216},[ID_Строка]) = true 
            then 0 
            else 1, Int64.Type},
        
        // те что включаются, делим по категориям 4 А + 4 П
        {"ID_КатегорияЛикв",each 
			if      List.Contains ({260, 270}, [ID_Строка]) = true then "А1"
            else if List.Contains ({210, 250, 280}, [ID_Строка]) = true then "А2"
            else if List.Contains ({220, 230, 240, 150, 170}, [ID_Строка]) = true then "А3"
            else if List.Contains ({110, 120, 130, 140, 160, 180}, [ID_Строка]) = true then "А4"
            else if List.Contains ({632, 633, 634, 635, 636, 637, 638, 639}, [ID_Строка]) = true then "П1" 
            else if List.Contains ({610, 620, 640, 650, 660, 670, 631}, [ID_Строка]) = true then "П2" 
            else if List.Contains ({510, 520, 530, 540, 550, 560}, [ID_Строка]) = true then "П3"
            else if List.Contains ({410, 420, 430, 440, 450, 460, 470, 480}, [ID_Строка]) = true then "П4" 
            else "0", type text}
        }),
   
    // делаем отступы слева в названиии подстатей
    Отступ = Table.ReplaceValue(AddColumnLikvid, 
        each [Статья], 
        each if Number.Mod([ID_Строка], 10) <> 0 then "    " & [Статья] else [Статья], 
        Replacer.ReplaceValue, {"Статья"} ),
    
   
/* СПРАВОЧНИКИ --------------------------------------------------------------------*/

    // справочник частей баланса
    Спр_ББ1_Часть = #table(type table 
            [
                ID_Часть = number, 
                ЧастьИмяПолное = text, 
                ЧастьИмяКраткое = text, 
                ЧастьАббревиутура = text
            ],
            {
                {1,"Активы","Актив","А"}, 
                {2,"Собственный капитал и обязательства", "Пассив","П"}
            } ),
    
    // справочник разделов баланса
    Спр_ББ2_Раздел = 
        let Источник = #table(type table 
            [
                ID_Раздел = number, 
                РазделПрефикс= text, 
                РазделИмя = text, 
                РазделИмяПолное = text,
                РазделИмяСокращенное = text, 
                РазделИмяАббревиатура = text
            ],
            {
                {1, "Раздел 1", "Долгосрочные активы", "Раздел 1 Долгосрочные активы", "ДолгАктивы", "ДА"}, 
                {2, "Раздел 2", "Краткосрочные активы", "Раздел 2 Краткосрочные активы", "КраткАктивы", "КА"},
                {3, "Раздел 3", "Собственный капитал", "Раздел 3 Собственный капитал", "СобКапитал", "СК"},
                {4, "Раздел 4", "Долгосрочные обязательства", "Раздел 4 Долгосрочные обязательства", "ДолгОбяз", "ДО"},
                {5, "Раздел 5", "Краткосрочные обязательства", "Раздел 5 Краткосрочные обязательства", "КраткОбяз", "КО"},
                {0, "Раздел 6", "Забалансовые счета", "Раздел 0 Забалансовые счета", "ЗабСчета", "ЗС"}
            } ),
        SelectRows = Table.SelectRows(Источник, each ([ID_Раздел] <> 0))
    in  SelectRows,

    // справочник групп статей баланса
    Спр_ББ3_Группа = #table(type table [ID_Группа = number, ГруппаИмя = text],
		{
			{1,"Группа"}, 
            {2,"Подгруппа"}
        } ),

    // справочник уровней баланса
    Спр_ББ_Уровень = 
        let Источник = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WMlTSUbqw+GLDhX0XNl3YemHvxR4FsNDyCxsuNl5sutijFKsTrWSEqQostODChgvbL2wBCu0GqzPGVAcWWnixCWge0LSL/WB1JpjqwELzgQJbQPYiVMcCAA==", BinaryEncoding.Base64), Compression.Deflate)), 
            let _t = ((type nullable text) meta [Serialized.Text = true]) 
            in type table [ID_Уровень = _t, УровеньПрефикс = _t, УровеньИмя = _t]),
            tr = Table.TransformColumnTypes(Источник,{
					{"ID_Уровень", Int64.Type}, 
					{"УровеньИмя", type text}
				})
        in  tr,

    // справочник статей баланса (статьи берутся из источника)
    Спр_ББ4_Статья = 
        let Источник        = EXPAND [[ID_Строка],[Статья]],
            lstColName      = {"Статья", "СтатьяИмя", "СтатьяСтрока"},   // Имена колонок: старое имя + новое имя, имя доп колонки
            TblDistinct     = Table.Distinct(Источник, {"ID_Строка"}),
            RenameColumns   = Table.RenameColumns(TblDistinct,{{"Статья", "СтатьяИмя"}} ),
            AddColumn       = Table.AddColumn(RenameColumns, "СтатьяСтрока", each Text.From([ID_Строка]) & " " & [СтатьяИмя]),
            SelectRows      = Table.SelectRows( AddColumn, each true),  // Не знаю, зачем так

            // множественная замена в столбце - часть ячейки
            МультиЗамена = [ 
                dict = List.Buffer(Table.FromValue({ 
                        {"Налог на добавленную стоимость", "НДС"},
                        {"НДС по приобретенным товарам, работам, услугам", "НДС по приобретенным Т,Р,У"},
                        {"Долгосрочные активы, предназначенные для реализации", "Долгосрочные активы для реализации"},
                        {"Долгосрочные обязательства по лизинговым платежам", "Долгосрочные обязательства по лизингу"},
                        {"Денежные средства и эквиваленты денежных средств", "Денежные средства и их эквиваленты"}, 
                        {"кредиторская задолженность", "КЗ"},                
                        {"Вложения в долгосрочные активы", "Вложения в ДА"},
                        {"дебиторская задолженность", "ДЗ"} 
                    }) [Value]),
                func = (x)=> List.Accumulate(dict, x, (state, current) => Text.Replace(state, current{0}, current{1} )),
                transform = Table.TransformColumns(SelectRows, {   
                            {"ID_Строка", (x)=>x, Int64.Type}, 
                            {"СтатьяИмя",    func, type text}, 
                            {"СтатьяСтрока", func, type text} 
                        })
                ][transform]
            
            /*  Пользовательский1 = Table.InsertRows(МультиЗамена, 61,
                    {[ID_Строка = 800, СтатьяИмя = "Прочие", СтрокаСтатья = "800 Прочие"]} ) */
		in МультиЗамена,

/* РАЗВОРАЧИВАЕМ ------------------------------------------------------------------------------------*/

    // добавляем значения колонок с ID, если переключатель IsReportViewGroup = True
    Join = [ Join_Часть   = Table.Join(Отступ,      "ID_Часть",  Спр_ББ1_Часть,  "ID_Часть"),
             Join_Раздел  = Table.Join(Join_Часть , "ID_Раздел", Спр_ББ2_Раздел, "ID_Раздел"),
             Join_Группа  = Table.Join(Join_Раздел, "ID_Группа", Спр_ББ3_Группа, "ID_Группа"),
             Join_Строка  = Table.Join(Join_Группа, "ID_Строка", Спр_ББ4_Статья, "ID_Строка"),
             Join_Уровень = Table.Join(Join_Строка, "ID_Уровень",Спр_ББ_Уровень, "ID_Уровень")
        ][Join_Уровень],

    
/* РЕЗУЛЬТАТ ---------------------------------------------------------------------------------------*/ 

    // выводим итоговую таблицу 
    IsReportViewGroup = true,
    End_Step2 = if IsReportViewGroup 
                    then Join       // узкую таблицу без доп колонок, только с кодами
                    else Отступ,    // широкую таблицу с допколонками + значениями групп ID "все в одном"
    
    // ставим метки на шагах запроса
    // Value.Metadata(А_Баланс_Загрузка)[Expand_]
    Result = End_Step2 meta [Expand_ = EXPAND]
in
    Result