// ОПУ

let
	// Функция загрузки функций
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
	
    /* ШАГ 1 ***********************************************************************************************/
    
    
    // подключаемся к папке
    path  = Value.Metadata(PathFolder)[List]{1},
    load = 
        let from  = Folder.Files(path),
            filtr = Table.SelectRows(from, each [Extension] = ".xls" or [Extension] = ".xlsx" and [Attributes][Hidden] = false),
            tbl   = Table.SelectColumns(filtr,{"Name","Content"})
        in  tbl,

    // выбираем нужные листы в файле 
    tr = Table.TransformColumns(load,
        {"Content", (x)=>
            let a = Excel.Workbook(x),     // соджержимое файла, если лист в файле один, то (x)=>Excel.Workbook(x){1}[Data]}
                b = Table.SelectRows(a, each Text.Length([Name]) = 15), // выбираем нужный лист (код организации в названии листа = 15 знаков)
                c = b{0}[Data]
            in  c
        }),

    // добавили колонки с расчетами
    addColumns = 
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

    // создаем функцию выборки и чистки двух столбцов
    JoinExpand_1 =
        [   a = Table.NestedJoin(addColumns,"ID_Форма",Спр_ФормыОтчетности,"ID_Форма","tmp"),
            b = Table.ExpandTableColumn(a,"tmp", {"Название"})
        ][b],
    Filter = Table.SelectRows(JoinExpand_1, each ([ID_Форма] = "0212")) [[Date],[Content]],
    
    to = Table.Buffer(Filter),
    
/* ШАГ2 **********************************************************************************************/

/* ПОДКЛЮЧЕНИЕ --------------------------------------------------------------------------------------*/
   // Определяем список колонок будущей таблицы
    lstNewColNаmes = {"Date", "Статья", "ID_Строка", "Текущий период", "Прошлый период"},

// ШАГ 2 - Общая чистка
    f2=(x)=>
        let /*a1 = Table.Skip(x, 18),          // убираем первые строки
            a2 = Table.RemoveLastN(a1, 10),  // убираем последние строки*/
            a3 = Table.SelectColumns(x,{"Column1", "Column2", "Column3", "Column4"}), // оставляем колонки

            // Чистим текст во всей таблице (" " в начале для удаления двойных пробелов!)
            ClearData = fxClear(a3, " #(lf)#(00A0)", true, true),
            // k=(x)=> try Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " ") otherwise x,
            // ClearData = Table.TransformColumns(Select, {}, k),
            
            // Список имен колонок
            ColNames = Table.ColumnNames(ClearData),

            // Меняем тип данных колонки ID_Строка для дальшейшей фильтрации
            Transform = Table.TransformColumns(ClearData,{ 
                    { ColNames{1}, (x)=> try Number.From(x) otherwise null , Int64.Type}
                } ),
        
            // Фильтруем строки ПО ИМЕНИ колонки, не по [полю]
            // Table.SelectRows(RenameColumns, each {ColNames{1}} <> null), не работает!
            f=(x)=> x <> null and x > 9 and x < 1000 , // код строки в ОПУ начинается с 10
            Фильтр = Table.SelectRows (Transform, 
                    (row) => f (Record.Field (row, ColNames{1} ) ) 
                ),

            // Меняем названия колонок
            RenameColumns = Table.RenameColumns(Фильтр, List.Zip({
                    List.Buffer( ColNames ), 
                    List.Skip(lstNewColNаmes, 1)
                }) )
    in RenameColumns,

/* ДОБАВЛЕНИЕ СТОЛБЦОВ */

    AddColumnf2 = Table.AddColumn(to, "Step2", each f2([Content])),
    
    Источник   = AddColumnf2[[Date],[Step2]],
    
    EXPAND = Table.ExpandTableColumn(
        Источник, 
        "Step2", 
            List.Skip(lstNewColNаmes, 1) , 
            List.Skip(lstNewColNаmes, 1)
    ),
    
    // трансформируем столбцы с указанием типа
    TransformAllColumns = Table.TransformColumns(EXPAND,{ 
        {"Date",           (x)=> Date.From(x), Date.Type},
        {"Статья",         (x)=> x, Text.Type},
        {"ID_Строка",      each try Number.From(_) otherwise null, type number},
        {"Текущий период", each try Number.From(_) otherwise _, Number.Type},
        {"Прошлый период", each try Number.From(_) otherwise _, Number.Type}
    } ),
    
    // Меняем 0 на null. Чтобы в визуализации не появлялся.
    Replace0ToNull = Table.ReplaceValue(TransformAllColumns, 
        0, null, Replacer.ReplaceValue, 
        {"Текущий период", "Прошлый период"}
    ),

    // Убираем текст внутри скобок и сами скобки
    УбираемТекстВСкобках = Table.TransformColumns ( Replace0ToNull, 
        { "Статья", (x) =>                                                   // if Text.Contains (x, "(") = true
            if fnCompareGroups({"010","100","±"}, x) = true               // проверяем колонку x на наличие символ из списка
            then Text.BeforeDelimiter(x, "(", {0, RelativePosition.FromEnd}) // если есть, то оставляем текст ДО "(" => ПОСЛЕ первой скобки все убираем
            else x, type text } 
    ),

    // делаем отступ, если номер строки не четный
    ДобавляемОтступ = Table.ReplaceValue(УбираемТекстВСкобках, 
                each [Статья], 
                each if Number.Mod([ID_Строка], 10) <> 0  then "    " & [Статья] else [Статья], 
                Replacer.ReplaceValue,{"Статья"}),
    
    AddedColumns = fxAddMultiColumns(ДобавляемОтступ, {
        { "ID_Группа", each 
            if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1)) = true 
            then 1 
            else 2, Int64.Type},
        { "ID_Прибыль", each 
            if List.Contains ({30,60,90,140,150,201,210,240},[ID_Строка]) = true 
            then 1 
            else 2, Int64.Type},
        { "ID_ВидДеят", each 
            if List.Contains ({10..90},[ID_Строка]) = true then 1
            else if List.Contains ({100..140},[ID_Строка]) = true then 2
            else 3, Int64.Type}
        }),

    ColNames = Table.ColumnNames(AddedColumns),
    
    // перегруппировка колонок
    ReorderColumns = fnReorderSubsetOfColumns(AddedColumns, 
            { ColNames{0}, ColNames{5}, ColNames{7}, ColNames{6}, ColNames{2} }, 
            0),
    
    // добавляем колонку с номером года
    AddColumn_Year = Table.AddColumn(ReorderColumns, "Год", each Date.Year([Date]), Int64.Type),
    
    
    /* СПРАВОЧНИКИ */
    
    // справочник видов деятельности
    Спр_ОПУ_ВидДеятельности = 
        let Источник = #table(type table 
            [   
                ID_ВидДеятельности = Int64.Type, 
                ID_ДиапазонСтрок = text, 
                ВидДеятельностиИмя = text 
            ],
            {
                {1, "10..90",   "Текущая деятельность"}, 
                {2, "100..140", "Инвестиционная деятельность"},
                {3, "80..100",  "Другие показатели"}
            }
        )
        in Источник,

    // справочник видов показателей прибыли
    Спр_ОПУ_ВидПрибыли =
        let Источник = #table(type table [ ID_Прибыль = Int64.Type, ПрибыльВид = text ],
        {
            {1, "Показатели прибыли"}, 
            {2, "Другие показатели"}
        } )
    in Источник,

    // справочник групп статей
    Спр_ОПУ_Группа = 
        let Источник = #table(type table [ID_Группа = number, ГруппаИмя = text],{     
                {1,"Группа"}, 
                {2,"Подгруппа"}
            } )
        in Источник,
    
    // справочник статей
    Спр_ОПУ_Статья = 
        let
            Источник        = Table.Distinct(EXPAND[[ID_Строка],[Статья]], {"ID_Строка"}),
            AddCol_СтрокаID = Table.AddColumn(Источник, "СтатьяСтрока", each Text.From([ID_Строка]) & " " & [Статья]),
            RenameColumns   = Table.RenameColumns(AddCol_СтрокаID,{{"Статья", "СтатьяИмя"}}),
            
            // Делаем множественную замену в подстроке столбца
            МультиЗамена = [ 
                dict = List.Buffer(Table.FromValue({ 
                        {"продукции, товаров, работ, услуг", ""},
                        {"основных средств", "ОС"},
                        {"нематериальных активов", "НА"},
                        {"долгосрочных активов", "ДА"},
                        {"инвестиционной деятельности", "ИнвДеят"},
                        {"финансовой деятельности", "ФинДеят"},
                        {"отложенных налоговых активов", "ОНА"},
                        {"отложенных налоговых обязательств", "ОНО"}
                    }) [Value]),
                func = (x)=> List.Accumulate(dict, x, (state, current) => Text.Replace(state, current{0}, current{1} )),
                b1   = Table.TransformColumns(RenameColumns, {"СтатьяИмя", func}),
                b2   = Table.TransformColumns(b1, {"СтатьяСтрока", func})
                ][b2],
            AddIndexColumn = Table.AddIndexColumn(МультиЗамена, "СтрокаИндекс", 1, 1, Int64.Type)
        in  AddIndexColumn,
    
       
    /* JOIN */
    
    // обединяем таблицу фактов со справочниками с формированием дополнительных колонок-срезов
    Join = 
        let Join_ВидДеят    = Table.Join(AddColumn_Year,  "ID_ВидДеят", Спр_ОПУ_ВидДеятельности,"ID_ВидДеятельности"),
            Join_ВидПрибыли = Table.Join(Join_ВидДеят,    "ID_Прибыль", Спр_ОПУ_ВидПрибыли,"ID_Прибыль"),
            Join_Группа     = Table.Join(Join_ВидПрибыли, "ID_Группа",  Спр_ОПУ_Группа,"ID_Группа"),
            Join_Строка     = Table.Join(Join_Группа,     "ID_Строка",  Спр_ОПУ_Статья,"ID_Строка")
        in  Join_Строка,
    
    // двигаем колонку "Год" в начало, после даты
    Reorder = Table.ReorderColumns(Join,
        {"Date", "Год", "ID_Группа", "ID_ВидДеят", "ID_Прибыль", "ID_Строка", "Статья", "Текущий период", "Прошлый период"}
    ),
      
    AddDataSet = Table.AddColumn(Reorder, "DataSet", each 2,Int64.Type),    

    // выводим: true - узкую таблицу, fase - таблицу ""все в одном" с допколонками со значениями групп
    result = if IsReportViewGroup 
             then AddDataSet
             else AddColumn_Year
in
    result