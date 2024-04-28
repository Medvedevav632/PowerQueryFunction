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
    path  = "G:\=EXCEL BI - мои проекты\2023.01 - Анализ финансовой отчетности (Гом ОПС)\= Отчетность",
    load = 
        let from  = Folder.Files(path),
            filtr = Table.SelectRows(from, each [Extension] = ".xls" or [Extension] = ".xlsx" and [Attributes][Hidden] = false),
            tbl   = Table.SelectColumns(filtr,{"Name","Content"})
        in  tbl,

    // выбираем нужные листы в файле 
    tr = Table.TransformColumns(load,
        {"Content", (x)=>
            let a = Excel.Workbook(x),                                  // соджержимое файла, если лист в файле один, то (x)=>Excel.Workbook(x){1}[Data]}
                b = Table.SelectRows(a, each Text.Length([Name]) = 15), // выбираем нужный лист (нам нужен с кодом организации в названии листа = 15 знаков)
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
    Filter = Table.SelectRows(JoinExpand_1, each ([ID_Форма] = "0412")) [[Date],[Content]],
    
    to = Table.Buffer(Filter),
	
// определяем список колонок будущей таблицы
    lstNewColNаmes = {"Date", "Статья", "ID_Строка", "Текущий период", "Прошлый период"},

    // добавляем колонку с исходником нужного отчета 4 - ОДДС
    AddColumn1 = Table.AddColumn(to, "Step1", each [Content] ),


/* ЧИСТКА */

    // общая чистка. Ориентируемся в основном на номер статьи
    f2=(x)=>
        let 
            a3 = Table.SelectColumns(x,{"Column1", "Column2", "Column3", "Column4"}), // оставляем колонки
            // Чистим текст во всей таблице (" " в начале для удаления двойных пробелов!)
            ClearData = fxClear(a3, " #(lf)#(00A0)", true, true),
            
            // Список имен колонок
            ColNames = Table.ColumnNames(ClearData),

            // Меняем тип данных колонки ID_Строка для дальшейшей фильтрации
            Transform = Table.TransformColumns(ClearData,{ 
                    { ColNames{1}, (x)=> try Number.From(x) otherwise null , Int64.Type}
                } ),
        
            // Фильтруем строки ПО ИМЕНИ колонки
            f=(x)=> x <> null and x > 9 and x < 1000, // код строки в ОПУ начинается с 10
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

    AddColumnf2 = Table.AddColumn(AddColumn1, "Step2", each f2([Step1])),
    
    Источник   = AddColumnf2[[Date],[Step2]],
    
    // разворачиваем подтаблицы
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
    
    // меняем 0 на null (чтобы в визуализации не появлялся)
    Replace0ToNull = Table.ReplaceValue(TransformAllColumns, 
        0, null, Replacer.ReplaceValue, 
        {"Текущий период", "Прошлый период"}
    ),

  
  /* ОБОГАЩЕНИЕ */
  
  // добавляем колонки
  AddedColumns = fxAddMultiColumns(Replace0ToNull, {
        { "СтатьяСтрока", each Text.From([ID_Строка]) & " " & [Статья] },
        
        { "ID_Группа",  each 
                if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1)) = true 
                then 1 
                else 2 , Int64.Type}, // 1- Группа, 2 - Подгруппа
        
        { "ID_ВидДеят", each if List.Contains ({20..40},[ID_Строка]) = true then 1 // "Текущая"
                        else if List.Contains ({50..70},[ID_Строка]) = true then 2 // "Инвестиционная"
                        else if List.Contains ({80..100},[ID_Строка]) = true then 3 // "Финансовая"
                        else 4, Int64.Type}, // "Другое"
        
        { "ID_ВидДвиж", each if List.Contains ({20..24, 50..55, 80..84},[ID_Строка]) = true then "Поступление" 
                        else if List.Contains ({30..34, 60..64, 90..95},[ID_Строка]) = true then "Выбытие" 
                        else if List.Contains ({40},[ID_Строка]) = true then "РезультатТек" 
                        else if List.Contains ({70},[ID_Строка]) = true then "РезультатИнв" 
                        else if List.Contains ({100},[ID_Строка]) = true then "РезультатФин" 
                        else if List.Contains ({110},[ID_Строка]) = true then "РезультатОбщий" 
                        else "Другое", type text}
        }),

    ReorderColumns = Table.ReorderColumns(
        AddedColumns,
        {"Date", "ID_Группа", "ID_ВидДеят", "ID_ВидДвиж", "ID_Строка", "Статья", "СтатьяСтрока", 
         "Текущий период", "Прошлый период"}
    ),
    
    Год = Table.AddColumn(ReorderColumns, "Год", each Date.Year([Date]), Int64.Type),
   
    ReorderColumns2 = Table.ReorderColumns(Год,
        {"Date", "Год", "ID_Группа", "ID_ВидДеят", "ID_ВидДвиж", "ID_Строка", "Статья", "СтатьяСтрока", 
        "Текущий период", "Прошлый период"}),
    
    Отступ = Table.ReplaceValue(
        ReorderColumns2, 
        each [Статья], 
        each if Number.Mod([ID_Строка], 10) <> 0  
             then "    " & [Статья] 
             else [Статья], 
        Replacer.ReplaceValue,{"Статья"}
        ),


/* СПРАВОЧНИКИ */  

    // справочник
    Спр_ОДДС_Виды_Деятельности = 
        let
            Источник = #table(type table [ВидДеятельностиДиапазонСтрок = text, ВидДеятельности = text],{
                    {"20..40","Текущая деятельность"}, 
                    {"50..70","Инвестиционная деятельность"},
                    {"80..100","Финансовая деятельность"}, 
                    {"110..140"," Другое"}
                }
        ),
            #"Добавлен индекс" = Table.AddIndexColumn(Источник, "ID_Индекс", 1, 1, Int64.Type),
            #"Переупорядоченные столбцы" = Table.ReorderColumns(#"Добавлен индекс",{"ID_Индекс", "ВидДеятельности", "ВидДеятельностиДиапазонСтрок"}),
            #"Переименованные столбцы" = Table.RenameColumns(#"Переупорядоченные столбцы",{{"ID_Индекс", "ID_ВидДеятельности"}})
        in    #"Переименованные столбцы",

    // справочник
    Спр_ОДДС_ВидыДвижения = 
        let
        Источник = #table(type table [ВидДвиженияДиапазонСтрок = text, ID_ВидДвижения = text],{
                {"20..24","Поступило"},
                {"50..55","Поступило"}, 
                {"80..84","Поступило"}, 
                {"30..34","Выбыло"},
                {"60..64","Выбыло"},
                {"90..95","Выбыло"},
                {"40","РезультатТек"}, 
                {"70","РезультатИнв"}, 
                {"100","РезультатФин"}, 
                {"110","РезультатОбщий"}, 
                {"120..140","Другие"}
            }
        ),
            #"Добавлен индекс" = Table.AddIndexColumn(Источник, "ID_Индекс", 1, 1, Int64.Type),
            #"Переупорядоченные столбцы" = Table.ReorderColumns(#"Добавлен индекс",{"ID_Индекс", "ID_ВидДвижения", "ВидДвиженияДиапазонСтрок"})
        in #"Переупорядоченные столбцы",

    // справочник
    Спр_ОДДС_Статья = 
        let
            Источник      = Table.Distinct(Replace0ToNull[[ID_Строка],[Статья]],{"ID_Строка"}),
            
            // добавляем отступ к строкам подгруппы
            ReplaceValue  = Table.ReplaceValue(
                        Источник, 
                        each [Статья], 
                        each if Number.Mod([ID_Строка], 10) <> 0  
                            then "    " & [Статья] 
                            else [Статья], 
                        Replacer.ReplaceValue,{"Статья"}
                    ),
            
            ДобавляетмОтступ = Table.AddColumn(ReplaceValue, "СтатьяСтрока", each Text.From([ID_Строка]) & " " & [Статья]),

            // Делаем множественную замену в подстроке столбца
                МультиЗамена = [ 
                    dict = List.Buffer(Table.FromValue({ 
                            {"от покупателей продукции, товаров, заказчиков работ, услуг", "покупателей товаров"},
                            {"от покупателей материалов и других запасов", "покупателей других запасов"},
                            {"доходы от участия в уставном капитале других организаций", "доходы от дивидендов"},
                            {"основных средств, нематериальных активов и других долгосрочных активов", "ОС, НА и других ДА"},
                            {"вклады собственника имущества (учредителей, участников)", "вклады собственника"},
                            {"на выплаты дивидендов и других доходов от участия в уставном капитале организации", "на выплаты дивидендов"},
                            {"текущей деятельности", "ТекДеят"},
                            {"инвестиционной деятельности", "ИнвДеят"},
                            {"финансовой деятельности", "ФинДеят"}
                        }) [Value]),
                    func = (x)=> List.Accumulate(dict, x, (state, current) => Text.Replace(state, current{0}, current{1} )),
                    b1   = Table.TransformColumns(ДобавляетмОтступ, {"Статья", func}),
                    b2   = Table.TransformColumns(b1, {"СтатьяСтрока", func})
                    ][b2]
            in  МультиЗамена,

    // выводим: true - узкую таблицу (шаг AddColumn_Year), fase - таблицу ""все в одном" с допколонками (шаг Отступ)
    IsReportViewGroup = true,
    result = if IsReportViewGroup then Отступ else Replace0ToNull
in
    result
	
	
	
// ДОПОЛНЕНИЕ

/* 
Убираем текст внутри скобок и сами скобки
    УбираемТекстВСкобках = Table.TransformColumns ( Replace0ToNull, 
        { "Статья", (x) =>                                                   // if Text.Contains (x, "(") = true
            if fnCompareGroups({"010","100","±"}, x) = true               // проверяем колонку x на наличие символ из списка
            then Text.BeforeDelimiter(x, "(", {0, RelativePosition.FromEnd}) // если есть, то оставляем текст ДО "(" => ПОСЛЕ первой скобки все убираем
            else x, type text } 
    ),
*/