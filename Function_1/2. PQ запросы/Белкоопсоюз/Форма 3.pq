let
    
//  СПРАВОЧНИКИ
    // Названия разделов отчета 3 БКС
    Спр_Разделы = 
        let Источник = #table(type table 
            [   ID_Раздел     = number, 
                РазделПрефикс = text, 
                РазделИмя     = text, 
                Str1          = number,
                Str2          = number],
            {   {1, "Раздел 1", "ИЗДЕРЖКИ ОБРАЩЕНИЯ (РАСХОДЫ) НА РЕАЛИЗОВАННЫЕ ТОВАРЫ И ТОВАРЫ, ПЕРЕДАННЫЕ ФИЛИАЛАМИ", 10100, 10700}, 
                {2, "Раздел 2", "ГРУППИРОВКА ИЗДЕРЖЕК ОБРАЩЕНИЯ ПО ЭЛЕМЕНТАМ ЗАТРАТ (строка 11500)", 21000, 25000},
                {3, "Раздел 3", "ДОХОДЫ, РАСХОДЫ И ФИНАНСОВЫЕ РЕЗУЛЬТАТЫ ОТ РЕАЛИЗАЦИИ ТОВАРОВ И ПЕРЕДАЧИ ТОВАРОВ ФИЛИАЛАМИ", 30000, 38005},
                {4, "Раздел 4", "ОСНОВНЫЕ СТАТИСТИЧЕСКИЕ ПОКАЗАТЕЛИ ТОРГОВО-ЗАКУПОЧНОЙ ДЕЯТЕЛЬНОСТИ", 40000, 47000},
                {5, "Раздел 5", "ДАННЫЕ О ПРОДАЖЕ ТОВАРОВ НАСЕЛЕНИЮ В КРЕДИТ В РОЗНИЧНОЙ ТОРГОВЛЕ", 50000, 53100},
                {6, "Раздел 6", "ЗАПАСЫ ТОВАРОВ И ТАРЫ НА КОНЕЦ ОТЧЕТНОГО ПЕРИОДА", 60000, 61000} }),
            AddCol  = Table.AddColumn(Источник, "РазделИмяПрефикс", each [РазделПрефикс] & " " & [РазделИмя]),
            Reorder = Table.ReorderColumns(AddCol,
					let nms = Table.ColumnNames(AddCol), lst = {"РазделИмяПрефикс"}, col =3  // 0 - если в начало
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
                {"1", "Розничная торговля","1,2,3,6"}, 
                {"2", "Оптовая торговля","1,2,3,6"},
                {"3", "Всего по торговле","1,2,3,6"},
                {"4", "Общественное питание","1,2,3,6"},
                {"5", "Заготовки","1,2,3,6"},
                {"6", "Итого","1,2,3,6"},
                {"1", "Отчетный период","4"}, 
                {"2", "Прошлый период","4"},
                {"1", "Сумма","5"}
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

//   ЗАГРУЗКА 
    
    from = Folder.Files("G:\=EXCEL BI - мои проекты\2023.05 - Отчетность БКС\Исходники\форма 3")[[Name],[Content]],
    
    // вытягиваем дату из названия файла (4 знака с конца)
    Date = 
        let f=(x)=> Date.From(Text.Middle(x, 7, 4)),
            trs = Table.TransformColumns(from, {"Name", f, Date.Type}),
            ren = Table.RenameColumns(trs,{"Name", "Year"}),
            to  = Table.TransformColumns(ren, {"Content", (x)=>Excel.Workbook(x)} )
        in  to,

    // оставляем листы с номерами в названии, это коды организаций  / еси в новую колонку sel = Table.AddColumn(Date, "Content2", each f([Content]))
    Transform = 
        let f=(x)=> Table.SelectRows(x, each try Value.Is( Number.From([Name]), type number) = true otherwise null),
            sel = Table.TransformColumns(Date,{"Content", f})
        in  sel,

    Expand1 = Table.ExpandTableColumn(Transform, "Content", {"Name", "Data"}, {"ID_Предприятие", "Data"}),
    
// ЧИСТКА И ТРАНСФОРМАЦИЯ

    // работаем в подтаблицах
    Clear = 
        let f=(x)=> 
                let 
                    // фильтруем колонки (оставляем 9 колонок (макс кол-во с данными для этого отчета), т.к. в некоторых отчетах есть больше колонок и вылазит ошибка
                    // фильтруем строки (название листа = число)
                    SelCol = Table.SelectColumns(x, List.Generate(()=> 1, (x)=> x<=9, (x)=>x+1, (x)=>"Column" & Text.From(x) )),        // оставляем 9 колонок
                    RemCol = Table.RemoveColumns(SelCol, {"Column1"}),  // убираем первую колонку с нумерацией статей 
                    SelRow = Table.SelectRows(RemCol, each try Value.Is( Number.From([Column3]), type number) = true otherwise null),   // оставляем строки с числов в названии
                    
                    // меняем имена столбцов, по порядку, потом по номеру название учитывая разделы
                    lstOld = List.Buffer(Table.ColumnNames(SelRow)),                // старый список имен столбцов из подтаблиц (все одинаковые)
                    lstNew = {"Статья", "ID_Строка", "1", "2", "3", "4", "5", "6"}, // новый список имен столбцов
                    rename = Table.RenameColumns(SelRow, List.Zip( {lstOld,lstNew} )), 
                    
                    // чистим (меняем знаки из списка на "") от пробелов, переносов строки и др. Лучше сейчас, до разворота, быстрее будет
                    // если в ячейке null, то ошибка поэтому try-otherwise
                    k=(x)=> try Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " ") otherwise x,
                    transform = Table.TransformColumns(rename,{"Статья", k})  
                in  transform,
            
            tr = Table.TransformColumns(Expand1,{"Data", f})  
        in  tr,
    
    // разворачиваем и переворачиваем таблицу
    Expand2 = Table.ExpandTableColumn(Clear, "Data", Table.ColumnNames(Clear{0}[Data]),Table.ColumnNames(Clear{0}[Data])),
    Unpivot = Table.UnpivotOtherColumns(Expand2, {"Year", "ID_Предприятие", "Статья", "ID_Строка"}, "Заголовок", "Значение"),

    // меняем тип колонок с трансформацией
    // Статья - чистим текст от пробелов внутри и снаружи и непечатных символов
    TransCol = Table.TransformColumns(
            Unpivot,{ 
                { "Year",       (x)=> Date.From(x), Date.Type},
                { "ID_Предприятие", (x)=> Number.From(x), Int64.Type},
                { "Статья",     (x)=>Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " "), type text},
                { "ID_Строка",  (x)=> try Number.From(x) otherwise null, Int64.Type},
                { "Заголовок",  (x)=> Number.From(x), Int64.Type}, // if Number.From(x) = 0 then null else Number.From(x), Int64.Type
                { "Значение",   (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type}
            } ),

// ОБОГАЩЕНИЕ

    // функция мультидобавления колонок
    fxAddMultiColumns = (table as table, addCol as list) as table => List.Accumulate(addCol, table, (t, d) => Table.AddColumn(t, d{0}, d{1}, d{2}?)),
    
    // добавляем колонки с цифровым ID атрибутов
    AddColumns = fxAddMultiColumns(TransCol, {
        // статьи с номерами строк от 100 до 399 соответствуют Части 1 (Активы), другие Части 2 (Пассивы)
        { "ID_Раздел",  each Number.From(Text.Start(Text.From([ID_Строка]), 1)), Int64.Type},
        
        // номер строки + название статьи
        { "СтатьяСтрока",  each Text.From([ID_Строка]) & " " & Text.From([Статья]) , type text},
        
        // уникальный номер заголовка в пределах отчета (только столбцы со значениями!): 1 знак - номер раздела, 2 знак - номер заголока в пределах раздела
        { "ID_Заголовок", each Number.From(Text.From([ID_Раздел]) & Text.From([Заголовок])), Int64.Type},
              
        // если первая буква большая, то группа (1), если маленькая, то подгруппа (2)
        { "ID_Группа", each 
                if      Text.End(Text.From([ID_Строка]), 3) = "000" then 1 
                else if Text.End(Text.From([ID_Строка]), 2) = "00" then 2
                else 3, Int64.Type}
        /*
        // уровень иерархии: 1 - Часть, 2 - Раздел, 3 - Статья (равно 0 в конце № строки), 4 - Подстрока (не равно 0 в конце № строки)
        { "ID_Уровень", each if List.Contains ({300, 700},[ID_Строка]) = true then 1 
                        else if List.Contains ({190, 290, 390, 490, 590, 690},[ID_Строка]) = true then 2 
                        else if Number.Mod([ID_Строка], 10) <> 0  then 4
                        else 3, Int64.Type}
        */
        }),

/* РАЗВОРАЧИАЕМ */

    // добавляем значения колонок с ID, если переключатель IsReportViewGroup = True
    Join = 
        let Join_Заголо  = Table.Join(AddColumns, "ID_Заголовок",   Спр_Заголовки,   "ID_Заголовок"),
            Join_Предпр  = Table.Join(Join_Заголо, "ID_Предприятие", Спр_Предприятия, "ID_Предприятие"),
            Join_Раздел  = Table.Join(Join_Предпр,  "ID_Раздел",     Спр_Разделы,     "ID_Раздел")
        in  Join_Раздел
in
    Join