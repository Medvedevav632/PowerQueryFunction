let
    
//  СПРАВОЧНИКИ
    // Названия разделов статотчета "1п райпо"
    Спр_Разделы = 
        let Источник = #table(type table 
            [   ID_Раздел     = number, 
                РазделПрефикс = text, 
                РазделИмя     = text
            ],
            {   {1, "Раздел 1", "Раздел I. Производство промышленной продукции (услуг промышленного характера)"}, 
                {2, "Раздел 2", "Раздел II. Основные показатели работы промышленности, тыс. руб."}
            }),
            AddCol  = Table.AddColumn(Источник, "РазделИмяПрефикс", each [РазделПрефикс] & " " & [РазделИмя])
        in  AddCol,

    // справочник единиц измерения (для замены)
    // lst1 = List.Buffer({"Гкал", "м.п.", "млн. руб.", "млн. шт.", "млн.руб.", "т.дал", "тонн", "туб", "тыс. кв. м", "тыс. руб.", "тыс. штук","тыс.дал", "тыс.кВт.ч", "тыс.пар", "тыс.руб.", "тыс.штук"}),
    Спр_ЕдИзм =
        let lst1 = {"Гкал", "м.п.", "млн. руб.", "млн. шт.", "млн.руб.", "т.дал", "тонн", "туб", "тыс. кв. м", "тыс. руб.", "тыс. штук","тыс.дал", "тыс.кВт.ч", "тыс.пар", "тыс.руб.", "тыс.штук"},
            lst2 = {"Гкал", "м.п.", "млн. руб.", "млн. шт.", "млн.руб.", "т.дал", "тонн", "туб", "тыс. кв. м", "тыс. руб.", "тыс. штук","тыс.дал", "тыс.кВт.ч", "тыс.пар", "тыс.руб.", "тыс.штук"}, 
            from = #table( {"Что","На что"}, List.Zip({lst1, lst2}))
        in  from,
    
    // справочник единиц измерения (для замены)
    Спр_ВидТоваров =
        let Источник = #table(type table 
            [
                ID_КатегорияТоваров = number, 
                КатегорияТоваров = text
            ],
            {
                {1,"Продовольственные товары"}, 
                {2,"Непродовольственные товары"} 
            })
        in  Источник,

    // заголовки разделов
    Спр_Заголовки =
        let Источник = #table(type table 
            [   
                IndexIn = text,      // порядковый номер заголовка в пределах раздела отчета!
                ЗаголовокИмя = text, // имя заголовка 
                НомерРаздела = text  // номера разделов, в которых есть этот заголовок
            ],
            {   
                {"1", "Наименование вида продукции (услуги)","1"}, 
                {"2", "Единица измерения","1"},
                {"3", "Фактическое производство","1"},
                {"4", "Фактическое производство за предыдущий год","1"},
                {"5", "Темп роста, %","1"},
                {"6", "Стоимость продукции в фактических ценах","1"},
                {"7", "Стоимость продукции в фактических ценах за предыдущий год","1"}, 
                
                {"1", "Hаименование показателя","2"},
                {"2", "Отчетный год","2"},
                {"3", "Предыдущий год","2"},
                {"4", "Темп изменения, %","2"}
            }),
            
            // разворачиваем по строкам через запятую
            ExpandList = Table.ExpandListColumn(Table.TransformColumns(Источник, 
                            {{"НомерРаздела", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), 
                            let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}}), 
                            "НомерРаздела"),
            
            // объединяем два столбца в новый используя функцию НомерРаздела + Индекс
            CombineCol = Table.CombineColumns(ExpandList, 
                            {"НомерРаздела", "IndexIn"},
                            Combiner.CombineTextByDelimiter("", QuoteStyle.None),
                            "ID_Заголовок"),
            TransCol   = Table.TransformColumnTypes(CombineCol,{{"ID_Заголовок", Int64.Type}})
        in  TransCol,

//   ЗАГРУЗКА 
    
    from = Folder.Files("G:\=EXCEL BI - мои проекты\2023.05 - Отчетность (БКС)\Отчетность\Форма 1-п райпо")[[Name],[Content]],
    
    // вытягиваем дату из названия файла (4 знака с конца)
    Date = 
        let f=(x)=> Date.From(Text.Middle(x, 16, 4)),
            trs = Table.TransformColumns(from, {"Name", f, Date.Type}),
            ren = Table.RenameColumns(trs,{"Name", "Year"}),
            to  = Table.TransformColumns(ren, {"Content", (x)=>Excel.Workbook(x)} )
        in  to,
//  Колонак Период
    // ИзвлекаемТекст1 = Table.TransformColumns(УдаленныеСтолбцы, {{"Период", each Text.BeforeDelimiter(_, "."), type text}}),
    // ИзвлекаемТекст2 = Table.TransformColumns(ИзвлекаемТекст1, {{"Период", each Text.AfterDelimiter(_, " "), type text}}),
    // ИзмененныйТип1 = Table.TransformColumnTypes(ИзвлекаемТекст2,{{"Период", Int64.Type}}),

    // оставляем листы с номерами в названии, это коды организаций  / еси в новую колонку sel = Table.AddColumn(Date, "Content2", each f([Content]))
    Transform = 
        let f=(x)=> Table.SelectRows(x, each try Value.Is( Number.From([Name]), type number) = true otherwise null),
            sel = Table.TransformColumns(Date,{"Content", f})
        in  sel,

    Expand1 = Table.ExpandTableColumn(Transform, "Content", {"Name", "Data"}, {"ID_Предприятие", "Data"}),
    #"Строки с примененным фильтром" = Table.SelectRows(Expand1, each true) //  ([ID_Предприятие] = "1")
,
    
    Replace1 = Table.Buffer(
        Table.ReplaceValue(#"Строки с примененным фильтром", 
            each [ID_Предприятие], 
            each Record.FieldOrDefault([251600080200001 = 5079, 253100000000000 = 5001, 253500000000000 = 5005, 
                                        253900000000000 = 5009, 254100000000000 = 5011, 254900000000000 = 5020,
                                        227100000000000 = 5555], // 5555 - такого номера нет, временный
                    Text.From([ID_Предприятие]),
                    [ID_Предприятие]), 
            Replacer.ReplaceValue, {"ID_Предприятие"})
    ),

// ЧИСТКА И ТРАНСФОРМАЦИЯ

    // работаем в подтаблицах
    Clear = 
        let f=(x)=> 
                let 
                    // фильтруем колонки (оставляем countCol (7) колонок = максимальное количество колонок с данными
                    countCol = 7,
                    SelCol = Table.SelectColumns(x, List.Generate(()=> 1, (x)=> x<=countCol, (x)=>x+1, (x)=>"Column" & Text.From(x) )),
                    // RemCol = Table.RemoveColumns(SelCol, {"Column1"}),  // убираем первую колонку с нумерацией статей в этом отчете этого нет
                    
                    // удаляем (меняем на "") пробелы, переносы, табуляции др. (лучше до extraxt и unpivot)
                    // если в ячейке null, то ошибка поэтому try-otherwise
                    k=(x)=> try Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " ") otherwise x,
                    transform = Table.TransformColumns(SelCol,{"Column1", k}),

                    // делаем колонку 2 текстовую
                    TransCol   = Table.TransformColumnTypes(transform,{{"Column2", type text}}),
                    
                    lst = List.Buffer({"тонн","тыс", "млн", "кал", "дал", "туб", "м.п."}),
                    f=(tbl)=> try List.ContainsAny({tbl[Column2]}, lst, Text.Contains) otherwise null,
                    SelRow1 = Table.SelectRows(TransCol, f),
                    
                    t=(tbl)=> List.Contains (List.Buffer({"А".."Я","A".."Z"}), Text.Start(tbl[Column1],1) ),
                    SelRow2 = Table.SelectRows(SelRow1, t),

                    // меняем имена столбцов, по порядку, потом по номеру название учитывая разделы
                    lstOld = List.Buffer(Table.ColumnNames(SelRow2)),                       // старый список имен столбцов из подтаблиц (все одинаковые)
                    lstNew = List.Buffer({"Статья", "ЕдИзм", "1", "2", "3", "4", "5"}),  // новый список имен столбцов = 7
                    rename = Table.RenameColumns(SelRow2, List.Zip( {lstOld,lstNew} ))
                in  rename,
              tr = Table.TransformColumns(Replace1,{"Data", f})  
        in  tr,

// ЧИСТКА И ТРАНСФОРМАЦИЯ
    // работаем в подтаблицах
    
    // разворачиваем и переворачиваем таблицу
    Expand2 = Table.ExpandTableColumn(Clear, "Data", Table.ColumnNames(Clear{0}[Data]),Table.ColumnNames(Clear{0}[Data])),
    Unpivot = Table.UnpivotOtherColumns(Expand2, {"Year", "ID_Предприятие", "Статья", "ЕдИзм"}, "Заголовок", "Значение"),

    // меняем тип колонок с трансформацией
    TransCol = Table.TransformColumns(
            Unpivot,{ 
                { "Year",       (x)=> Date.From(x), Date.Type},
                { "ID_Предприятие", (x)=> Number.From(x), Int64.Type},
                { "Статья",     (x)=>Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " "), type text},
                { "ЕдИзм",      (x)=> x, type text},
                { "Заголовок",  (x)=> Number.From(x), Int64.Type},
                { "Значение",   (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type}
            } ),

// ОБОГАЩЕНИЕ

    // функция мультидобавления колонок
    fxAddMultiColumns = (table as table, addCol as list) as table => List.Accumulate(addCol, table, (t, d) => Table.AddColumn(t, d{0}, d{1}, d{2}?)),
    
    // добавляем колонки с цифровым ID атрибутов
        // { "СтатьяСтрока", each Text.From([ID_Строка]) & " " & Text.From([Статья]) , type text},         // номер строки + название статьи
        // { "ID_Группа",    each if Text.End(Text.From([ID_Строка]), 1) = "0" then 1 else 2, Int64.Type}  // если последняя цифра 0, то группа (1), иначе подгруппа (2)
    AddColumns = fxAddMultiColumns(TransCol, {
        // статьи с номерами строк от 100 до 399 соответствуют Части 1 (Активы), другие Части 2 (Пассивы)
        { "ID_Раздел",  each 1, Int64.Type},
    
        // уникальный номер заголовка в пределах отчета (только столбцы со значениями!): 1 знак - номер раздела, 2 знак - номер заголока в пределах раздела
        { "ID_Заголовок", each Number.From(Text.From([ID_Раздел]) & Text.From([Заголовок])), Int64.Type}
        }),

// ОБЪЕДИНЯЕМ

    // добавляем значения колонок с ID, если переключатель IsReportViewGroup = True
    Join = 
        let Join_Заголо  = Table.Join(AddColumns, "ID_Заголовок",   Спр_Заголовки,   "ID_Заголовок"),
            Join_Предпр  = Table.Join(Join_Заголо, "ID_Предприятие", Спр_Предприятия2, "Код")
            
        in  Join_Предпр,
    #"Строки с примененным фильтром1" = Table.SelectRows(Join, each ([Статья] <> "ИТОГО НЕПРОДОВОЛЬСТВЕННЫЕ ТОВАРЫ" and [Статья] <> "ИТОГО ПРОДОВОЛЬСТВЕННЫЕ ТОВАРЫ"))
in
    #"Строки с примененным фильтром1"