let

/* ПОДКЛЮЧЕНИЕ --------------------------------------------------------------------------------------*/
    from = ВсеФормыВпапке1,
   
/* ЧИСТКА ПОДТАБЛИЦ ---------------------------------------------------------------------------------*/

    // список начальных колонок итоговой таблицы: общие + обязательные (форму еще не добавили)
    lst1 = {"DataSet","Организация","Date"},
    lst2 = {"Статья", "Строка_ID", "На начало", "На конец"},

    // добавляем подтаблицы 2. [Data] - колонка в исходной подтаблице. 1 = БАЛАНС
    AddColTable2Report = Table.AddColumn(from, "Step1", each [Content]{1}[Data]),

    // добавляем подтаблицы 3 (можно через функцию: each f([Step1])),
    AddColTable3Clear = Table.AddColumn(AddColTable2Report, "Step2", (x)=>
        let 
            // оставляем нужные колонки в подтаблице (статья, номер статьи, на начало, на конец)
            Select    = Table.SelectColumns(x[Step1],{"Column1", "Column5", "Column6", "Column7"}),   
            
            // чистим текст во всей таблице (" ", двойные пробелы, табуляцию,)
            ClearData = fxPack[fxClear](Select, " #(lf)#(00A0)", true, true),                       
            
            // вторая колонка с из списка {"Статья", "Строка_ID", "На начало", "На конец"} = "Строка_ID"
            ColStrokeID  = Table.ColumnNames(ClearData){1},

            // меняем тип данных "Строка_ID" для дальнейшей фильтрации строк
            Transform = Table.TransformColumns(ClearData,{{ ColStrokeID, (x)=> try Number.From(x) otherwise null, Int64.Type}} ),
        
            // фильтруем строки с НомеромСтатьи ПО ИМЕНИ колонки, не по [полю], т.к. each {ColNames{1}} <> null, не работает!
            f=(x)=> x<>null and x>100 and x<1000,
            Filter = Table.SelectRows (Transform, (row) => f(Record.Field(row,ColStrokeID)) ),

            // меняем названия колонок в подтаблице кроме первых на нужные - lst2
            RenameColumns = fxPack[fxRenameColumns](Filter, Table.ColumnNames(Filter), lst2)
    
        in RenameColumns
    ),
    
    // конец первого этапа обработки подтаблиц
    _Result_Step1 = AddColTable3Clear [[DataSet],[Организация],[Date],[Step2]],


/* ТРАНСФОРМАЦИЯ ---------------------------------------------------------------------------------------*/

    // Разворачиваем подтаблицы. Список столбцов без первой колонки "Date", которой внутри нет
    // надежней будет колонки определять сразу до разворота  ColNames = Table.ColumnNames(from2[Step2]{0})
    EXPAND = Table.ExpandTableColumn(_Result_Step1, "Step2", lst2, lst2),
    
    // меняем тип колонок с трансформацией. Можно делать это и в подтаблицах, до разворота
    // Три допколонки (Датасет + Организация + Дата) + 4 обязательные (Статья, Строка, На начало, На конец)
    Transform = Table.TransformColumns(
            EXPAND,{ 
                { lst1{0}, (x)=> Number.From(x), Int64.Type},   // датасет
                { lst1{1}, (x)=> x, type text},                 // организация
                { lst1{2}, (x)=> Date.From(x), Date.Type},      // дата
                
                { lst2{0}, (x)=> x, type text}, // статья
                { lst2{1}, (x)=> try Number.From(x) otherwise null, Int64.Type}, // строка
                { lst2{2}, (x)=> if Number.From(x) = 0 then null else Number.From(x), Int64.Type}, // на начало
                { lst2{3}, (x)=> if Number.From(x) = 0 then null else Number.From(x), Int64.Type}  // на конец
            } ),

    
/* ОБОГАЩЕНИЕ -----------------------------------------------------------------------------------------*/

    // Обязательно должны быть 4 стандартные колонки: Статья, Строка, Сумма на начало, Сумма на конец
    AddColumns0 = fxPackRep[fxAddColumnOffset] (Transform, "balance", 0),   // добавляем колонку с номером отчета в начало таблицы
    AddColumns1 = fxPackRep[fxAddColBalanceID] (AddColumns0),               // добавляем колонки с цифровыма ID атрибутами (Часть, Раздел, Группа, Статья, Уровень)
    AddColumns2 = fxPackRep[fxAddColBalanceLiq] (AddColumns1),              // добавляем колонки для анализа ликвидности (Статус, Категория, Группа)


/* СПРАВОЧНИКИ ---------------------------------------------------------------------------------------*/
 
    // справочник статей баланса
    Спр_ББ_Статья = 
        let
            Источник        = EXPAND [[Строка_ID],[Статья]],
            lstColName      = {"Статья", "СтатьяИмя", "СтатьяСтрока"},   // Имена колонок: старое имя + новое имя, имя доп колонки
            TblDistinct     = Table.Distinct(Источник, {"Строка_ID"}),
            RenameColumns   = Table.RenameColumns(TblDistinct,{{"Статья", "СтатьяИмя"}} ),
            AddColumn       = Table.AddColumn(RenameColumns, "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [СтатьяИмя]),
            SelectRows      = Table.SelectRows(AddColumn, each true),  // Не знаю, зачем так

            // Делаем множественную замену в подстроке столбца
            МультиЗамена = [ 
                dict = List.Buffer(Table.FromValue({ 
                        {"Налог на добавленную стоимость",                      "НДС"},
                        {"НДС по приобретенным товарам, работам, услугам",      "НДС по приобретенным Т,Р,У"},
                        {"Долгосрочные активы, предназначенные для реализации", "Долгосрочные активы для реализации"},
                        {"Долгосрочные обязательства по лизинговым платежам",   "Долгосрочные обязательства по лизингу"},
                        {"Денежные средства и эквиваленты денежных средств",    "Денежные средства и их эквиваленты"}, 
                        {"кредиторская задолженность",                          "КЗ"},                
                        {"Вложения в долгосрочные активы",                      "Вложения в ДА"},
                        {"дебиторская задолженность",                           "ДЗ"} 
                    }) [Value]),
                func = (x)=> List.Accumulate(dict, x, (state, current) => Text.Replace(state, current{0}, current{1} )),
                transform = Table.TransformColumns(SelectRows, {   
                            {"Строка_ID", (x)=>x, Int64.Type}, 
                            {"СтатьяИмя",    func, type text}, 
                            {"СтатьяСтрока", func, type text} 
                        })
                ][transform]
            
            /*  Пользовательский1 = Table.InsertRows(МультиЗамена, 61,
                    {[ID_Строка = 800, СтатьяИмя = "Прочие", СтрокаСтатья = "800 Прочие"]} ) */
    in МультиЗамена,


/* ДОБАВЛЯЕМ КОЛОНКИ ИЗ СПРАВОЧНИКОВ  -----------------------------------------------------------------*/

    // добавляем значения колонок с ID, если переключатель IsReportViewGroup = True
    Join = [ Join_Часть   = Table.Join(AddColumns2,     "Часть_ID",     Справочники[БАЛАНС][Спр_ББ_Часть],   "Часть_ID"),
             Join_Раздел  = Table.Join(Join_Часть ,     "Раздел_ID",    Справочники[БАЛАНС][Спр_ББ_Раздел],  "Раздел_ID"),
             Join_Группа  = Table.Join(Join_Раздел,     "Группа_ID",    Справочники[БАЛАНС][Спр_ББ_Группа],  "Группа_ID"),
             Join_Уровень = Table.Join(Join_Группа,     "Уровень_ID",   Справочники[БАЛАНС][Спр_ББ_Уровень], "Уровень_ID"),
             Join_Строка  = Table.Join(Join_Уровень,    "Строка_ID",    Спр_ББ_Статья,                       "Строка_ID")
        ][Join_Строка],
    
/* НАВОДИМ КРАСОТУ ---------------------------------------------------------------------------------*/ 

    // добавляет отступ в колонке Статья
        OffsetChar = fxPackRep[fxOffsetColumn](Join),

    // перегруппировка колонок согласно кастомного списка в Excel
        ReorderColumns = fxPackRep[fxReorderAllColumns]( 
            OffsetChar, 
            Value.Metadata(СправочникСтолбцов)[balance][Имя] 
        ),

/* РЕЗУЛЬТАТ ---------------------------------------------------------------------------------------*/ 

    // выводим или узкую таблицу или с допколонками со значениями групп ID "все в одном"
    _Result_Step2 = if IsReportViewGroup 
                    then ReorderColumns 
                    else AddColumns2,
    
    // ставим метки на шагах запроса (потом использовать to = Value.Metadata(А_Баланс_Загрузка)[Expand_]_
    Result = _Result_Step2 meta [Expand_ = EXPAND]
in
    Result