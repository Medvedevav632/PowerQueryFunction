************************************
// СКРИПТ ЗАГРУЗКИ ИСХОДНЫХ ДАННЫХ

// ПАРАМЕТРЫ - "строка" записать как отдельный запрос
// ПутьОтчеты
"G:\=EXCEL BI - мои проекты\2023 - Анализ финансовой отчетности\= Отчетность" meta [IsParameterQuery=true, List={"G:\=EXCEL BI - мои проекты\2023 - Анализ финансовой отчетности\= Отчетность"}, DefaultValue=..., Type="Any", IsParameterQueryRequired=true]

// IsReportViewGroup
false meta [IsParameterQuery=true, List={true, false}, DefaultValue=true, Type="Logical", IsParameterQueryRequired=true]


// BaseScript
let
    // Список всех файлов в папке с подпапками
	from  = Folder.Files(ПутьОтчеты),
    
	// Оставляем только файлы Excel
	filtr = Table.SelectRows(from, each [Extension] = ".xls" 
									 or [Extension] = ".xlsx" 
								    and [Attributes][Hidden] = false),
    
	// Оставляем нужные колонки
	tbl   = Table.SelectColumns(filtr,{"Name","Folder Path","Content"}),

	// Вариант обработки файлов с исходными данными
    iVarLoad = 3,
    
    // Функция загрузки исходных файлов из  папки
    f=(x)=>
        // Вариант 1 - 
        if iVarLoad = 1 then 
            [ a = Excel.Workbook(x),
              c = a{0}[Data] // первый лист
            ][c]
        
        // Вариант 2 - 
        else if iVarLoad = 2 then 
            [ a = Excel.Workbook(x),
              c = a{1}[Data] // второй лист
            ][c]
        
        // Вариант 3: Если в файле 1 лист tr = Table.TransformColumns(tbl, {"Content", (x) => Excel.Workbook(x) {1} [Data] })
        else
            Excel.Workbook(x){1}[Data],
			
	f2=(x)=>
        [ a1 = Excel.Workbook(x) {1} [Data],
          a2 = a1
        ][a2],	
	
  
    // Применяем функцию трансформации
	tr = Table.TransformColumns(tbl,{"Content",f}),

    
	
	fnAddMultiColumns = (table as table, addedColumns as list) as table =>
        List.Accumulate(
            addedColumns,
            table,
            (t, d) => Table.AddColumn(t, d{0}, d{1}, d{2}?)),
    
	// Добавили колонки с расчетами (можно много сразу, функцию менять не надо)
	// Особенность. В файле xls дата как 01.16, в xlsx как 01.12.2016. Из за этого ошибка. 
	// Чтбы это обойти обрабатываем ошибку через try если верно и otherwise если ошибка
    AddedColumns = fnAddMultiColumns(tr, {
        { "Date",           each try       Date.From("01.01.20" & Text.End([Content][Column1]{0},2)) // если файл xls
                                 otherwise Date.From([Content][Column1]{0}+ #duration(30, 0, 0, 0)), type date}, // если файл xlsx + добаваляем 30 дней
        { "ID_Предприятие", each Text.From([Content][Column1]{1}), type text },
        { "ID_Форма",       each Text.Start([Content][Column2]{0},4), type text}
    }),

	// Создаем функцию выборки и чистки двух столбцов:
	// a - 
	// b - 
    JoinExpand_1 =
        [ a = Table.NestedJoin(AddedColumns, "ID_Форма", Спр_ФормыОтчетности, "ID_Форма", "tmp"),
          b = Table.ExpandTableColumn(a, "tmp", {"Название"} )
        ][b],
    

//--------------------------------------------------
// ШАГ2 - Чистим БАЛАНС (Источник = ВсеФормыВпапке,)

  // Фильтруем список отчетов по коду формы отчетности (0112 - это баланс)
  ФильтрФормы = Table.SelectRows(Источник, each ([ID_Форма] = "0112")) [[Date],[Content]],
    
  // Разворачиваем таблицу
  ExpandData = Table.ExpandTableColumn(
    ФильтрФормы, 
    "Content", 
    {"Column1", "Column2", "Column3", "Column4"}, 
    {"Column1", "Column2", "Column3", "Column4"}
  ),

  // Меняем тип данных
  ColTypes = Table.TransformColumnTypes(
    ExpandData, 
	{
		{"Column2", Int64.Type}, 
		{"Column3", Int64.Type}, 
		{"Column4", Int64.Type}
	}
  ),

  // Убирам ошибки (иногда могут появляться)
  RemoveErrors = Table.RemoveRowsWithErrors(ColTypes , {"Column2"}),
 
  // Фильтруем таблицу
  Фильтр = [
    a = Table.SelectRows(RemoveErrors, each ([Column2] <> null and [Column2] >= 99 and [Column2] <= 1000)), 
    b = Table.SelectRows(a, each not List.ContainsAny({[Column1]}, {"Справочно"}, Text.Contains))
  ][b], 

  //Замена "в том числе:" на ""
  Замена = [
    a = Table.ReplaceValue(Фильтр, "в том числе:", "", Replacer.ReplaceText, {"Column1"}), 
    b = Table.TransformColumns(a, {{"Column1", Text.Trim, type text}})
  ][b], 

  //Переименовываем колонки
  RenameColumns = Table.RenameColumns(
    Замена, 
    {
      {"Column1", "Статья"}, 
      {"Column2", "Строка"}, 
      {"Column3", "Конец периода"}, 
      {"Column4", "Начало периода"}
    }
  ),

  //Замена "Денежные средства и эквиваленты денежных средств" (в прошлых годах есть отличие в названии)
  ReplaceСтатья = Table.ReplaceValue(
    RenameColumns, 
    each [Статья], 
    each Record.FieldOrDefault(
      [Денежные средства и их эквиваленты = "Денежные средства и эквиваленты денежных средств"], 
      Text.From([Статья]), // через , можно продолжть список замен [..., Уставный капитал = 5001, ...]                                                                           
      [Статья]
    ),  
    Replacer.ReplaceValue, 
    {"Статья"}
  ), 

    // Удаляем нули
	ReplaceNull = Table.ReplaceValue(ReplaceСтатья,0,null,Replacer.ReplaceValue,{"Конец периода", "Начало периода"}),
    
    // Функция добавления нескольких колонок
	AddMultiColumns = (table as table, addedColumns as list) as table =>
        List.Accumulate(addedColumns, table, (t, d) => Table.AddColumn(t, d{0}, d{1}, d{2}?)),

    // Добавляеим колонки
	AddedColumns1 = AddMultiColumns(ReplaceNull, {
        { "ID_Часть",  each if List.Contains ({110..399}, [Строка]) = true then 1 else 2, Int64.Type},
        { "ID_Раздел", each fnRazdel(Text.Start(Text.From([Строка]),1)), Int64.Type},
        { "ID_Группа", each if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1)) = true then 1 else 2, Int64.Type}
        //{ "ID_Уровень", each if List.ContainsAny( {[Статья]}, {"ИТОГО","БАЛАНС"}, Text.Contains) then 1 else 2, Int64.Type}
        }),

    // уровень иерархии в балансе: 1 - баланс, 2 - итоги разделов, 3 - строки с 0 в конце, 4 - подстроки (статьи в группе запасы и др.)
    AddedColumns2 = AddMultiColumns(AddedColumns1, {
        { "ID_Уровень", each if List.Contains ({300, 700},[Строка]) = true then 1 
                        else if List.Contains ({190, 290, 390, 490, 590,690},[Строка]) = true then 2 
                        else if Number.Mod([Строка], 10) <> 0  then 4
                        else 3, Int64.Type}
        }),

    // Добаляем колонку "Год"
	AddColumn_Year = Table.AddColumn(AddedColumns2, "Год", each Date.Year([Date]), Int64.Type),
    
	// Делаем список названий колонок
	ColNames = Table.ColumnNames(AddColumn_Year),
    
    // Меням порядок столбцов
	ReorderColumns = fnReorderSubsetOfColumns(
        AddColumn_Year, 
        {ColNames{0}, ColNames{9}, ColNames{5}, ColNames{6}, ColNames{7}, ColNames{8}, ColNames{2}}, 
        0
    ),
    
    //Переименовываем колонки
	RenameColumns2 = Table.RenameColumns(
        ReorderColumns,{
            {"Конец периода", "На конец"}, 
            {"Начало периода", "На начало"}, 
            {"Строка", "ID_Строка"}}
    ),
    
    //Делаем отступ в строках подгруппы
	Отступ = Table.ReplaceValue(RenameColumns2, 
                each [Статья], 
                each if Number.Mod([ID_Строка], 10) <> 0  then "    " & [Статья] else [Статья], 
                Replacer.ReplaceValue,{"Статья"}),
    
    // Добавляем значения колонок с ID, если переключатель IsReportViewGroup = True
    Join = [ Join_Часть = Table.Join(Отступ,"ID_Часть",Спр_ББ1_Часть,"ID_Часть"),
            Join_Раздел = Table.Join(Join_Часть ,"ID_Раздел",Спр_ББ2_Раздел,"ID_Раздел"),
            Join_Группа = Table.Join(Join_Раздел,"ID_Группа",Спр_ББ3_Группа,"ID_Группа"),
            Join_Строка = Table.Join(Join_Группа,"ID_Строка",Спр_ББ4_Статья,"ID_Строка")
        ][Join_Строка],

    
    //Выводим: узкую таблицу (если Параметр IsReportViewGroup = True) или с допколонками со значениями групп ID "все в одном"
    result = if IsReportViewGroup then Join else Отступ
in 
   result
   
   
   
// СПРАВОЧНИКИ
   
// Спр_ББ1_Часть
let
    Источник = #table(type table [ID_Часть = number, ЧастьПрефикс= text, ЧастьИмя = text],{     
            {1,"Актив","Активы"}, 
            {2,"Пассив", "Собственный капитал и обязательства"}
          }
)
in
    Источник

// Спр_ББ2_Раздел
let
    Источник = #table(type table [ID_Раздел = number, РазделПрефикс= text, РазделИмя = text, ИмяСокр = text],{     
            {1,"Раздел 1","Долгосрочные активы","ДолгАктивы"}, 
            {2,"Раздел 2", "Краткосрочные активы", "КраткАктивы"},
            {3,"Раздел 3","Собственный капитал","СобКапитал"},
            {4,"Раздел 4","Долгосрочные обязательства","ДолгОбяз"},
            {5,"Раздел 5","Краткосрочные обязательства","КраткОбяз"},
            {0,"Раздел 6","Забалансовые счета","ЗабСчета"}
          }
),
    #"Добавлен пользовательский объект" = Table.AddColumn(Источник, "РазделПолноеИмя", each [РазделПрефикс] & " " & [РазделИмя])
in
    #"Добавлен пользовательский объект"

// Спр_ББ3_Группа
let
    Источник = #table(type table [ID_Группа = number, ГруппаИмя = text],{     
            {1,"Группа"}, 
            {2,"Подгруппа"}
          }
)
in
    Источник

// Спр_ББ_Заголовки
let
    Источник = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WMlTSUbow98K+C3subL3YoHCx8WLTxQYgd9eFHUqxOtFKRiD5hRebLmwASvRc7FeKjQUA", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Номер = _t, Название = _t]),
    #"Измененный тип" = Table.TransformColumnTypes(Источник,{{"Номер", Int64.Type}, {"Название", type text}})
in
    #"Измененный тип"

// Спр_ББ_Уровень
let
    Источник = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WMlTSUbqw+GLDhX0XNl3YemHvxR4FsNDyCxsuNl5surBDKVYnWskIUxVYaMGFDRe2X9gCFNp9sRus0hhTJVhoIdCsDRebLvZATTTBVAcWmg8U2AK2Ga46FgA=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [ID_Уровень = _t, УровеньПрефикс = _t, ИмяУровня = _t]),
    #"Измененный тип" = Table.TransformColumnTypes(Источник,{{"ID_Уровень", Int64.Type}, {"ИмяУровня", type text}})
in
    #"Измененный тип"
	
	
	
	// Спр_ФормыОтчетности
let
    Источник = #table(type table [ID_Форма = text, Название = text, КраткоеНазвание = text, Примечание = text],{     
            {"0112","Баланс","ББ",""}, 
            {"0212","Отчет и прибылях и убытках", "ОПУ",""},
            {"0312","Отчет о доходах, расходах и финансовых результатах","ОДРФР",""},
            {"0412","Отчет о движении денежных средств","ОДДС",""},
            {"0512","Приложение к бухгалтерскому балансу","ПББ",""},
            {"3012","Отчет об изменении капитала","ОИК",""}
          }
)
in
    Источник





// fnRazdel
let
    Источник = (НомерРаздела as text) => 
let
    Номер = Record.FieldOrDefault([1 = 1, 2 = 2, 4 = 3, 5 = 4, 6 = 5, 0 = 0], НомерРаздела, null)
in 
    Номер
in
    Источник




// fnReorderSubsetOfColumns
/*  FnReorderSubsetOfColumns - Переупорядочиваение столбцов без привязки к названию других

    Функция получает:
        tbl - таблицу, 
        reorderedColumns - подмножество переупорядоченных имен столбцов в виде списка и 
        offset - индекс (номер колонки куда вставлять полонки по списку), начинающийся с нуля. 
        Затем выполняется переупорядочение очень надежным способом, без ссылок на имена других столбцов.

    Пример = FnReorderSubsetOfColumns(Буфер, {"Строка", "Код"}, 0)
*/
 
let
    Источник = (tbl as table, reorderedColumns as list, offset as number)
   as table =>
      Table.ReorderColumns(
         tbl,
         List.InsertRange(
            List.Difference(
            Table.ColumnNames(tbl),
            reorderedColumns
         ),
         offset,
         reorderedColumns
      )
   )
in
	Источник

// Спр_ББ4_Статья
let
    Источник = Б_Баланс_Чистка[[Статья],[Строка]],
    #"Переупорядоченные столбцы" = Table.ReorderColumns(Источник,{"Строка", "Статья"}),
    #"Удаленные дубликаты" = Table.Distinct(#"Переупорядоченные столбцы", {"Строка"}),
    #"Переименованные столбцы" = Table.RenameColumns(#"Удаленные дубликаты",{{"Строка", "ID_Строка"},{"Статья", "Статья Имя"}}),
    AddColumn = Table.AddColumn(#"Переименованные столбцы", "Строка Статья", each Text.From([ID_Строка]) & " " & [Статья Имя]),
    
    // Делаем множественную замену в подстроке столбца
   	МультиЗамена = [ 
           dict = List.Buffer(Table.FromValue({ 
                {"Hалог на добавленную стоимость", "НДС"},
                {"Долгосрочные активы, предназначенные для реализации", "Долгосрочные активы для реализации"},
                {"Долгосрочные обязательства по лизинговым платежам", "Долгосрочные обязательства по лизингу"},
                {"Денежные средства и эквиваленты денежных средств", "Денежные средства и их эквиваленты"}, 
                {"кредиторская задолженность", "КЗ"},
                {"дебиторская задолженность", "ДЗ"} 
               }) [Value]),
		   func = (x)=> List.Accumulate(dict, x, (state, current) => Text.Replace(state, current{0}, current{1} )),
		   b1   = Table.TransformColumns(AddColumn, {"Статья Имя", func}),
       b2   = Table.TransformColumns(b1, {"Строка Статья", func})
		 ][b2]
in
    МультиЗамена


let
    Источник = killerDateTableRu(#date(2016, 01, 1), #date(2021, 12, 31), null, true)
in
    Источник


// killerDateTableRu
// функция killerDateTableRu создает таблицу с датами. Дополняет ее различными полезными для Time Intelligence столбцами. 
// В случае, если четвертый аргумент принимает Дни 1, то функция добавляет столбец с выходными.
// Оригинальная идея philbritton https://gist.github.com/philbritton/9677152
// пример: killerDateTableRu(#date(2017, 2, 1), #date(2017, 2, 4), "RU-ru", 1)
// ошибка в четвертом параметре. Надо не 1 а TRUE ставить или тип данных менять
// пример: killerDateTableRu(#date(2017, 2, 1), #date(2017, 2, 4), "RU-ru", true)
 
// Описание полей результирующей таблицы:
// ===================================================
// Date - дата
// Year - год
// QuarterOfYear - номер квартала
// MonthOfYear - номер месяца в году
// MonthDayNumber - номер дня в месяце
// DateInt - дата в форме целого числа
// YearMonthNumber - МесяцГод в форме 201701
// MonthName - название месяца в формате Январь
// MonthInCalendar - Месяц год в формате Янв 2017
// QuarterInCalendar - КварталГод в формате Q1 2017
// DayInWeek - номер дня недели
// DayOfWeekName - название дня недели
// WeekEnding - дата окончания недели
// StartOfWeek - дата начала недели
// StartOfMonth - дата начала месяца
// WeekOfYear - номер недели в году
// DayOfYear - номер дня от начала года
// SequentialMonthNumber - номер месяца в таблице по порядку
// SequentialMonthNumberReverse - номер месяца в таблице в обратном порядке
// SequentialWeekNumber - номер недели в таблице по порядку
// SequentialWeekNumberReverse - номер недели в таблице в обратном порядке
// SequentialDayNumber - номер дня в таблице по порядку
// SequentialDayNumberReverse - номер дня в таблице в обратном порядке
 
( 
  StartDateParameter        as date, 
  EndDateParameter          as date, 
  CultureParameter          as nullable text, 
  IsRuHolidaysParameter     as nullable logical
) =>
 
let
    MONTHS_ORDERED_BY_DEFAULT = {
        "Январь", "Февраль", "Март",
        "Апрель", "Май", "Июнь", 
        "Июль", "Август", "Сентябрь", 
        "Октябрь", "Ноябрь", "Декабрь"
      },
 
    getProductionCalendar = ( URL as text ) as nullable table =>
      let
          GetDataFromGovRu =
            Csv.Document( 
              Web.Contents( URL ),
              [ Delimiter  = ",",
                Encoding   = 65001,
                QuoteStyle = QuoteStyle.None ]
            ),
 
          #"Повышенные заголовки" = 
            Table.PromoteHeaders( GetDataFromGovRu ),
          
          #"Переименовали поле в год" = 
            Table.RenameColumns( #"Повышенные заголовки",
              { 
                {"Год/Месяц", "Год"}
              }
            ),
 
          #"Выбрали год и месяц" =
            Table.SelectColumns( #"Переименовали поле в год", 
              { "Год" } & MONTHS_ORDERED_BY_DEFAULT 
            ),
 
          #"Спрямили календарь" = 
            Table.UnpivotOtherColumns( #"Выбрали год и месяц", 
              { "Год" }, "Месяц", "День"
            ),
          
          #"Добавили номер месяца" = 
            Table.AddColumn( #"Спрямили календарь", "НомерМесяца", each 
              1 + List.PositionOf( MONTHS_ORDERED_BY_DEFAULT, [Месяц] ), 
              type number
            ),
          
          #"Преобразовали строку дат в список" =
            Table.TransformColumns( #"Добавили номер месяца",
              { 
                { "День", each Text.Split( _, "," ) }
              }
            ),
 
          #"Развернули список дней" = 
            Table.ExpandListColumn( #"Преобразовали строку дат в список",
              "День" 
            ),
 
          #"Добавили сокращенный день" = 
            Table.AddColumn( #"Развернули список дней", "Сокращенный День", each 
              Text.Contains( [День], "*" )
            ),
 
          #"Добавили выходной день" = 
            Table.AddColumn( #"Добавили сокращенный день", "Выходной День", each 
              not [Сокращенный День]
            ),
 
          #"Удалили признак сокращенного дня" = 
            Table.ReplaceValue( #"Добавили выходной день",
              "*", "", Replacer.ReplaceText, { "День" }
            ),
 
          #"Уточнли тип полей календаря" = 
            Table.TransformColumnTypes( #"Удалили признак сокращенного дня",
              {
                { "Год", Int64.Type },
                { "День", Int64.Type },
                { "НомерМесяца", Int64.Type }
              }
            ),
 
          #"Рассчитали дату" = 
            Table.AddColumn( #"Уточнли тип полей календаря", "Дата", each
              #date( [Год], [НомерМесяца], [День] )
            )
      in
          #"Рассчитали дату",
 
// Примечание: 
//  источник взят отсюда: http://data.gov.ru/opendata/7708660670-proizvcalendar
//  На момент 24.05.2018 в календаре содержатся описание праздников с 1999 года по 2025 год.
    PROD_СALENDAR_URL =
      "http://data.gov.ru/opendata/7708660670-proizvcalendar/data-20180410T1145-structure-20180410T1145.csv?encoding=UTF-8",
 
    ProductionCalendar = getProductionCalendar( PROD_СALENDAR_URL ),
 
    ListOfShortday = 
      Table.SelectRows( ProductionCalendar, each [Сокращенный День] = true )[Дата],
    
    ListOfHoliday =
      Table.SelectRows( ProductionCalendar, each [Выходной День] = true )[Дата],
 
    CheckIfThereIsShortdayList = 
      List.Buffer( 
        try ListOfShortday otherwise {} 
      ),
 
    CheckIfThereIsHolidayList = 
      List.Buffer( 
        try ListOfHoliday otherwise {} 
      ),
 
    DayCount = 
      Duration.Days( 
        Duration.From( 
          EndDateParameter - StartDateParameter 
        ) 
      ) + 1,
    
    MainList = 
      List.Dates( StartDateParameter, DayCount, #duration( 1, 0, 0, 0 ) ),
    
    TableFromList = 
      Table.FromList( MainList, 
        Splitter.SplitByNothing(), { "Date" } 
      ),    
    
    ChangedType = 
      Table.TransformColumnTypes( TableFromList, 
        { 
          {"Date", type date } 
        } 
      ),
    
    InsertYear = 
      Table.AddColumn( ChangedType, "Year", each 
        Date.Year( [Date] ), 
        Int64.Type
      ),
    
    InsertQuarter = 
      Table.AddColumn( InsertYear, "QuarterOfYear", each 
        Date.QuarterOfYear( [Date] ),
        Int64.Type
      ),
    
    InsertMonth = 
      Table.AddColumn( InsertQuarter, "MonthOfYear", each
        Date.Month( [Date] ),
        Int64.Type 
      ),
    
    InsertDay = 
      Table.AddColumn( InsertMonth, "MonthDayNumber", each
        Date.Day( [Date] ),
        Int64.Type
      ),
    
    InsertDayInt = 
      Table.AddColumn( InsertDay, "DateInt", each 
        [Year] * 10000 + [MonthOfYear] * 100 + [MonthDayNumber],
        Int64.Type
      ),
    
    InsertYearMonthNumber = 
      Table.AddColumn( InsertDayInt, "YearMonthNumber", each 
        [Year] * 100 + [MonthOfYear] * 1,
        Int64.Type
      ),
    
    InsertMonthName = 
      Table.AddColumn( InsertYearMonthNumber, "MonthName", each 
        Date.ToText( [Date], "MMMM", CultureParameter ),
        type text
      ),
    
    InsertCalendarMonth = 
      Table.AddColumn( InsertMonthName, "MonthInCalendar", each
        ( try ( Text.Range( [MonthName], 0, 3 ) ) otherwise [MonthName] ) & " " & Number.ToText( [Year] ),
        type text
      ),
    
    InsertCalendarQtr = 
      Table.AddColumn( InsertCalendarMonth, "QuarterInCalendar", each 
        "Q" & Number.ToText( [QuarterOfYear] ) & " " & Number.ToText( [Year] ),
        type text
      ),
    
    InsertDayWeek = 
      Table.AddColumn(InsertCalendarQtr, "DayInWeek", each 
        Date.DayOfWeek( [Date], Day.Monday ) + 1, 
        Int64.Type
      ),
    
    InsertDayName = 
      Table.AddColumn(InsertDayWeek, "DayOfWeekName", each 
        Date.ToText( [Date], "dddd", CultureParameter ),
        type text
      ),
    
    InsertWeekEnding = 
      Table.AddColumn(InsertDayName, "WeekEnding", each 
        Date.EndOfWeek( [Date], Day.Monday ),
        type date
      ),
    
    InsertedStartofWeek = 
      Table.AddColumn(InsertWeekEnding, "StartOfWeek", each
        Date.StartOfWeek( [Date], Day.Monday ),
        type date
      ),
    
    InsertedStartofMonth = 
      Table.AddColumn(InsertedStartofWeek, "StartOfMonth", each 
        Date.StartOfMonth( [Date] ),
        type date
      ),
    
    InsertWeekofYear = 
      Table.AddColumn(InsertedStartofMonth, "WeekOfYear", each 
        Date.WeekOfYear( [Date], Day.Monday ),
        Int64.Type
      ),
    
    InsertDayofYear = 
      Table.AddColumn(InsertWeekofYear, "DayOfYear", each 
        Date.DayOfYear( [Date] ),
        Int64.Type
        ),
    
    listBufferMonths = 
      List.Buffer(
        List.Distinct(
          InsertDayofYear[StartOfMonth]
        )
        ),
    
    AddedNumberOfMonth = 
      Table.AddColumn( InsertDayofYear, "SequentialMonthNumber", each
        List.PositionOf( listBufferMonths, [StartOfMonth]) + 1,
        Int64.Type
      ),
    
    SequentialMonthNumberReverse = 
      Table.AddColumn( AddedNumberOfMonth, "SequentialMonthNumberReverse", each
        List.PositionOf(
          List.Reverse( listBufferMonths ),
          [StartOfMonth]
        ) + 1,
        Int64.Type
      ),
    
    listBufferWeeks =
      List.Buffer( 
        List.Distinct( SequentialMonthNumberReverse[StartOfWeek] )
      ),
    
    AddedNumberOfWeeks = 
      Table.AddColumn( SequentialMonthNumberReverse, "SequentialWeekNumber", each
        List.PositionOf( listBufferWeeks, [StartOfWeek] ) + 1,
        Int64.Type
      ),
 
    AddedNumberOfWeeksReverse = 
      Table.AddColumn( AddedNumberOfWeeks, "SequentialWeekNumberReverse", each
        List.PositionOf( List.Reverse( listBufferWeeks ), [StartOfWeek] ) + 1,
        Int64.Type
      ),
    
    InsertSequentialDayNumber =
      Table.AddIndexColumn(
        AddedNumberOfWeeksReverse,
        "SequentialDayNumber",
        1,
        1
      ),
 
    InsertSequentialDayNumberReverse =
      Table.AddIndexColumn( 
        InsertSequentialDayNumber,
        "SequentialDayNumberReverse",
        List.Max( InsertSequentialDayNumber[SequentialDayNumber] ),
        -1 
      ),
 
    insertRuHolidaysColumn = 
      Table.AddColumn( InsertSequentialDayNumberReverse, "Holiday", each
        if List.Count( CheckIfThereIsHolidayList ) = 0 
        then null
        else 
          if List.Contains( CheckIfThereIsHolidayList, [Date]) 
          then true 
          else false
      ),
 
    insertRuShortdaysColumn = 
      Table.AddColumn( insertRuHolidaysColumn, "Shortday", each
        if List.Count( CheckIfThereIsShortdayList ) = 0 
        then null
        else 
          if List.Contains( CheckIfThereIsShortdayList, [Date]) 
          then true 
          else false
      )
    
in
  if IsRuHolidaysParameter
  then insertRuShortdaysColumn
  else InsertSequentialDayNumberReverse