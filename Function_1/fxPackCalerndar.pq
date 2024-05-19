let Calendar = 
	[ readme = "Библиотека календарей v 1.1",
	
	
	// ---------------------------------------------------------------------------------------	
	// Мой календарь, с которым работаю
	fxMyCalendar = (StartDate as date, EndDate as date, optional Culture as nullable text) as table =>
		let
			DayCount        = Duration.Days(Duration.From(EndDate - StartDate)),
			Source          = List.Dates(StartDate,DayCount,#duration(1,0,0,0)),
			TableFromList   = Table.FromList(Source, Splitter.SplitByNothing()),    
			ChangedType     = Table.TransformColumnTypes(TableFromList,{{"Column1", type date}}),
			RenamedColumns  = Table.RenameColumns(ChangedType,{{"Column1", "Date"}}),
			
			InsertYear          = Table.AddColumn(RenamedColumns, 		"Year", 			each Date.Year([Date])),
			InsertQuarter       = Table.AddColumn(InsertYear, 			"QuarterOfYear", 	each Date.QuarterOfYear([Date])),
			InsertMonth         = Table.AddColumn(InsertQuarter, 		"MonthOfYear", 		each Date.Month([Date])),
			InsertDay           = Table.AddColumn(InsertMonth, 			"DayOfMonth", 		each Date.Day([Date])),
			InsertDayInt        = Table.AddColumn(InsertDay, 			"DateInt", 			each [Year] * 10000 + [MonthOfYear] * 100 + [DayOfMonth]),
			InsertMonthName     = Table.AddColumn(InsertDayInt, 		"MonthName", 		each Date.ToText([Date], "MMMM", Culture), type text),
			InsertCalendarMonth = Table.AddColumn(InsertMonthName, 		"MonthInCalendar", 	each (try(Text.Range([MonthName],0,3)) otherwise [MonthName]) & " " & Number.ToText([Year])),
			InsertCalendarQtr   = Table.AddColumn(InsertCalendarMonth, 	"QuarterInCalendar", each "Q" & Number.ToText([QuarterOfYear]) & " " & Number.ToText([Year])),
			InsertDayWeek       = Table.AddColumn(InsertCalendarQtr, 	"DayInWeek", 		each Date.DayOfWeek([Date])),
			InsertDayName       = Table.AddColumn(InsertDayWeek, 		"DayOfWeekName", 	each Date.ToText([Date], "dddd", Culture), type text),
			InsertWeekEnding    = Table.AddColumn(InsertDayName, 		"WeekEnding", 		each Date.EndOfWeek([Date]), type date)    
		in  InsertWeekEnding,

	// ---------------------------------------------------------------------------------------	
	// Календарпь Уварова
	fxCalendar1 = (StartDate as date, EndDate as date, optional Culture as nullable text) as table =>
		let
			DayCount        = Duration.Days(Duration.From(EndDate - StartDate)),
			Source          = List.Dates(StartDate,DayCount,#duration(1,0,0,0)),
			TableFromList   = Table.FromList(Source, Splitter.SplitByNothing()),    
			ChangedType     = Table.TransformColumnTypes(TableFromList,{{"Column1", type date}}),
			RenamedColumns  = Table.RenameColumns(ChangedType,{{"Column1", "Date"}}),
			
			InsertYear          = Table.AddColumn(RenamedColumns, "Year", each Date.Year([Date])),
			InsertQuarter       = Table.AddColumn(InsertYear, "QuarterOfYear", each Date.QuarterOfYear([Date])),
			InsertMonth         = Table.AddColumn(InsertQuarter, "MonthOfYear", each Date.Month([Date])),
			InsertDay           = Table.AddColumn(InsertMonth, "DayOfMonth", each Date.Day([Date])),
			InsertDayInt        = Table.AddColumn(InsertDay, "DateInt", each [Year] * 10000 + [MonthOfYear] * 100 + [DayOfMonth]),
			InsertMonthName     = Table.AddColumn(InsertDayInt, "MonthName", each Date.ToText([Date], "MMMM", Culture), type text),
			InsertCalendarMonth = Table.AddColumn(InsertMonthName, "MonthInCalendar", each (try(Text.Range([MonthName],0,3)) otherwise [MonthName]) & " " & Number.ToText([Year])),
			InsertCalendarQtr   = Table.AddColumn(InsertCalendarMonth, "QuarterInCalendar", each "Q" & Number.ToText([QuarterOfYear]) & " " & Number.ToText([Year])),
			InsertDayWeek       = Table.AddColumn(InsertCalendarQtr, "DayInWeek", each Date.DayOfWeek([Date])),
			InsertDayName       = Table.AddColumn(InsertDayWeek, "DayOfWeekName", each Date.ToText([Date], "dddd", Culture), type text),
			InsertWeekEnding    = Table.AddColumn(InsertDayName, "WeekEnding", each Date.EndOfWeek([Date]), type date)    
		in  InsertWeekEnding,
	
	// Празднки PФ
	fxCalendar2 = (Год) =>
		let
			TextYearParam = Text.From(Год) & (if Год=2020 then "b" else ""),
			Source = Lines.FromBinary(Web.Contents("http://www.consultant.ru/law/ref/calendar/proizvodstvennye",[RelativePath = Text.From(TextYearParam) & "/?"])),
			#"Converted to Table" 	= Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
			#"Filtered Rows" 		= Table.SelectRows(#"Converted to Table", each Text.StartsWith([Column1], "#(tab)#(tab)#(tab)<td class=")),
			#"Replaced Value" 		= Table.ReplaceValue(#"Filtered Rows","</tr><tr>","",Replacer.ReplaceText,{"Column1"}),
			#"Added Index" 			= Table.AddIndexColumn(#"Replaced Value", "#Month", 1, 1),
			#"Added Prefix" 		= Table.TransformColumns(#"Added Index", {{"Column1", each Text.Split(_,"</td><td class="), type text}}),
			#"Expanded Custom" 		= Table.ExpandListColumn(#"Added Prefix", "Column1"),
			#"Filtered Rows1" 		= Table.SelectRows(#"Expanded Custom", each not Text.Contains([Column1], "inactively")),
			#"Replaced Value1" 		= Table.ReplaceValue(#"Filtered Rows1","#(tab)#(tab)#(tab)<td class=","",Replacer.ReplaceText,{"Column1"}),
			#"Split Column by Delimiter" = Table.SplitColumn(#"Replaced Value1", "Column1", Splitter.SplitTextByEachDelimiter({">"}, QuoteStyle.Csv, false), {"Day type", "Day"}),
			#"Extracted Text Before Delimiter" = Table.TransformColumns(#"Split Column by Delimiter", {{"Day", each Text.BeforeDelimiter(_, "<"), type text}}),
			#"Changed Type" 		= Table.TransformColumnTypes(#"Extracted Text Before Delimiter",{{"Day", Int64.Type}}),
			#"Added Custom" 		= Table.AddColumn(#"Changed Type", "Year", each Год, Int64.Type),
			#"Added Custom1" 		= Table.AddColumn(#"Added Custom", "Date", each #date([Year],[#"#Month"],[Day]), type date),
			#"Inserted Month Name" 	= Table.AddColumn(#"Added Custom1", "Month", each Date.MonthName([Date]), type text),
			#"Inserted Day Name" 	= Table.AddColumn(#"Inserted Month Name", "Day name", each Date.DayOfWeekName([Date]), type text)
		in
			#"Inserted Day Name",
	
	// Календарь 3
	fxCalendar3 = (StartDate,EndDate) =>
		let
			Source = List.Dates,
			#"Invoked FunctionSource" = Source(StartDate, Duration.Days(DateTime.Date(DateTime.FixedLocalNow()) - EndDate) + 1, #duration(1, 0, 0, 0)),
			#"Table from List" = Table.FromList(#"Invoked FunctionSource", Splitter.SplitByNothing(), null, null, ExtraValues.Error),
			#"Changed Type1" = Table.TransformColumnTypes(#"Table from List",{{"Column1", type date}}),
			#"Renamed Columns" = Table.RenameColumns(#"Changed Type1",{{"Column1", "Дата"}}),

			// Добавление столбцов
			Год                 = Table.AddColumn(#"Renamed Columns", 	"Год", 			each Date.Year([Дата]), Int64.Type),
			#"Номер месяца"     = Table.AddColumn(Год, 					"Месяц", 		each Date.ToText([Дата],"MM"), type text),
			#"Номер дня"        = Table.AddColumn(#"Номер месяца", 		"Номер дня", 	each Date.ToText([Дата],"dd"), type text),
			#"Название дня"     = Table.AddColumn(#"Номер дня", 		"День", 		each Date.ToText([Дата],"ddd"), type text),
			#"Название месяца"  = Table.AddColumn(#"Название дня", 		"Мес", 			each Date.ToText([Дата],"MMM"), type text),
			#"Номер квартала"   = Table.AddColumn(#"Название месяца", 	"Квартал", 		each Date.QuarterOfYear([Дата]), type text),
			#"Номер недели"     = Table.AddColumn(#"Номер квартала", 	"Номер недели", each Date.WeekOfYear([Дата]), Int64.Type),
			#"Читабельная Неделя" = Table.AddColumn(#"Номер недели", 	"Неделя", 		each Date.ToText(Date.StartOfWeek([Дата]),"MM-dd")&"|"&Date.ToText(Date.EndOfWeek([Дата]),"MM-dd"), type text),

			#"Эта-прошлая" = Table.AddColumn(#"Читабельная Неделя", "Эта/прошлая неделя", 
					each if Date.IsInCurrentWeek([Дата]) = true  and Date.From([Дата]) <> Date.From(DateTime.LocalNow()) then "Эта" 
					else if Date.IsInPreviousWeek([Дата]) = true and Date.AddDays([Дата],7) < Date.From(DateTime.LocalNow()) then "Прошлая" 
					else "", type text),

			#"Номер дня недели" = Table.AddColumn(#"Эта-прошлая", "№ дня недели", 
					each Date.DayOfWeek([Дата]) + 1, Int64.Type),

			#"Этот/прошлый месяц" = Table.AddColumn(#"Номер дня недели", "Этот/прошлый месяц", 
					each if Date.IsInCurrentMonth([Дата]) = true and Date.From([Дата]) <> Date.From(DateTime.LocalNow()) then "Этот месяц" 
					else if Date.IsInPreviousMonth([Дата]) = true then "Прошлый месяц" 
					else "", type text),

			#"Эта/прошлая/позапрошлая" = Table.AddColumn(#"Этот/прошлый месяц", "Прошлая/Позапрошлая", 
					each if Date.IsInPreviousWeek([Дата]) = true  then "Прошлая" 
					else if Date.IsInPreviousWeek(Date.AddDays([Дата],7)) = true then "Позапрошлая" 
					else "", type text)
			in
				#"Эта/прошлая/позапрошлая",

	//----------------------------------------------------------------------------------------------------------------------------------------
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
	
	fxCalendarKill = ( 
		StartDateParameter        as date, 
		EndDateParameter          as date, 
		CultureParameter          as nullable text, 
		IsRuHolidaysParameter     as nullable logical ) =>
	
	let
		MONTHS_ORDERED_BY_DEFAULT = {
			"Январь", "Февраль", "Март",
			"Апрель", "Май", "Июнь", 
			"Июль", "Август", "Сентябрь", 
			"Октябрь", "Ноябрь", "Декабрь"
		},
	
		getProductionCalendar = ( URL as text ) as nullable table =>
			let
				GetDataFromGovRu = Csv.Document(  Web.Contents( URL ), [ Delimiter  = ",", Encoding   = 65001, QuoteStyle = QuoteStyle.None ] ),
		
				#"Повышенные заголовки"      = Table.PromoteHeaders( GetDataFromGovRu ),
				#"Переименовали поле в год"  = Table.RenameColumns( #"Повышенные заголовки", { {"Год/Месяц", "Год"} }),
				#"Выбрали год и месяц"       = Table.SelectColumns( #"Переименовали поле в год",  { "Год" } & MONTHS_ORDERED_BY_DEFAULT ),
				#"Спрямили календарь"        = Table.UnpivotOtherColumns( #"Выбрали год и месяц", { "Год" }, "Месяц", "День"),
				#"Добавили номер месяца"     = Table.AddColumn( #"Спрямили календарь", "НомерМесяца", each 1 + List.PositionOf( MONTHS_ORDERED_BY_DEFAULT, [Месяц] ), type number),
				#"Преобразовали в список" 	 = Table.TransformColumns( #"Добавили номер месяца",{ { "День", each Text.Split( _, "," ) }} ),
				#"Развернули список дней"    = Table.ExpandListColumn( #"Преобразовали в список", "День" ),
				#"Добавили сокращенный день" = Table.AddColumn( #"Развернули список дней", "Сокращенный День", each Text.Contains( [День], "*" ) ),
				#"Добавили выходной день"    = Table.AddColumn( #"Добавили сокращенный день", "Выходной День", each not [Сокращенный День] ),
				#"Удалили признак сокр. дня" = Table.ReplaceValue( #"Добавили выходной день", "*", "", Replacer.ReplaceText, { "День" } ),
				#"Уточнили тип полей" 		 = Table.TransformColumnTypes( #"Удалили признак сокр. дня",{
						{ "Год", Int64.Type },
						{ "День", Int64.Type },
						{ "НомерМесяца", Int64.Type } }
					),
				#"Рассчитали дату" = Table.AddColumn( #"Уточнили тип полей", "Дата", each #date( [Год], [НомерМесяца], [День] ) )
			in
				#"Рассчитали дату",
	
	// Примечание: 
	//  источник взят отсюда: http://data.gov.ru/opendata/7708660670-proizvcalendar
	//  На момент 24.05.2018 в календаре содержатся описание праздников с 1999 года по 2025 год.
	
		PROD_СALENDAR_URL =
		"http://data.gov.ru/opendata/7708660670-proizvcalendar/data-20180410T1145-structure-20180410T1145.csv?encoding=UTF-8",
	
		ProductionCalendar = getProductionCalendar( PROD_СALENDAR_URL ),
	
		ListOfShortday = Table.SelectRows( ProductionCalendar, each [Сокращенный День] = true )[Дата],
		ListOfHoliday  = Table.SelectRows( ProductionCalendar, each [Выходной День] = true )[Дата],
		
		CheckIfThereIsShortdayList = List.Buffer( try ListOfShortday otherwise {} ),
		CheckIfThereIsHolidayList  = List.Buffer( try ListOfHoliday otherwise {} ),

		DayCount      = Duration.Days( Duration.From( EndDateParameter - StartDateParameter ) ) + 1,
		MainList      = List.Dates( StartDateParameter, DayCount, #duration( 1, 0, 0, 0 ) ),
		TableFromList = Table.FromList( MainList, Splitter.SplitByNothing(), { "Date" } ),    
		ChangedType   = Table.TransformColumnTypes( TableFromList,  { {"Date", type date } } ),
		
		InsertYear    = Table.AddColumn( ChangedType,   "Year",           each Date.Year( [Date] ), Int64.Type ),
		InsertQuarter = Table.AddColumn( InsertYear,    "QuarterOfYear",  each Date.QuarterOfYear( [Date] ),Int64.Type),
		InsertMonth   = Table.AddColumn( InsertQuarter, "MonthOfYear",    each Date.Month( [Date] ),Int64.Type),
		InsertDay     = Table.AddColumn( InsertMonth,   "MonthDayNumber", each Date.Day( [Date] ),Int64.Type),
		InsertDayInt  = Table.AddColumn( InsertDay,     "DateInt",        each [Year] * 10000 + [MonthOfYear] * 100 + [MonthDayNumber],Int64.Type),

		InsertYearMonthNumber = Table.AddColumn( InsertDayInt, "YearMonthNumber",    each [Year] * 100 + [MonthOfYear] * 1, Int64.Type),
		InsertMonthName =       Table.AddColumn( InsertYearMonthNumber, "MonthName", each Date.ToText( [Date], "MMMM", CultureParameter ), type text),
		InsertCalendarMonth =   Table.AddColumn( InsertMonthName, "MonthInCalendar", each ( try ( Text.Range( [MonthName], 0, 3 ) ) otherwise [MonthName] ) & " " & Number.ToText( [Year] ), type text),
		InsertCalendarQtr =     Table.AddColumn( InsertCalendarMonth, "QuarterInCalendar", each "Q" & Number.ToText( [QuarterOfYear] ) & " " & Number.ToText( [Year] ), type text),

		// 1Кв 22
		InsertCalendarQtrRus = Table.AddColumn( InsertCalendarQtr,     "QuarterInCalendarR", each Number.ToText( [QuarterOfYear] ) & "Кв" & " " & Number.ToText( [Year] ), type text),  
		InsertDayWeek        = Table.AddColumn( InsertCalendarQtrRus,  "DayInWeek",      each Date.DayOfWeek( [Date], Day.Monday ) + 1, Int64.Type ),
		InsertDayName        = Table.AddColumn( InsertDayWeek,         "DayOfWeekName",  each Date.ToText( [Date], "dddd", CultureParameter ), type text),
		InsertWeekEnding     = Table.AddColumn( InsertDayName,         "WeekEnding",     each Date.EndOfWeek( [Date], Day.Monday ), type date),
		InsertedStartofWeek  = Table.AddColumn( InsertWeekEnding,      "StartOfWeek",    each Date.StartOfWeek( [Date], Day.Monday ), type date),
		InsertedStartofMonth = Table.AddColumn( InsertedStartofWeek,   "StartOfMonth",   each Date.StartOfMonth( [Date] ), type date ),
		InsertWeekofYear     = Table.AddColumn( InsertedStartofMonth,  "WeekOfYear",     each Date.WeekOfYear( [Date], Day.Monday ), Int64.Type),
		InsertDayofYear      = Table.AddColumn( InsertWeekofYear,      "DayOfYear",      each Date.DayOfYear( [Date] ), Int64.Type),

		listBufferMonths 	= List.Buffer( List.Distinct( InsertDayofYear[StartOfMonth] ) ),
		
		AddedNumberOfMonth 	= Table.AddColumn( InsertDayofYear, "SequentialMonthNumber", each List.PositionOf( listBufferMonths, [StartOfMonth]) + 1, Int64.Type),
		
		SequentialMonthNumberReverse = 
			Table.AddColumn( AddedNumberOfMonth, "SequentialMonthNumberReverse", each
				List.PositionOf( List.Reverse( listBufferMonths ), [StartOfMonth]) + 1,
				Int64.Type
			),
		
		listBufferWeeks = List.Buffer( List.Distinct( SequentialMonthNumberReverse[StartOfWeek] ) ),
		
		AddedNumberOfWeeks 			= Table.AddColumn( SequentialMonthNumberReverse, "SequentialWeekNumber", each List.PositionOf( listBufferWeeks, [StartOfWeek] ) + 1, Int64.Type),
		AddedNumberOfWeeksReverse 	= Table.AddColumn( AddedNumberOfWeeks, "SequentialWeekNumberReverse", each List.PositionOf( List.Reverse( listBufferWeeks ), [StartOfWeek] ) + 1, Int64.Type),
		InsertSequentialDayNumber 	= Table.AddIndexColumn( AddedNumberOfWeeksReverse, "SequentialDayNumber", 1, 1 ),
	
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
				if 	List.Count( CheckIfThereIsShortdayList ) = 0  then null else 
				if 	List.Contains( CheckIfThereIsShortdayList, [Date])  then true else false
		)
		
	in
		if IsRuHolidaysParameter
		then insertRuShortdaysColumn
		else InsertSequentialDayNumberReverse,


	// --------------------------------------------------------------------------------------------------------------------------------------
	// Календарь Товарища
	// fxComrad = (StartDate as date, EndDate as date, optional Culture as nullable text) as table =>
	fxComrad = (StartDate as date, EndDate as date, optional Culture as nullable text) as table =>
		let
			start = StartDate, 	// Excel.CurrentWorkbook(){[Name = "Параметры"]}[Content]{0}[Значение]
			finish = EndDate, 	// Excel.CurrentWorkbook(){[Name = "Параметры"]}[Content]{1}[Значение]
			tab_generate = Table.FromList(
					{Number.From(start) .. Number.From(finish)}, 
					Splitter.SplitByNothing(), {"Дата"} ), 

			col_type = Table.TransformColumnTypes( tab_generate, {{"Дата", type date}} ), 
			
			tab_add_col_1 = Table.AddColumn( col_type, 		"Год", 			each Date.Year([Дата]), 					Int64.Type ), 
			tab_add_col_2 = Table.AddColumn( tab_add_col_1, "Месяц", 		each Date.Month([Дата]),					Int64.Type ), 
			tab_add_col_3 = Table.AddColumn( tab_add_col_2, "День", 		each Date.Day([Дата]), 						Int64.Type ), 
			tab_add_col_4 = Table.AddColumn( tab_add_col_3, "Квартал", 		each Date.QuarterOfYear([Дата]), 			Int64.Type ), 
			tab_add_col_5 = Table.AddColumn( tab_add_col_4, "Неделя года", 	each Date.WeekOfYear([Дата], Day.Monday), 	Int64.Type ), 
			tab_add_col_6 = Table.AddColumn( tab_add_col_5, "День недели", 	each Date.DayOfWeek([Дата], Day.Monday) + 1,Int64.Type )
		in  tab_add_col_6,

	
	// -------------------------------------------------------------------------------------------------------------------------------------
	// Календарь с сайта буржуйского
	// https://forum.enterprisedna.co/t/extended-date-table-power-query-m-function/6390
	fnDateTable2 = 
	[ 
		fnDateTable = ( StartDate as date, EndDate as date, optional FYStartMonthNum as number, optional Holidays as list, optional WDStartNum as number, optional AddRelativeNetWorkdays as logical ) as table =>
			let
				FYStartMonth = List.Select( {1..12}, each _ = FYStartMonthNum ){0}? ?? 1,
				WDStart = List.Select( {0..1}, each _ = WDStartNum ){0}? ?? 0,
				CurrentDate = Date.From( DateTime.FixedLocalNow()),
				DayCount = Duration.Days( Duration.From( EndDate - StartDate)) +1,
				Source = List.Dates( StartDate, DayCount, #duration(1,0,0,0)),
				AddToday = if List.Contains(Source, CurrentDate) then Source else List.Combine( {Source, {CurrentDate}}),
				ToTable = Table.FromList(AddToday, Splitter.SplitByNothing(), type table [Date = Date.Type] ),
				InsertYear = Table.AddColumn(ToTable, "Year", each Date.Year([Date]), type number),
				InsertYearOffset = Table.AddColumn(InsertYear, "CurrYearOffset", each Date.Year([Date]) - Date.Year( Date.From(CurrentDate)), type number),
				InsertCompletedYear = Table.AddColumn(InsertYearOffset, "YearCompleted", each Date.EndOfYear([Date]) < Date.From( Date.EndOfYear(CurrentDate)), type logical),

				InsertQuarterNum = Table.AddColumn(InsertCompletedYear, "Quarter Number", each Date.QuarterOfYear([Date]), type number),
				InsertQuarter = Table.AddColumn(InsertQuarterNum, "Quarter", each "Q" & Number.ToText([Quarter Number]), type text),
				InsertStartOfQuarter = Table.AddColumn(InsertQuarter, "Start of Quarter", each Date.StartOfQuarter([Date]), type date),
				InsertEndOfQuarter = Table.AddColumn(InsertStartOfQuarter, "End of Quarter", each Date.EndOfQuarter([Date]), type date),
				InsertCalendarQtr = Table.AddColumn(InsertEndOfQuarter, "Quarter & Year", each "Q" & Number.ToText( Date.QuarterOfYear([Date])) & Date.ToText([Date], [Format = " yyyy"]), type text),
				InsertQuarternYear = Table.AddColumn(InsertCalendarQtr, "QuarternYear", each [Year] * 10 + [Quarter Number], type number),
				InsertQuarterOffset = Table.AddColumn(InsertQuarternYear, "CurrQuarterOffset", each ((4 * Date.Year([Date])) +  Date.QuarterOfYear([Date])) - ((4 * Date.Year(Date.From(CurrentDate))) +  Date.QuarterOfYear(Date.From(CurrentDate))), type number),
				InsertCompletedQuarter = Table.AddColumn(InsertQuarterOffset, "QuarterCompleted", each Date.EndOfQuarter([Date]) < Date.From(Date.EndOfQuarter(CurrentDate)), type logical),

				InsertMonth = Table.AddColumn(InsertCompletedQuarter, "Month", each Date.Month([Date]), type number),
				InsertStartOfMonth = Table.AddColumn(InsertMonth, "Start of Month", each Date.StartOfMonth([Date]), type date),
				InsertEndOfMonth = Table.AddColumn(InsertStartOfMonth, "End of Month", each Date.EndOfMonth([Date]), type date),
				InsertCalendarMonth = Table.AddColumn(InsertEndOfMonth, "Month & Year", each Text.Proper( Date.ToText([Date], [Format = "MMM yyyy"])), type text),
				InsertMonthnYear = Table.AddColumn(InsertCalendarMonth , "MonthnYear", each [Year] * 100 + [Month], type number),
				InsertMonthOffset = Table.AddColumn(InsertMonthnYear, "CurrMonthOffset", each ((12 * Date.Year([Date])) +  Date.Month([Date])) - ((12 * Date.Year(Date.From(CurrentDate))) +  Date.Month(Date.From(CurrentDate))), type number),
				InsertCompletedMonth = Table.AddColumn(InsertMonthOffset, "MonthCompleted", each Date.EndOfMonth([Date]) < Date.From(Date.EndOfMonth(CurrentDate)), type logical),
				InsertMonthName = Table.AddColumn(InsertCompletedMonth, "Month Name", each Text.Proper( Date.ToText([Date], "MMMM")), type text),
				InsertMonthShort = Table.AddColumn( InsertMonthName, "Month Short", each Text.Proper( Date.ToText([Date], "MMM")), type text),
				InsertMonthInitial = Table.AddColumn(InsertMonthShort, "Month Initial", each Text.Start([Month Name], 1) & Text.Repeat( Character.FromNumber(8203), Date.Month([Date]) ), type text),
				InsertDayOfMonth = Table.AddColumn(InsertMonthInitial, "Day of Month", each Date.Day([Date]), type number),
			
				InsertWeekNumber = Table.AddColumn(InsertDayOfMonth, "Week Number", each
					if Number.RoundDown((Date.DayOfYear([Date])-(Date.DayOfWeek([Date], Day.Monday)+1)+10)/7)=0
					then Number.RoundDown((Date.DayOfYear(#date(Date.Year([Date])-1,12,31))-(Date.DayOfWeek(#date(Date.Year([Date])-1,12,31), Day.Monday)+1)+10)/7)
					else if (Number.RoundDown((Date.DayOfYear([Date])-(Date.DayOfWeek([Date], Day.Monday)+1)+10)/7)=53 and (Date.DayOfWeek(#date(Date.Year([Date]),12,31), Day.Monday)+1<4))
					then 1 else Number.RoundDown((Date.DayOfYear([Date])-(Date.DayOfWeek([Date], Day.Monday)+1)+10)/7), type number),
				InsertStartOfWeek = Table.AddColumn(InsertWeekNumber, "Start of Week", each Date.StartOfWeek([Date], Day.Monday), type date),
				InsertWeekEnding = Table.AddColumn(InsertStartOfWeek, "End of Week", each Date.EndOfWeek( [Date], Day.Monday), type date),
				InsertCalendarWk = Table.AddColumn(InsertWeekEnding, "Week & Year", each "W" & Text.PadStart( Text.From( [Week Number] ), 2, "0") & " " & Text.From(Date.Year( Date.AddDays( Date.StartOfWeek([Date], Day.Monday), 3 ))), type text ),
				InsertWeeknYear = Table.AddColumn(InsertCalendarWk, "WeeknYear", each Date.Year( Date.AddDays( Date.StartOfWeek([Date], Day.Monday), 3 )) * 100 + [Week Number],  Int64.Type),
				InsertWeekOffset = Table.AddColumn(InsertWeeknYear, "CurrWeekOffset", each (Number.From(Date.StartOfWeek([Date], Day.Monday))-Number.From(Date.StartOfWeek(CurrentDate, Day.Monday)))/7, type number),
				InsertCompletedWeek = Table.AddColumn(InsertWeekOffset, "WeekCompleted", each Date.EndOfWeek( [Date], Day.Monday) < Date.From(Date.EndOfWeek(CurrentDate, Day.Monday)), type logical),
			
				InsertDayWeek = Table.AddColumn(InsertCompletedWeek, "Day of Week Number", each Date.DayOfWeek([Date], Day.Monday) + WDStart, Int64.Type),
				InsertDayName = Table.AddColumn(InsertDayWeek, "Day of Week Name", each Text.Proper( Date.ToText([Date], "dddd" )), type text),
				InsertDayInitial = Table.AddColumn(InsertDayName, "Day of Week Initial", each Text.Proper(Text.Start([Day of Week Name], 1)) & Text.Repeat( Character.FromNumber(8203), Date.DayOfWeek([Date], Day.Monday) + WDStart ), type text),
				InsertDayOfYear = Table.AddColumn(InsertDayInitial, "Day of Year", each Date.DayOfYear([Date]), Int64.Type),
				InsertDayInt = Table.AddColumn(InsertDayOfYear, "DateInt", each [Year] * 10000 + [Month] * 100 + [Day of Month], type number),
				InsertDayOffset = Table.AddColumn(InsertDayInt, "CurrDayOffset", each Number.From([Date]) - Number.From(CurrentDate), type number),
				InsertIsAfterToday = Table.AddColumn(InsertDayOffset, "IsAfterToday", each not ([Date] <= Date.From(CurrentDate)), type logical),
				InsertIsWorkingDay = Table.AddColumn(InsertIsAfterToday, "IsWeekDay", each if Date.DayOfWeek([Date], Day.Monday) > 4 then false else true, type logical),
				InsertIsHoliday = Table.AddColumn(InsertIsWorkingDay, "IsHoliday", each if Holidays = null then "Unknown" else List.Contains( Holidays, [Date] ), if Holidays = null then type text else type logical),
				InsertIsBusinessDay = Table.AddColumn(InsertIsHoliday, "IsBusinessDay", each if [IsWeekDay] = true and [IsHoliday] <> true then true else false, type logical),
				InsertDayType = Table.AddColumn(InsertIsBusinessDay, "Day Type", each if [IsHoliday] = true then "Holiday" else if [IsWeekDay] = false then "Weekend" else if [IsWeekDay] = true then "Weekday" else null, type text),

				InsertISOYear = Table.AddColumn( InsertDayType, "ISO Year", each Date.Year( Date.AddDays( Date.StartOfWeek([Date], Day.Monday), 3 )), type number),
				InsertISOqNum = Table.AddColumn(InsertISOYear, "ISO Quarter Number", each if [Week Number] >39 then 4 else if [Week Number] >26 then 3 else if [Week Number] >13 then 2 else 1, Int64.Type),
				InsertISOqtr = Table.AddColumn(InsertISOqNum, "ISO Quarter", each "Q" & Number.ToText([ISO Quarter Number]), type text),
				InsertISOQuarter = Table.AddColumn(InsertISOqtr, "ISO Quarter & Year", each "Q" & Number.ToText([ISO Quarter Number]) & " " & Number.ToText([ISO Year]), type text),
				InsertISOqNy = Table.AddColumn(InsertISOQuarter, "ISO QuarternYear", each [ISO Year] * 10 + [ISO Quarter Number], type number),

				// BufferTable = Table.Buffer(Table.Distinct( InsertISOqNy[[ISO Year], [DateInt]])),
				// InsertISOday = Table.AddColumn(InsertISOqNy, "ISO Day of Year", (OT) => Table.RowCount( Table.SelectRows( BufferTable, (IT) => IT[DateInt] <= OT[DateInt] and IT[ISO Year] = OT[ISO Year])),  Int64.Type),
				AddFY = Table.AddColumn(InsertISOqNy, "Fiscal Year", each "FY" & (if [Month] >= FYStartMonth and FYStartMonth >1 then Text.From([Year] +1) else Text.From([Year])), type text),
				//AddFYs = Table.AddColumn(AddFY, "Fiscal Year short", each "FY" & (if [Month] >= FYStartMonth and FYStartMonth >1 then Text.PadEnd( Text.End( Text.From([Year] +1), 2), 2, "0") else Text.End( Text.From([Year]), 2)), type text),
				AddFQ = Table.AddColumn(AddFY, "Fiscal Quarter", each "FQ" & Text.From( Number.RoundUp( Date.Month( Date.AddMonths( [Date], - (FYStartMonth -1) )) / 3 )) & " " & (if [Month] >= FYStartMonth and FYStartMonth >1 then Text.From([Year] +1) else Text.From([Year])), type text),
				AddFQnYr = Table.AddColumn(AddFQ, "FQuarternYear", each (if [Month] >= FYStartMonth and FYStartMonth >1 then [Year] +1 else [Year]) * 10 + Number.RoundUp( Date.Month( Date.AddMonths( [Date], - (FYStartMonth -1) )) / 3 ), type number),
				AddFM = Table.AddColumn(AddFQnYr, "Fiscal Period Number", each if [Month] >= FYStartMonth and FYStartMonth >1 then [Month] - (FYStartMonth-1) else if [Month] >= FYStartMonth and FYStartMonth =1 then [Month] else [Month] + (12-FYStartMonth+1), type number),
				AddFP = Table.AddColumn(AddFM, "Fiscal Period", each "FP" & Text.PadStart( Text.From([Fiscal Period Number]), 2, "0") & " " & (if [Month] >= FYStartMonth and FYStartMonth >1 then Text.From([Year] +1) else Text.From([Year])), type text),
				AddFMnYr = Table.AddColumn(AddFP , "FPeriodnYear", each (if [Month] >= FYStartMonth and FYStartMonth >1 then [Year] +1 else [Year]) * 100 + [Fiscal Period Number], type number),
				FYCalendarStart = #date( Date.Year(StartDate)-1, FYStartMonth, 1 ),
				InsertFFD = Table.AddColumn( AddFMnYr, "FiscalFirstDay", each if [Month] >= FYStartMonth and FYStartMonth >1 then #date( Date.Year([Date])+1, FYStartMonth, 1) else #date( Date.Year([Date]), FYStartMonth, 1), type date ),

				InitTable = Table.FromList( List.Transform( {Number.From(FYCalendarStart) .. Number.From(EndDate)}, Date.From), Splitter.SplitByNothing(), type table [DateFW = Date.Type]),
				AddFFD = Table.AddColumn( InitTable, "FiscalFirstDay", each if Date.Month([DateFW]) < FYStartMonth then #date(Date.Year([DateFW]), FYStartMonth, 1) else #date(Date.Year([DateFW]) + 1, FYStartMonth, 1)),
				AddFWSD = Table.AddColumn( AddFFD, "FWStartDate", each Date.AddYears(Date.StartOfWeek([DateFW], Day.Monday), 1)),
				Group1 = Table.Group( AddFWSD, {"FiscalFirstDay", "FWStartDate"}, {{"AllRows", each _, type table [DateFW = nullable date, FiscalFirstDay = date, FWStartDate = date]}}),
				Group2 = Table.Group( Group1, {"FiscalFirstDay"}, {{"AllRows2", each _, type table [FiscalFirstDay = date, FWStartDate = date, AllRows = table]}}),
				AddIndex = Table.AddColumn( Group2, "Custom", each Table.AddIndexColumn([AllRows2], "Fiscal Week Number", 1, 1) )[[Custom]],
				ExpandG2 = Table.ExpandTableColumn( AddIndex, "Custom", {"FiscalFirstDay", "FWStartDate", "AllRows", "Fiscal Week Number"}, {"FiscalFirstDay", "FWStartDate", "AllRows", "Fiscal Week Number"}), 
				ExpandG1 = Table.ExpandTableColumn( ExpandG2, "AllRows", {"DateFW"}, {"DateFW"} )[[DateFW], [Fiscal Week Number]],
				MergeFYW = Table.Join( InsertFFD, {"Date"}, ExpandG1, {"DateFW"}, JoinKind.LeftOuter, JoinAlgorithm.SortMerge ),
				FWlogic = List.Contains( {null}, FYStartMonthNum),
				UpdateFYWeek = if FWlogic then Table.ReplaceValue(MergeFYW, each [Fiscal Week Number], each if FYStartMonth =1 then [Week Number] else [Fiscal Week Number], Replacer.ReplaceValue, {"Fiscal Week Number"}) else MergeFYW,
				AddFYW = Table.AddColumn( UpdateFYWeek, "Fiscal Week", each if FWlogic then "F" & [#"Week & Year"] else if FYStartMonth =1 then "FW" & Text.PadStart( Text.From([Fiscal Week Number]), 2, "0") & Date.ToText([Date], " yyyy") else if Date.Month([Date]) < FYStartMonth then "FW" & Text.PadStart( Text.From([Fiscal Week Number]), 2, "0") & Date.ToText([Date], " yyyy") else "FW" & Text.PadStart(Text.From([Fiscal Week Number]), 2, "0") & " " & Text.From( Date.Year([Date])+1), type text),
				InsertFWeeknYear = Table.AddColumn(AddFYW, "FWeeknYear", each if FWlogic then [WeeknYear] else (if FYStartMonth =1 then Date.Year([Date]) else if Date.Month([Date]) < FYStartMonth then Date.Year([Date]) else Date.Year([Date])+1) * 100 + [Fiscal Week Number],  Int64.Type),
				
				CurrentDateRecord = Table.SelectRows(InsertFWeeknYear, each ([Date] = CurrentDate)),
				CurrentISOyear = CurrentDateRecord{0}[ISO Year],
				CurrentISOqtr = CurrentDateRecord{0}[ISO Quarter Number],
				CurrentYear = CurrentDateRecord{0}[Year],
				CurrentMonth = CurrentDateRecord{0}[Month],
				CurrentFiscalFirstDay = CurrentDateRecord{0}[FiscalFirstDay],
				PrevFiscalFirstDay = Date.AddYears(CurrentFiscalFirstDay, -1),
				CurrentFY = CurrentDateRecord{0}[Fiscal Year],
				CurrentFQ = CurrentDateRecord{0}[FQuarternYear],
				CurrentFP = CurrentDateRecord{0}[FPeriodnYear],
				CurrentFW = CurrentDateRecord{0}[FWeeknYear],

				InsertISOYrOffset = Table.AddColumn(InsertFWeeknYear, "ISO CurrYearOffset", each [ISO Year] - CurrentISOyear, type number),
				InsertISOQtrOffset = Table.AddColumn(InsertISOYrOffset, "ISO CurrQuarterOffset", each ((4 * [ISO Year]) +  [ISO Quarter Number]) - ((4 * CurrentISOyear) + CurrentISOqtr), type number),
				InsertFYoffset = Table.AddColumn(InsertISOQtrOffset, "Fiscal CurrYearOffset", each try (if [Month] >= FYStartMonth then [Year]+1 else [Year]) - (if CurrentMonth >= FYStartMonth then CurrentYear+1 else CurrentYear) otherwise null, type number),
				InsertCurrentFY = Table.AddColumn(InsertFYoffset, "IsCurrentFY", each if [Fiscal Year] = CurrentFY then true else false, type logical),
				InsertCurrentFQ = Table.AddColumn(InsertCurrentFY, "IsCurrentFQ", each if [FQuarternYear] = CurrentFQ then true else false, type logical),
				InsertCurrentFP = Table.AddColumn(InsertCurrentFQ, "IsCurrentFP", each if [FPeriodnYear] = CurrentFP then true else false, type logical),
				InsertCurrentFW = Table.AddColumn(InsertCurrentFP, "IsCurrentFW", each if [FWeeknYear] = InsertISOYrOffset then true else false, type logical),
				InsertPYTD = Table.AddColumn(InsertCurrentFW, "IsPYTD", each if CurrentYear-1 = [Year] and [Day of Year] <= CurrentDateRecord{0}[Day of Year] then true else false, type logical),
					ListPrevFYDates = List.Buffer( Table.SelectRows( Table.ExpandTableColumn( Table.NestedJoin(
						Table.AddIndexColumn( Table.RenameColumns( Table.TransformColumnTypes( Table.FromList( List.Dates( PrevFiscalFirstDay, Number.From(CurrentFiscalFirstDay-PrevFiscalFirstDay),#duration(1,0,0,0)), Splitter.SplitByNothing()),{{"Column1", type date}}), {{"Column1", "DateFY"}}), "Index", 1, 1), {"Index"}, 
						Table.AddIndexColumn( Table.RenameColumns( Table.TransformColumnTypes( Table.FromList( List.Dates( Date.AddYears( PrevFiscalFirstDay, -1), Number.From( PrevFiscalFirstDay - Date.AddYears( PrevFiscalFirstDay, -1)),#duration(1,0,0,0)), Splitter.SplitByNothing()),{{"Column1", type date}}), {{"Column1", "DateFY"}}), "Index", 1, 1)
						, {"Index"}, "Table", JoinKind.LeftOuter), "Table", {"DateFY"}, {"PrevDateFY"}), each [DateFY] <= CurrentDate)[PrevDateFY] ),
				InsertPFYTD = Table.AddColumn(InsertPYTD, "IsPFYTD", each if [Fiscal CurrYearOffset] = -1 and List.Contains(ListPrevFYDates, [Date] ) then true else false, type logical),
				InsertNetWorkdays = if AddRelativeNetWorkdays = true then Table.AddColumn(InsertPFYTD, "Relative Networkdays", each fxNETWORKDAYS( StartDate, [Date], Holidays ), type number ) else InsertPFYTD,
				fxNETWORKDAYS = (StartDate, EndDate, optional Holidays as list) =>
					let
					ListOfDates = List.Dates( StartDate, Number.From(EndDate-StartDate)+1, Duration.From(1) ),
					DeleteHolidays = if Holidays = null then ListOfDates else List.Difference( ListOfDates, List.Transform(Holidays, Date.From )),
					DeleteWeekends = List.Select( DeleteHolidays, each Date.DayOfWeek( _, Day.Monday) < 5 ),
					CountDays = List.Count( DeleteWeekends)
					in
					CountDays,
				RemoveToday = Table.RemoveColumns( if EndDate < CurrentDate then Table.SelectRows(InsertNetWorkdays, each ([Date] <> CurrentDate)) else InsertNetWorkdays, {"Day of Year", "FiscalFirstDay"}), 
				ChType = Table.TransformColumnTypes(RemoveToday,{{"Year", Int64.Type}, {"Quarter Number", Int64.Type}, {"Month", Int64.Type}, {"Day of Month", Int64.Type}, {"DateInt", Int64.Type}, {"Day of Week Number", Int64.Type}, {"ISO CurrYearOffset", Int64.Type}, {"ISO QuarternYear", Int64.Type}, {"ISO CurrQuarterOffset", Int64.Type}, {"Week Number", Int64.Type}, {"WeeknYear", Int64.Type}, {"MonthnYear", Int64.Type}, {"QuarternYear", Int64.Type}, {"FQuarternYear", Int64.Type}, {"Fiscal Period Number", Int64.Type}, {"FPeriodnYear", Int64.Type}, {"CurrWeekOffset", Int64.Type}, {"CurrMonthOffset", Int64.Type}, {"CurrQuarterOffset", Int64.Type}, {"CurrYearOffset", Int64.Type}, {"Fiscal CurrYearOffset", Int64.Type}, {"Fiscal Week Number", Int64.Type}}),
				ReorderCols = Table.ReorderColumns(ChType,{"Date", "Year", "CurrYearOffset", "YearCompleted", "Quarter Number", "Quarter", "Start of Quarter", "End of Quarter", "Quarter & Year", "QuarternYear", "CurrQuarterOffset", "QuarterCompleted", "Month", "Start of Month", "End of Month", "Month & Year", "MonthnYear", "CurrMonthOffset", "MonthCompleted", "Month Name", "Month Short", "Month Initial", "Day of Month", "Week Number", "Start of Week", "End of Week", "Week & Year", "WeeknYear", "CurrWeekOffset", "WeekCompleted", "Day of Week Number", "Day of Week Name", "Day of Week Initial", "DateInt", "CurrDayOffset", "IsAfterToday", "IsWeekDay", "IsHoliday", "IsBusinessDay", "Day Type", "ISO Year", "ISO CurrYearOffset", "ISO Quarter Number", "ISO Quarter", "ISO Quarter & Year", "ISO QuarternYear", "ISO CurrQuarterOffset", "Fiscal Year", "Fiscal CurrYearOffset", "Fiscal Quarter", "FQuarternYear", "Fiscal Period Number", "Fiscal Period", "FPeriodnYear", "DateFW", "Fiscal Week Number", "Fiscal Week", "FWeeknYear", "IsCurrentFY", "IsCurrentFQ", "IsCurrentFP", "IsCurrentFW", "IsPYTD", "IsPFYTD"}),
				ListCols = if FWlogic then Table.RemoveColumns(ReorderCols,{"ISO Quarter Number", "Fiscal Year", "Fiscal Quarter", "FQuarternYear", "Fiscal Period Number", "Fiscal Period", "FPeriodnYear", "DateFW", "Fiscal Week Number", "Fiscal Week", "FWeeknYear", "Fiscal CurrYearOffset", "IsCurrentFQ", "IsCurrentFP", "IsCurrentFW"}) else Table.RemoveColumns(ReorderCols,{"Fiscal Period Number", "DateFW", "Fiscal Week Number", "ISO Quarter Number"})
			in  ListCols,
		
		Documentation = 
		[
			Documentation.Name =  			" fxCalendar", 
			Documentation.Description = 	" Date table function to create an ISO-8601 calendar", 
			Documentation.LongDescription = " Date table function to create an ISO-8601 calendar", 
			Documentation.Category = 		" Table", 
			Documentation.Version = 		" 2.02: full code review",
			Documentation.Source = 			" local", 
			Documentation.Author = 			" Melissa de Korte", 
			Documentation.Examples = { 
				[	Description =  " See: https://forum.enterprisedna.co/t/extended-date-table-power-query-m-function/6390", 
					Code = " Optional paramters: #(lf)
						(FYStartMonthNum) Month number the fiscal year starts, Januari if omitted #(lf) 
						(Holidays) Select a query (and column) that contains a list of holiday dates #(lf) 
						(WDStartNum) Switch default weekday numbering from 0-6 to 1-7 by entering a 1 #(lf)
						(AddRelativeNetWorkdays) if true adds a Relative Networkdays column to the date table #(lf)
						#(lf)
						Important to note: #(lf)
						[Fiscal Week] starts on a Monday and can contain less than 7 days in a First- and/or Last Week of a FY #(lf)
						[IsWeekDay] does not take holiday dates into account  #(lf)
						[IsBusinessDay] does take optional holiday dates into account  #(lf)
						[IsPYTD] and [IsPFYTD] compare Previous [Day of Year] with the Current [Day of Year] number, so dates don't align in leap years #(lf)
						IMPORTANT! No Fiscal columns will be added if the (FYStartMonthNum) is omitted", 
					Result = " " 
				] }
		],
			
		result = Value.ReplaceType( fnDateTable, Value.ReplaceMetadata( Value.Type( fnDateTable ), Documentation ))
	] [result]

]
in
    Calendar