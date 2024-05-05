let

MyLib1 =[
	readme = "БИБЛИОТЕКА моих функий 1",
	Text.ContainsAny = (str, needles) as logical => List.Count(needles),
	to = Expression.Evaluate(Text.FromBinary(Binary.Buffer(File.Contents("G:\=EXCEL BI - мои проекты\2024.01 - Анализ Отчетности\2. PQ запросы\fxPackGlobal2.pq"))), #shared)
],

MyOffice =[
	readme = "БИБЛИОТЕКА функий по работе с файлами Word"
],

UTIL =[ // утилиты 
	LoadPq = // загрузка функций из файла. x - полный путь + имя файла с расширением
		(path)=>Expression.Evaluate(Text.FromBinary(Binary.Buffer(File.Contents(path))), #shared),
	
	Load   = // загрузка функций с Гитхаба / локальной папки
		[	func = (fnName as text, optional BasePath as text) as function => // as function  к чему обязывает?!
                let
                    DefaultPath = "G:\=EXCEL PQ\PQ-function-load\",
                    GitHubPath  = "https://raw.githubusercontent.com/hohlick/pquery/master/",
                        
                    BasePath = if (BasePath <> null) then BasePath else DefaultPath,
                    Path = BasePath & (if Text.End(BasePath, 1) <> "\" then "\" else ""),
                    File = Path & fnName & ".pq",
                    
                    Function = try Expression.Evaluate(Text.Replace(fnName, ".", "_"), #shared)  //если путь уже импортирован в рабочую книгу, просто используйте существующий путь
                    otherwise try  Expression.Evaluate(Text.FromBinary(Binary.Buffer(File.Contents(File))), #shared) //если еще не импортирован, попробуйте загрузить его из текстового файла в папке
                    otherwise      Expression.Evaluate(Text.FromBinary(Binary.Buffer(Web.Contents(GitHubPath & fnName & ".pq"))), #shared) // если папка не найдена - взять из GitHubPath
                in
                    Function
        ][func],
		
	LoadPath = // загрузка пути с листа текущей книги
			"Excel.CurrentWorkbook(){[Name=""Table1""]}[Content]{0}[Path]",				

	
	Files.SheetsInFolder = // загрузка файлов из папки
		/*
			Create a table with info on all sheets in any Excel files in a particular folder.
			Usage:
				Files.SheetsInFolder = Load("Files.SheetsInFolder"),
				Source = Files.SheetsInFolder("C:\path\to\my\folder\")
			Result: [a table containing the binary file content, file names, sheet tables, and sheet names for each sheet in each spreadsheet in the given folder]
		*/
		(folderPath as text) as table =>
		let
			Source 			= Folder.Files(folderPath),
			FilteredRows 	= Table.SelectRows(Source, each Text.Start([Extension],3) = ".xl"),
			RemovedOtherColumns = Table.SelectColumns(FilteredRows,{"Content", "Name"}),
			NoTemps 		= Table.SelectRows(RemovedOtherColumns, each not Text.StartsWith([Name], "~$")),
			InsertedCustom 	= Table.AddColumn(NoTemps, "Sheets", each Excel.Workbook([Content])),
			Expanded 		= Table.ExpandTableColumn(InsertedCustom, "Sheets", {"Data", "Name","Kind"}, {"D","N","K"}),
			NoPrintAreas 	= Table.SelectRows(Expanded, each not Text.Contains([N], "$")),
			OnlySheets 		= Table.SelectRows(NoPrintAreas, each Text.Contains([K], "Sheet"))
		in
			OnlySheets,
	
	fxSwitch = // возвращяет парный элемент из списка списков или пары списков
		/* 	Возвращяет парный элемент из списка
			Пример:
				1. 	to = F[fxSwitch]("B", { {"A", 1}, {"B", 2}, {"C", 3} })								// Return: 2
				2. 	to = F[fxSwitch]("TE", {"QB", "RB", "WR", "TE", "K", "D/ST"}, {1, 2, 3, 4, 5, 6}) 	// Return: 1
				3. 	to = F[fxSwitch](
							Date.DayOfWeek(DateTime.From("03/05/24")), // число
							{0, 1, 2, 3, 4, 5, 6}, 
							{"Воскресенье", "Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота"}, 
							null
						)
				3. 	to = F[fxSwitch](
							3, 	// Date.Month(DateTime.From(date))
							{1,2,3,4,5,6,7,8,9,10,11,12}, 
							{"Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"}, 
							null
						)
			
		*/
		(	value as any, 							// входящее значение
			casesOrPairs as list, 					// где искать, список списков или пара списков. Можно через List.Zip словарь сделать
			optional resultsOrDefault as any, 		// что возвращать
			optional default as any					// если ошибка
		) =>
		let
			hasPairs = List.First(casesOrPairs) is list,
			usingPairs = 
				let
					targetPosition = List.PositionOf(casesOrPairs, value, Occurrence.First, (case, theValue) => theValue = case{0})
				in
					if targetPosition = -1 then resultsOrDefault else casesOrPairs{targetPosition}{1},
			usingCases = 
				let
					cases = casesOrPairs,
					results = resultsOrDefault
				in
					if List.IsEmpty(cases) or List.IsEmpty(results) 
					then default 
					else 	if value = List.First(cases) 
							then List.First(results) 
							else @fxSwitch(value, List.Skip(cases, 1), List.Skip(results, 1), default)
		in
			if hasPairs then usingPairs else usingCases,
	
	fxSwitchDAX = // замена значения в ячейке (вся ячейка). Аналог реализации SWITCH в DAX
		/*	Описание:	замена значения в ячейке (вся ячейка). Могут быть логические переменные
						Первый элемент списка - проверяемое значение, далее идут пары, последним все остальные значения
			Результат:	любое значение из таблицы подстановки
			Сайт: 		https://xxlbi.com/blog/switch-true-in-power-query/
			Автор: 		by Daniil Maslyuk
			Пример:	1.	= Switch({	2, 
									1, "A", 
									2, "B"}
								), 						=> Result = "B"
			
					2.	= Switch({	"b", 
									"a", "first", 
									"b", "second", 
									"c"})				=> Result = "second"
					
					3.	= Switch({	true, 
									1 > 2, "first", 
									0 = 1, "second", 
									"other"})			=> Result = "other"
		*/
		[func = 
			(input as list) as any =>
				let
					Source         = List.Buffer(input),
					Expression     = List.First(Source),
					SkipExpression = List.Skip(Source),
					HasElse        = Number.IsEven(List.Count(Source)),
					ElseValue      = if HasElse then List.Last(Source) else null,
					ValuesResults  = List.RemoveLastN(SkipExpression, Number.From(HasElse)),
					Values         = List.Alternate(ValuesResults, 1, 1, 1),
					Results        = List.Alternate(ValuesResults, 1, 1, 0),
					FirstResult    = List.PositionOf(Values, Expression),
					FunctionResult = if FirstResult = -1 then ElseValue else Results{FirstResult}
				in
					FunctionResult,
			
			fnType = type function (input as list) as any
				meta [
					Documentation.Name = "Switch",
					Documentation.LongDescription = "Evaluates an expression against a list of values and returns one of multiple possible result expressions.",
					Documentation.Examples = {
						[Description = "Simple Switch expression", Code = "Switch({2, 1, ""A"", 2, ""B""})", Result = """B"""],
						[Description = "An equivalent of SWITCH(TRUE... in DAX", Code = "Switch({true, 1 > 2, ""A"", 1 < 2, ""B"", ""No result""})", Result = """B"""]
					}
				],
			result = Value.ReplaceType(func, fnType)
			][result],
	
	
	fxSwitchLite = // надо доделать
		(input) => 
		let values = {	{"E", "Employee"}, 
						{"S", "SCYC"}, 
						{"M", "Medical"},
						{input, "Undefined"} 
					}, 
			Result = List.First(List.Select(values, (x)=> x{0}=input)){1} 
		in  Result
],

TABLE =[ // создание готовых таблиц
	from1 = #table(type table [Product=text, Month=text, Sales=number], {} ),
	from2 = #table({"a".."d"},List.Repeat({{1..4}},5)),
	from3 = Table.FromList(List.Accumulate({1..6}, {}, (accum,x)=> accum & {Date.DayOfWeekName(x+1)}), null, {"ДеньНедели"}),
	from4 = Table.FromList(Text.Split("понедельник|вторник|среда|четверг|пятница|суббота|воскресенье","|"), null, {"ДеньНедели"}),
	from5 = // три столца. Изменять можно число строк и шаг
		#table(type table[Date=date,Month=number,Year=number],List.Transform(List.Dates(#date(2024,1,1), 10, #duration(1,0,0,0)),(x)=>{x, Date.Month(x), Date.Year(x)})),
	from6 = // таблица с одним столбцом со списком 12-ти месяцев (с возможностью изхменения формата)
		Table.Combine(List.Transform({1..12}, each Table.FromValue(Date.ToText(#date(2000,_,1),"MMMM"))))
],

libPQ = [ 
	Document = // функция документирования других функций
		( name as text, description as text,  valueOrExample as any, optional valueIfExample as any ) =>
		let 
			value 	 = if valueIfExample is null then valueOrExample else valueIfExample,
			examples = if valueIfExample is null then {} else valueOrExample
		in
			Value.ReplaceType(value, Value.Type(value) 
				meta [	Documentation.Name = name,
						Documentation.Description = description,
						// [Description = "", Code="", Result =""]
						Documentation.Examples = examples
				]
		),

	TEXT =[ // текстовые функции
		Text.Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz",
		Text.AlphaNumeric = Text.Alphabet & "0123456789",
		Number.Digits = {0,1,2,3,4,5,6,7,8,9}
	],
	
	SMALL =[
		Skip.AllMulti =  		// отсечение строк сверху и снизу таблицы по ключевым словам - по всей таблице + мульти выбор
			(x)=>(y)=>not List.ContainsAny({Text.Combine( List.Select( Record.ToList(y), (i)=>i is text))}, x, Text.Contains),
		
		Skip.RowsUpDown  =  	// по колонке + одно слово
			(x, txt)=>(y)=>not Text.Contains( Text.Combine( List.Select( Record.ToList(Record.SelectFields(y, txt)), (i)=>i is text)), x),

		fxIsNumber = 			// является ли первый символ числом
			(x)=>List.ContainsAny ({ Text.Start(x,1) }, {"0".."9"}, Text.Contains),
		
		Type.AllColumnText = 	// делаем все столбцы текстовыми
			(tbl as table) => Table.TransformColumnTypes(tbl, List.Transform(Table.ColumnNames(tbl), (x)=>{x,type text}) ),

		fxContainsKeyWord = 	// содержит ли таблица ключевое слово во всех столбцах в любом месте
			(table as table, keyword as text) as logical =>
				let f=(x)=> List.RemoveNulls(List.RemoveItems(x,{""})),
					g=(row as list)=> List.ContainsAny(row, {keyword}, Text.Contains),
					b = List.Transform(Table.ToColumns(table), f),
					h = List.AnyTrue(List.Transform(b, (x)=> g(x) )) in  h
		
	],

	//-------------------------------------------------------------------------------------
	REORDER = [
		Reorder.AllColumns =  	// перегруппировка всех столбцов таблицы согласно нового списка lst
			/*  
				Аргументы: 
				lst 	- 	новый список расположения столбцов список. 
							Он должен совпадать с исходным, меняеться только последовательность.
							Если имен каких-то столбуцов в новом списке не будет или ошибка в названии, то 
				offset 	- 	сдвиг относительно начала. 0 - без сдвига, 1 - первый столбец без изменений
				Пример: 1. 	на входе 4 колонки abcd. На выходе колонки a и b поставит в конец
							to = F[REORDER][Reorder.AllColumns](from,{"a","b"},2) 
							Result: {"c", "d", "a", "b"}
			*/
			[func = 
				(	tbl as table, 				//
					lst as list,				
					optional offset as number
				) => 
					let offset = if offset <> null then offset else 0
					in	Table.ReorderColumns(tbl,
						List.InsertRange(
							List.Difference(
								Table.ColumnNames(tbl),
								lst
							),
						offset, lst)
					),
			typ = type function (tbl as table, lst as list, optional offset as number) as table 
					meta [ 	Documentation.Name            = "Reorder.AllColumns (fxReorderAllColumns)",
							Documentation.LongDescription = "Перегруппировка всех столбцов таблицы согласно нового списка lst.
															 Список новых заголовков, число и названия колонок должны совпадать. <br>
															 Пример: lst = Value.Metadata(СправочникСтолбцов)[balance][Имя] <br>
															 		 F[REORDER][Reorder.AllColumns](from,{""a"",""b""},2)"],
			result = Value.ReplaceType(func,typ) 
			][result],
		
		Reorder.AllColumns2 = 	// перегруппировка всех столбцов таблицы согласно нового списка lst + сдвиг от начала
			[ to = 1
			]
	],
		
	LIST =[
		List.From = Document(
			"List.From", 
			"Преобразует текст, разделенный запятыми в список",
			{ [ Description = "Конвертирует текст в список", Code = "List.From(""{A, B, C}"")", Result = "{ ""A"", ""B"", ""C"" }"] },
			
			(simpleTextList as text) => [
				trimWhitespace 	= Text.Trim(simpleTextList),
				listToSplit 	= Text.TrimEnd(Text.TrimStart(trimWhitespace, "{"), "}"),
				Result 			= List.Transform(Text.Split(listToSplit, ","), each Text.Trim(_))
			] [Result]
			),
		
		List.ToText = Document(
			"List.ToText",
			"Конвертирует список в текст. Обратная функция List.From",
			{ [ Description = "Конвертирует текст в список", Code = "List.ToText({ 1, 2, ""abc"",1+2})", Result = """{12abc3}"""] },
			
			(list as list) => List.Accumulate(list, "{", (state, current) => state & Text.From(current)) & "}"
			),
	
		List.ToRecord = Document(
			/*  Transform a list of strings to a record using a given lambda (passed values k)
				The built-in Record.FromList only takes static lists...
				Usage:
				let
					List.ToRecord = Load("List.ToRecord"),
					list = {"a","b"}
				in
					List.ToRecord(list, (k) => Text.Upper(k))
				Result: [a="A", b="B"]	
			*/
			"List.ToRecord",
			"Конвертирует список в запись",
			{ [ Description = "Конвертирует список в запись", Code = "List.ToRecord({""a"",""b""}, (k) => Text.Upper(k))", Result = "[a=""A"", b=""B""]"] },
			
			(List as list, Lambda as function) as record =>
			let
				Transformed = List.Transform(List, Lambda)	//each Lambda(_)
			in
				Record.FromList(Transformed, List)
			)
	
	],

	//-------------------------------------------------------------------------------------
	CLEAN = [
		Column.RemoveToName = 
			// 1. Удаляем колонки по маске имени (начинается с ...)
			[ func = 
				(	tbl as table, 
					optional txt as text	// начальные символы имени колонки
				)=> 
				let from = if txt = null then "Column" else txt,
					to = Table.RemoveColumns(
							tbl, 
							List.Select(
								Table.ColumnNames(tbl), 
								(x)=>Text.StartsWith(x,from)
							) 
						)
				in to,
			
			fnType = type function (tbl as table, optional txt as text) as table
						meta [ 	Documentation.Name            = "Column.RemoveNameWithColumn - fxRemoveColumnsNameColumn",
								Documentation.LongDescription = " "
						],
			
			result = Value.ReplaceType(func, fnType)
			] [result],
				
		Column.RemoveNull = 
			// Удаляем в таблице пустые колонки - везде стоит null (fxRemoveColumnsNull)	
			(table as table, optional options) => [
				tbl = Table.Buffer(Table.Profile(table)), 
				lst = Table.SelectRows(tbl,(x)=>x[Count]=x[NullCount])[Column], 
				to  = Table.RemoveColumns(table,lst) 
			] [to],
		
		Row.RemoveNull  = 
			// удаление полностью пустых строк
			(tbl as table)=> 
				Table.SelectRows(tbl, each not List.IsEmpty(List.RemoveMatchingItems(List.Transform(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}), each try Text.Clean(Text.Trim(_)) otherwise _ ),{""})))

	],
	
	//-------------------------------------------------------------------------------------
	REPLACE = [
	
		Text.ReplaceFull = // заменяет ячейку целиком на основании словаря замен. Table.ToRows или Table.ToList
			/*
				to = Table.TransformColumns(add, {
						{"Область", 
						(x)=> List.First(
								List.Select(
									Table.ToRows(Замена), 
									(y)=> y{0}=x)
								){1}, // 1 столбец в таблице замен
						type text}
					})
			*/
			(tbl as table, input) =>
				List.First(
					List.Select(
						Table.ToRows(tbl), 	// переводим в список пар таблицу, состоящую из двух столбцов "Что менять" и "На что менять"
						each _{0}=input)	// 
				){1},						// меняем на второй столбец. Если 2, третий, если он есть


		Text.ReplaceAll_ = 
			/*
			Do multiple text replacements in one function call, passing the replacements as a list of lists
			Usage:
				Text.ReplaceAll = Load("Text.ReplaceAll"),
				Text.ReplaceAll("(test)", {
					{"(", "["},
					{")", "]"}
				})
			Result: "[test]"
			*/
			(str as text, Replacements as list) as text => 
				List.Accumulate(Replacements, str, (s, x) => Text.Replace(s, x{0}, x{1})),

        МультиЗамена = [ // Делаем множественную замену в подстроке столбца
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
                            {"Строка_ID", (x)=>x, Int64.Type}, 
                            {"СтатьяИмя",    func, type text}, 
                            {"СтатьяСтрока", func, type text} 
                        })
                ][transform],
		
		fxTextReplace = 
			//
			(dictionary as list, row) => List.Accumulate(dictionary, row, (s,c)=>Text.Replace(s,c{0},c{1})),
	
		fxTextReplaceList = [ // Заменяет в текстовой строке символы по списку
			/*  Функция : fxTextReplaceList    Отличная функция!!!
				Описание: Заменяет в любом месте тексте символы по списку, что на что
				Пример:  
					1. 	Если функция объявлена в теле запроса Text.ReplaceList = (InputText as text, ...) =>
						to = Text.ReplaceList("мама мыла раму",{"мама","мыла"},{"папа","чистил"}), 
					
					2.	Если отдельной функцией (тут источником другой запрос - в виде текста HTML код страницы - не список и не таблица!)
						to = fxTextReplaceList(#"HTML-код_All", {"scripts","<",">"}, {"","(",")"}, null)
						to = fxTextReplaceList("мама мыла раму",{"мама","мыла","раму"},{"папа","чистил","картошку"}, 1)
					
					3. 	Если заменить надо в колонке, то через Table.TransformColumns
						f=(x)=> fxTextReplaceList(x,
								{"А","Б","В"},
								{"1","2","3"}),
						tr = Table.TransformColumns(Transform3,{"ИмяКОлонки", f})
				
				OldList и  NewList лучше сделать таблицей замен или словарем и использовать как Data[Что] и Data[НаЧто] */
			func = (InputText as text,	// текст, где менять                
				OldList   as list, 	// список символов что менять    	{...} = > {"кот","+","1"}    / Data[Column1]{1} / {Source} / Data[Column1]
				NewList   as list,	// список символов на что менять	{...} = > {"собака","-","2"}
				optional Index as nullable number	// с какого вхождения начинать
				) => 
			
			let	ActualRow = if Index = null then 0 else Index,
				Result    = Text.Replace(InputText, OldList{ActualRow}, NewList{ActualRow}),
				NextRow   = ActualRow + 1,
				Output    = if   NextRow > (List.Count(OldList) - 1) 
							then Result          
							else @fxTextReplaceList(Result, OldList, NewList, NextRow)
			in 	Output,

			typ = type function (InputText as text, OldList as list, NewList as list,optional Index as nullable number) as table 
					meta [ 	Documentation.Name            = "fxTextReplaceList",
							Documentation.LongDescription = "Заменяет в любом месте тексте символы по списку, что на что <br>
									Пример: <br> 
									1. 	Если функция объявлена в теле запроса Text.ReplaceList = (InputText as text, ...) => <br>
										<i> to = Text.ReplaceList(""мама мыла раму"",{""мама"",""мыла""},{""папа"",""чистил""}), </i><br>
									
									2. 	Если отдельной функцией (тут источником другой запрос - в виде текста HTML код страницы - не список и не таблица!) <br>
										to = fxTextReplaceList(#""HTML-код_All"", {""scripts"",""<"","">""}, {"",""("","")""}, null) <br>
										to = fxTextReplaceList(""мама мыла раму"",{""мама"",""мыла"",""раму""},{""папа"",""чистил"",""картошку""}, 1) <br>
									
									3. 	Если заменить надо в колонке, то через Table.TransformColumns <br>
										f=(x)=> fxTextReplaceList(x, <br>
												{""А"",""Б"",""В""}, <br>
												{""1"",""2"",""3""}), <br>
										tr = Table.TransformColumns(Transform3,{""ИмяКОлонки"", f}) <br>
									
									OldList и  NewList лучше сделать таблицей замен или словарем и использовать как Data[Что] и Data[НаЧто]",
									
							Documentation.Examples = {
								[	Description = 	"Если функция объявлена в теле запроса Text.ReplaceList = (InputText as text, ...) =>", 
									Code 		= 	"to = Text.ReplaceList(""мама мыла раму"",{""мама"",""мыла""},{""папа"",""чистил""})",
									Result		=	".."
								], 
								[	Description = "Если отдельной функцией (тут источником другой запрос - в виде текста HTML код страницы - не список и не таблица!)",
									Code 		= "to = fxTextReplaceList(#""HTML-код_All"", {""scripts"",""<"","">""}, {"",""("","")""}, null) <br> 
									               to = fxTextReplaceList(""мама мыла раму"",{""мама"",""мыла"",""раму""},{""папа"",""чистил"",""картошку""}, 1)",
									Result		=".."
								]
							}
						],
			
			result = Value.ReplaceType(func,typ)
			][result]
	],

        fxParseHead = // сворачивание многоуровневых заголовков
		    [func = (
                table,				//
                num,				//
                optional delim,		//
                optional fill		//
                ) =>
                [ lst   = List.Buffer(
                                    Table.ToColumns(
                                        Table.TransformColumns(Table.Range(table,0,num), {}, Text.From)
                                    )
                                ), 
                        
                        delim = if delim = null then " " else delim,
                        
                        f=(x,y)=>List.Accumulate(	List.Zip({x,y}),
                                                    [ i=false, j={} ],
                                                    (s,c)=>[ i=s[i]=true or c{0}<>null, 
                                                            j = if i 
                                                                then s[j]&{c{0}} 
                                                                else s[j]&{c{0}??c{1}}]
                                                )[j],
                        g=(x)=> if fill=null 
                                then 	Text.Combine(x,delim) 
                                else 	Text.Combine(
                                            List.ReplaceValue(x,null,fill,Replacer.ReplaceValue),
                                            delim
                                        ),
                        
                        gen = List.Generate(    ()=>  [i = 0, l = lst{i}, o = l],
                                                (x)=> [i] < List.Count(lst),
                                                (x)=> [i = x[i] + 1, l = lst{i}, o = f(l, x[o])],
                                                (x)=> g(x[o])
                                            ),
                        out = 	Table.RenameColumns(
                                    Table.Range(table, num),
                                    List.Zip({Table.ColumnNames(table), gen})
                                )
            	] [out]
        	][func],
		
		fxFAQ = // справка по функциям. Например to = fxFAQ("list") 
			[func = (txt as text) =>
                let	list_1      = Record.ToTable(#shared), 
                    rows_select = Table.SelectRows( list_1, each Text.Contains(Text.Lower([Name]), txt) and Text.Contains([Name], ".") ), 
                    tab_add_col = Table.AddColumn( rows_select, "Meta", each Value.Metadata(Value.Type([Value])) ), 
                    cols_select = Table.SelectColumns(tab_add_col, {"Meta"}), 
                    col_expand_1 = 
                        Table.ExpandRecordColumn( cols_select, "Meta", 
                            {
                                "Documentation.Name", 
                                "Documentation.Description", 
                                "Documentation.LongDescription", 
                                "Documentation.Examples"
                            }, 
                            {"Name", "Description", "LongDescription", "Examples"}
                        ), 
                    col_expand_2 = Table.ExpandListColumn(col_expand_1, "Examples"), 
                    
                    col_expand_3 = 
                        Table.ExpandRecordColumn(
                            col_expand_2, 
                            "Examples", 
                            {"Description", "Code", "Result"}, 
                            {"Example Description", "Example Code", " Example Result"}
                        ),
                    Sort    = Table.Sort(col_expand_3,{{"Name", Order.Ascending}}),
                    Filtr   = Table.SelectRows(Sort, each ([Name] <> null))
                in Filtr
			] [func],
	
		fxAddMultiColumns = // создание нескольких столбцов
			[ func = (table as table, addedColumns as list) as table =>
					List.Accumulate(
						addedColumns, 
						table, 
						(t, d) => Table.AddColumn(t, d{0}, d{1}, d{2}?)
					),
				typ = type function (table as table, addedColumns as list) as table 
						meta [ 	Documentation.Name            = "fxAddMultiColumns",
								Documentation.LongDescription = "функция, создающая сразу несколько столбцов. Например = fxAddMultiColumns(tbl,{'дата', each 1})",
								Documentation.Examples = {
									[Description = "добавить один столбец", Code = "=fxAddMultiColumns(#table({""a"",""b"",""c""},{}),{1,""первый""})",Result="#table({""первый"",""b"",""c""},{})"], 
									[Description = "добавить два столбец", Code =  "=fxAddMultiColumns(#table({""a"",""b"",""c""},{}),{-1,""первый с конца""})",Result="#table({""a"",""b"",""первый с конца""},{})"]
								}
							],
				
				result = Value.ReplaceType(func,typ)
			][result],
			
	
		Table.fxClearAll = // очистка текста о символов + пустых столбцов и строк
		  	[func = (
				tbl, 						// таблица
				optional clearData, 		// удалить непечатные	
				optional clearColumn, 		// удалить пустые столбцы
				optional clearRows			// удалить пустые строки
			) =>

			[ 	// делаем все колонки текстовыми
				tr = Table.TransformColumns(tbl, {}, Text.From),
			
				// удаляем непечатные символы
				p = if clearData is null 
						then tr 
						else [ k  = (x) => 
									if x = null 
									then x 
									else Text.Combine( List.RemoveMatchingItems(Text.SplitAny(Text.Clean(x), clearData), {""}), " "),
							to = Table.TransformColumns(tr, {}, k)
							][to],

				// удаляем пустые колонки
				cl = if clearColumn is null 
						then p 
						else [ a = List.Sort(Table.ColumnNames(p)),
							b = Table.Profile(p, {{"tmp", (x) => true, List.NonNullCount}})[tmp],
							c = List.PositionOf(b, 0, Occurrence.All),
							d = List.Transform(c, (x) => a{x}),
							e = Table.RemoveColumns(p, d)
							][e],

				// удаляем пустые строки
				to = if clearRows is null 
						then cl 
						else Table.SelectRows(cl, each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null})))
				/*	// убираем двойные пробелы во всей таблице
					w = [ f = (this) => Text.Combine(List.Select(Text.Split(this, " "), each _ <> ""), " "),
							m = Table.TransformColumns(to, {}, f)
						][m]
				*/	  		
			] [to],
				
			typ = type function (tbl as table, optional clearData as text, optional clearColumn as text, optional clearRows as text) as table 
					meta [  Documentation.Name              = "Table.ClearAll",
						    Documentation.LongDescription   = "функция, создающая сразу несколько столбцов"],

			result = Value.ReplaceType(func,typ)
			] [result],
		
		
		Table.ReorderSubsetOfColumns = // переупорядочиваение столбцов без привязки к названию других
			/*  Функция получает:
					tbl - таблицу, 
					reorderedColumns - подмножество переупорядоченных имен столбцов в виде списка и 
					offset - индекс (номер колонки куда вставлять полонки по списку), начинающийся с нуля. 
					Затем выполняется переупорядочение очень надежным способом, без ссылок на имена других столбцов.
				Пример: to = FnReorderSubsetOfColumns(Буфер, {"Строка", "Код"}, 0)
				Одной строкой:
				func = (tbl, lst, n) => Table.ReorderColumns(tbl,List.InsertRange(List.Difference(Table.ColumnNames(tbl),lst),n,lst)),

			*/
            [func = (tbl as table, 
					reorderedColumns as list, 
					offset as number
                ) as table =>
            
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
			),
            
            typ = type function (tbl as table, reorderedColumns as list, offset as number) as table 
					meta [  Documentation.Name              = "Table.ReorderSubsetOfColumns",
						    Documentation.LongDescription   = "Функция получает:
                                tbl                 - таблицу, 
                                reorderedColumns    - подмножество переупорядоченных имен столбцов в виде списка и 
                                offset              - индекс (номер колонки куда вставлять полонки по списку), начинающийся с нуля. 
                                Затем выполняется переупорядочение очень надежным способом, без ссылок на имена других столбцов"
                ],
			
			result = Value.ReplaceType(func,typ)
        	][result],
		
		Table.fxRemoveEmptyColumns = // удаление пустых столбцов
			[ func = (table, optional options) =>
				let	tbl = Table.Buffer(Table.Profile(table)),
					lst = Table.SelectRows(tbl,(x)=>x[Count]=x[NullCount])[Column],
					to  = Table.RemoveColumns(table,lst)
				in  to,
			
				typ = type function (table as table, optional options as record) as table 
						meta [  Documentation.Name         = "fxTableRemoveEmptyColumns (@buchlotnik)",
								Documentation.LongDescription = "функция, удаляющая пустые столбцы из таблицы"],
				
				result = Value.ReplaceType(func,typ)
			][result],
				
		Table.fxRemoveEmptyRows = // Удаление пустых строк
			1,
		
		fxCompareGroups = // проверяет, если ли тексте символы из списка
			/* Например: 
				1. 	Создаем нужный список (группа) ключей и проверяем вот таким образом и отбираем строки:
					to = Table.SelectRows(Таблица, each fnCompareGroups(Ключи1, [Текст]) 
				2.  
			*/		
			(keys as list, text as text) as logical =>
			let Source = List.AnyTrue( List.Transform( keys, (key)=> Text.Contains(text, key, Comparer.OrdinalIgnoreCase) ) )
			in 	Source,

		TableRenameColumnsByPositions = // переименование столбца по номеру позиции
			(table,list) =>
				[   a = List.Buffer(Table.ColumnNames(table)),
					b = List.Count(a),
					c = (x) =>{a{if x{0}>0 then x{0}-1 else b+x{0}},x{1}},
					d = if list{0} is list then List.Transform(list,c) else c(list),
					e = Table.RenameColumns(table,d)
				]  [e],
		
		fxRenameColumns = // переименовать все столбцы по новому списку
        	[ func = (table,listOld,listNew) => 
				Table.RenameColumns(
                    table, 
                    List.Zip({ List.Buffer( listOld ), listNew})
				)
        	][func],
		
		fxРегион = // номер региона РБ
        	[ func = (КодРегиона as text) => 
                Record.FieldOrDefault(
                    [1 = "Брест", 2 = "Витебск", 3 = "Гомель", 4 = "Гродно", 5 = "Минск", 6 = "Могилев", 0 = "БКС"], 
                    КодРегиона, 
                    null
			    )
        	] [func]
	
], // конец библиотеки libPQ


extensionLibrary = [
	readme._extensionLibrary = "БИБЛИОТЕКА ФУНКЦИЙ _extensionLibrary",
	Document = 
		// 
		(name as text, description as text, valueOrExample as any, optional valueIfExample as any) =>
		let
			value = if valueIfExample is null then valueOrExample else valueIfExample,
			examples = if valueIfExample is null then {} else valueOrExample
		in
			Value.ReplaceType(value, Value.Type(value) meta [
				Documentation.Name = name, 
				Documentation.Description = description,  
				// [Description = "", Code="", Result =""]
				Documentation.Examples = examples
			]),
	
	//======================== 
	// Date
	//========================
	Date.Calendar =	Document(
			// Basic calendar
			"Date.Calendar",
			"Generate a calendar table for a given date span - can be text or proper dates. Current columns are Date, DayOfWeek, Month, MonthNum, WeekStartData, WeekStart, Year, YearMonth",
			{[ Description = "2016 calendar", Code ="PBI[Date.Calendar](""1/1/2016"", ""12/31/2016""", Result = "2016 calendar"]},
			(start as any, end as any) => 
				let
					StartDate 	= Date.From(start),
					EndDate 	= Date.From(end),
					Source 		= Date.DatesBetween(StartDate, EndDate),
					FromList 	= Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
					Date 		= Table.RenameColumns(FromList,{{"Column1", "Date"}}),
					DayOfWeek 	= Table.AddColumn(Date, "Day of Week", each Date.DayName([Date])),
					Month 		= Table.AddColumn(DayOfWeek, "Month", each Date.MonthName([Date])),
					MonthNum 	= Table.AddColumn(Month, "MonthNumber", each Date.Month([Date])),
					WeekStartDate = Table.AddColumn(MonthNum, "WeekStartDate", each Date.StartOfWeek([Date])),
					WeekStart	= Table.AddColumn(WeekStartDate, "Week Start", each [Month] & " " & Text.From(Date.Day([WeekStartDate]))),
					Year 		= Table.AddColumn(WeekStart, "Year", each Date.Year([Date])),
					YearMonth	= Table.AddColumn(Year, "YearMonth", each Number.From(Text.From([Year]) & (if [MonthNumber] < 10 then "0" else "") & Text.From([MonthNumber]))),
					Result 		= YearMonth
			in
					Result ),

	Date.DatesBetween = Document(
		"Date.DatesBetween",
		"Returns a list of dates in a given span (inclusive). Start and end parameters can be any order",
		{[Description = "Date range", Code = "PBI[Date.DatesBetween](""1/1/2016"", ""1/3/2016"")", Result="{""1/1/2016"", ""1/2/2016"", ""1/3/2016""}" ]},
		(start as any, end as any) => 
		let 
			StartDate = Date.From(start), 
			EndDate = Date.From(end),
			adjustedStart = List.Min({StartDate, EndDate}),
			adjustedEnd = List.Max({StartDate, EndDate}),
			GetDates = (start as date, end as date, dates as list)=> if start > end then dates else @GetDates(Date.AddDays(start, 1), end, List.Combine({dates, {start}})),
			Dates = GetDates(adjustedStart, adjustedEnd, {})
		in Dates ),

	Date.DayName = Document(
		"Date.DayName",
		"Returns the English day of the week name for a date",
		{[ Description = "Get the day name", Code="Date.DayName(""9/9/2016"")", Result="Friday"]},
		(date as any) => Switch(Date.DayOfWeek(DateTime.From(date)), {0, 1, 2, 3, 4, 5, 6}, {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"}, null)),

	Date.MonthName = Document(
		"Date.MonthName",
		"Returns the English month name for a date",
		{[ Description = "Get the month name", Code="PBI[Date.MonthName](""9/9/2016"")", Result = "September" ]},
		(date as any) => 
			let 
				monthNumber = if date is number then date else Date.Month(DateTime.From(date))
			in 
				Switch(
					monthNumber,
					{1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, 
					{"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}, null)),

	//======================== 
	// List
	//======================== 
	List.Flatten = Document(
		"List.Flatten",
		"Recursively flattens list elements. The end result is a single list",
		{ [ Description = "Flattening nested lists into one", Code = "List.Flatten({ 1, 2, Table.FromRecords({[x=1]}), {3, 4, 5} })", Result = "{ 1, 2, Table.FromRecords({[x=1]}), 3, 4, 5}"] },
		
		(list as list) => List.Accumulate(list, {}, (state, current) =>
			let
				currentListContent = if current is list then @List.Flatten(current) else {current}
			in
				List.Combine({state, currentListContent})
		)),

	List.From = Document(
		"List.From", 
		"Converts a text representation of a list into a list of the elements. Items are considered to be split by ,",
		{ [ Description = "Convert a text list", Code = "List.From(""{A, B, C}"")", Result = "{ ""A"", ""B"", ""C"" }"] },
		(simpleTextList as text) =>
			let
				trimWhitespace = Text.Trim(simpleTextList),
				listToSplit = Text.TrimEnd(Text.TrimStart(trimWhitespace, "{"), "}"),
				Result = List.Transform(Text.Split(listToSplit, ","), each Text.Trim(_))
			in
				Result),

	List.ToText = Document(
		"List.ToText",
		"Converts a list to a textual representation. Inverse of List.From",
		{ [ Description = "Conver to text", Code = "List.ToText({ 1, 2, 3})", Result = """{1, 2, 3}"""] },
		(list as list) =>
				List.Accumulate(list, "{", (state, current) => current & Text.From(current)) & "}"),

	//======================== 
	// Number
	//======================== 
	Number.Digits = {0,1,2,3,4,5,6,7,8,9},
	Number.ParseText = Document(
		"Number.ParseText",
		"Returns the first number (1+ consecutive digits) in the given text. Optionally, provide allowed characters to ignore them",
		{[ Description = "Allow commas", Code="PBI[Number.ParseText](""It's over 9,000!"", 0, {"",""})", Result = "9000" ]},
		(text as text, optional startIndex as number, optional allowCharacters as list) => 
			let
				consider = if startIndex is null then text else Text.Range(text,startIndex), 
				_allowCharacters = if allowCharacters is null then {} else allowCharacters,
				numberSeries = List.FirstN(List.Skip(Text.ToList(consider), each not Text.IsNumber(_)), each Text.IsNumber(_) or List.Contains(_allowCharacters, _))
			in 
				if text is null then null else Text.FromList(numberSeries)	),
	Number.ToLetters = Document(
			"Number.ToLetters",
			"Converts a number  (starting at 1) to an alphabet representation. Works like column headers in Excel.",
			{[ 
				Description = "Column 27", 
				Code = "PBI[Number.ToLetters](27)", 
				Result = "AB"
			]},
			(value as number) =>
				let
					GetLetter = (num as number) => 
						let 
							number = Number.Mod(num, 26),
							val = if number = 0 then 26 else number,
							valid = number < 26 and number > 0 
						in 
							if valid then Text.At(Text.Alphabet, val - 1) else error "Can't get letter for " & Text.From(num),
					func = (value as number, factor as number) =>
						let
							ThisLetter = GetLetter(Number.RoundDown(value/Number.Power(26, factor))),
							Result = if value <= Number.Power(26, factor) then "" else @func(value, factor+1) & ThisLetter
						in
							Result
				in
					if value <= 26 then GetLetter(value) else func(value, 1) & GetLetter(value)	),
	
	//======================== 
	// Splitters
	//======================== 
	Splitter.SplitTextByNonAlpha = Document(
		"Splitter.SplitTextByNonAlpha",
		"Splits text by characters that aren't [A-Za-z]",
		{[ Description = "Split text", Code="PBI[Splitter.SplitTextByNonAlpha](""A1B,C"")", Result = "{ ""A"", ""B"", ""C"" }"]},
		(line as text) => Splitter.SplitTextByNotIn(Text.Alphabet)	),

	Splitter.SplitTextByNotIn = Document(
		"Splitter.SplitTextByNotIn",
		"Splits text on any characters that aren't the provided 'safe' characters",
		{[ Description = "Split on non-alphanumeric", Code = "PBI[Splitter.SplitTextByNotIn](PBI[Text.Alphanumeric])(""Power BI is #1"")", Result = "{""Power BI is "", ""1""}"]},
		(safeCharacters as text) => (line as nullable text) =>
			if line is null then 
				{} 
			else
				List.Accumulate(Text.ToList(line), {null} , (state, current) => 
				let
					doSkip = not Text.Contains(safeCharacters, current),
					lastItem = List.Last(state),
					appendLast = lastItem<>null
				in
					if doSkip then 
						if lastItem is null then 
							state 
						else 
						List.Combine({state, {null}})
					else if appendLast then
						List.Combine({List.RemoveLastN(state, 1), {lastItem & current}})
					else  
						List.Combine({List.RemoveLastN(state, 1), {current}}))	),


	
	//===========================
	// Switch
	//===========================
	Switch = 
		/*

		*/
		Document(
			"Switch", 
			"Given a value, find it's paired item <br>"&
				"Switch(value as any, cases as list, results as list, optional default as any) <br>" &
				"Switch(value as any, pairs as list, optional default as any)",
			{ 
				[ Description = "Using separate lists", Code = "Switch(1, {1, 2, 3}, {""A"", ""B"", ""C""})", Result = "A"],
				[ Description = "Using one paired list", Code = "Switch(1, {{1, ""A""}, {2, ""B""}, {3, ""C""}})", Result = "A"]
			},
			(value as any, casesOrPairs as list, optional resultsOrDefault as any, optional default as any) =>
				let
				hasPairs = List.First(casesOrPairs) is list,
				usingPairs = 
					let
						targetPosition = List.PositionOf(casesOrPairs, value, Occurrence.First, (case, theValue) => theValue = case{0})
					in
						if targetPosition = -1 then resultsOrDefault else casesOrPairs{targetPosition}{1},
				usingCases = 
					let
						cases = casesOrPairs,
						results = resultsOrDefault
					in
						if List.IsEmpty(cases) or List.IsEmpty(results) then default else if value = List.First(cases) then List.First(results) else @Switch(value, List.Skip(cases, 1), List.Skip(results, 1), default)
				in
					if hasPairs then usingPairs else usingCases		),

	//===========================
	// Table
	//===========================
	Table.AddBlankRow = 
		// 
		(table as table) => Table.Combine({table, Table.FromRecords({[]})}),
	
	Table.DrillIntoColumn = 
		// Transforms a column's value into its nested value -- if it eventually finds only one. Consider the following column:
		//  MyCol
		//  -----
		//  "a"
		//  {{{"b"}}}
		//  Table.FromRecords({[MyCol=Table.FromRecords({[col=2]})})
		//  {}
		//
		// Table.DrillColumn(table, "MyCol") will convert it to
		//
		//  MyCol
		//  -----
		//  "a"
		//  "b"
		//  2
		//  null	
		(table as table, columnName as text) =>
		let
			FindValue = (value as any) => 
				if value is list then
				if List.Count(value) = 1 then @FindValue(List.First(value)) 
				else if List.Count(value) = 0 then null
				else error "Couldn't find single value"
				else if value is table then
				if Table.RowCount(value) = 1 then @FindValue(List.First(Table.ToColumns(value)))
											else if Table.RowCount(value) = 0 then null 
				else error "Couldn't find single value"
				else  value,
			Result = Table.TransformColumns(table, {{columnName, FindValue}})
		in
			Result,
		
	Table.EnterDataFormula = 
		Document(
			"Table.EnterDataFormula",
			"Convert a table into a formula that's compatible with the Enter Data UI. This function returns an expression you can copy and paste to editthe " &
				" table using Enter Data. Paste the formula, then click the gear icon next to the step name. This is an easy way to allow people to customize a table you've queried",
			{[ 
				Description = "Simple table", 
				Code = "Web.Page(Web.Contents(""https://en.wikipedia.org/wiki/List_of_Super_Bowl_champions"")){3}[Data]", 
				Result = "Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText(""jVVtb9s2EP4rRD4aRqI3S9TH2Gm7ZWkb1MGGom6Aq83IhBkyoOik2a8fT6Ski6Fu+yDdWea9Pffc8du3M342P7uVzrU/jrbZs7UTQgnb+q+lfzL/nFeLxIu0ror55pglWcK8viA6J3o96vVwJksSqvNRT5Ognn2fh1yuQClo2cq8/DCvmMbCPzmmUWaLkEYyn2GkNIhFEFUQvBN1FkQexGI2uP8kXtg73SjQO3YLzkrjMEgRn/NFEmrlpL66JLmnRM+JXhC9IvWR82lxWqvQz8KypTV6a9pYJxZ5nlex1mpMg5dEp98p5OR7TWEmqaYLkgY2eQ2avbegt7LdGlbUof1dApgLz/MASQCcF0EEpH3wDuJiNkfXWUAabT9YITRbwqvHeXsIPovBZ4S5jDWl3kdZjbqHfBbKmXUkGd1iA78ae2AfJGg36ZRH46QzTqKIrtLR1Wc4IA82FzemZZe68cRv2ReQu5Bs3vO/7JOt5kixEt886d45Vp1kg8u/oN1L3Tij2Rexaw9ST3qqAj95FIGmPFC4Jhl+lPAo2ZVRT/vgKeuHoeg9pWT2MqLnhCDkO6ccROSWoJx8NFZsLn7XO48pPBklcQCV6yNmdDJKPrbJj+LIejol9UmUj1Jr0RoH7E/pQWnQc9LPXNJ7ronngEkVyFaVszHh48MDKMOWUqlJN3VCZoDAUxMY6pwkiIASCmwu1u6c3ZijRDY8tpFeHQh5nIUq0r4O3Iotw//WApxTgnm5h5fDhHW3DGdhKDtRBOuMdsNHfhZDy7vCYoEdo8dhzkghePQP0C3uT+le2WovxUOfQfqmieX8Hkeuvh8M/eEtNB5XAXbShoesk3JM93YvFewE0hPYO2g8ejHbjHaExxEkha6kXzdag5M+oG5ATRuS9nF+UukKrGeqBr9ftNuHmT3xgDs6LKZ4BaR0g1wL1xdKAUaK3w+n7+DxCbo1tjxut6CFsFNGuAXehvhslfC9YGuQejIOjslogkv4SgrfgNUebDOWk77hNh1f/OfS+RXmx+o9qK3R0zZ8jHIncA5bIdiddPCL8/VpDCv/Nh5nD/jO460mrPBSD1Hw10qJZ9Fdsf5qexmi4IOuUxzXT5uLy8HgSjhrpGM30vyP09f+PmmNfvYLQLBraI5g/9voN3NscS3fiZ/wbzG+/wM="", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type text) meta [Serialized.Text = true]) in type table [Appearances = _t, Team = _t, Wins = _t, Losses = _t, Winning percentage = _t, Season(s) = _t])"
			]},
			(table as table) =>
				let
					Encoded = Table.JsonEncode(table),
					ColumnMeta= "[" & Text.Range(List.Accumulate(Table.ColumnNames(table), "", (state, current) => state & ", " & current & " = _t"), 2) & "]",
					Text = 
						"Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText(""" & Encoded & """, BinaryEncoding.Base64), Compression.Deflate))," & 
						" let _t = ((type text) meta [Serialized.Text = true]) in type table " & ColumnMeta & ")"
				in
					Text
		),
	
	Table.ExpandRecordColumn = 
		(table as table, columnName as text, optional fieldNames as list, optional newColumnNames as nullable list) => 
		// if fieldNames aren't specified, use the field names from the first row of the column.
		let
			_fieldNames = if fieldNames <> null then fieldNames else List.Buffer(Record.FieldNames(List.First(Table.Column(table, columnName)))),
			_newColumnNames = if newColumnNames <> null then newColumnNames else _fieldNames,
			Result = Table.ExpandRecordColumn(table, columnName, _fieldNames , _newColumnNames)
		in 
			Result,

	Table.FromListCrossJoin = 
		/* 
			Perform a cross join of lists. Example usage:
			Table.FromListCrossJoin({ {ColorsTable[ColorName], "Color"}, {NumbersTable[Number], "Number"}})
			Will give me a new table with two columns, "Color" and "Number" which contains one row for each possible
			combination of colors and numbers
			Table.FromListCrossJoin({{"Red", "Blue"}, "Color"}, {{1,2,3}, "Number"}}) = 
			Table.FromRecords({[Color="Red", Number=1],[Color="Red", Number = 2],[Color="Red", Number = 3],[Color="Blue", Number=1],[Color="Blue", Number=2],[Color="Blue", Number=3]})
		*/
		(listColumnNamePairs as any) =>
			let 
				remainingPairs 	= List.Skip(listColumnNamePairs, 1),
				current 		= List.First(listColumnNamePairs),
				theList 		= List.First(current),
				columnName 		= List.First(List.Skip(current),1),
				firstTable 		= Table.FromList(theList, null, {columnName}),
				doStuff = (table as table, remainingPairs as list) =>
					if List.Count(remainingPairs) <= 0 
					then table 
					else	let 
								current = List.First(remainingPairs),
								theList = List.First(current),
								columnName = List.First(List.Skip(current), 1),
								nextTable = Table.ExpandListColumn(Table.AddColumn(table, columnName, each theList), columnName)
							in 
								@doStuff(nextTable, List.Skip(remainingPairs, 1)),
				Result = doStuff(firstTable, remainingPairs)
			in
				Result,

	Table.JsonDecode = 
		// 
		(encoded as text) =>
		let
			Decompressed = Binary.Decompress(Binary.FromText(encoded, BinaryEncoding.Base64), Compression.Deflate),
			Decoded = Table.FromRows(Json.Document(Decompressed), let _t = ((type text) meta [Serialized.Text = true]) in type table [Name = _t, Number = _t])
		in
			Decoded,
	Table.JsonEncode = 
		// 
		(table as table) =>
		let
			Rows = Table.ToRows(table),
			Json = Json.FromValue(Rows),
			Compressed = Binary.Compress(Json, Compression.Deflate),
			Encoded = Binary.ToText(Compressed, BinaryEncoding.Base64)
		in
			Encoded,
	
	Table.ReplaceValueIf = 
		// Replaces a value if it matches a predicate
		(table as table, replaceIf as function, after as any, columnNameOrList as any) => 
		Table.ReplaceValue(
			table, null, after, 
			(text, old, new)=>
				if replaceIf(text) 
					then new 
					else text, 
			if columnNameOrList is list 
				then columnNameOrList 
				else {columnNameOrList}),

	
	Table.SplitColumnNames = 
		// Splits camelCased and PascalCased column names. 	
		(table as table) => Table.RenameColumns(table, List.Transform(Table.ColumnNames(table), each {_, Text.SplitCamelCase(_)})), 
	
	Table.SplitColumnText = 
		// Splits camelCased and PascalCased text in a column. 
		(table as table, columns as list) => if List.Count(columns) = 0 then table else Table.TransformColumns(@Table.SplitColumnText(table, List.Skip(columns, 1)), {{List.First(columns), Text.SplitCamelCase}}),

	Table.TransformColumn = 
		(table as table, column as text, transform as function) => Table.TransformColumns(table, {{column, transform}}),
	Table.RenameColumn = 
		(table as table, column as text, newName as text) => Table.RenameColumns(table, {{column, newName}}),
	Table.RenameAndTransformColumn = 
		(table, currentName as text, newName as text, transform as function) => Table.TransformColumn(Table.RenameColumns(table, {currentName, newName}), newName, transform),

	//===========================
	// Text
	//===========================
	Text.Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz",
	Text.AlphaNumeric = Text.Alphabet & "0123456789",
	Text.FromList = 
		// 
		(list as list) => List.Accumulate(list, "", (state, current) => state & Text.From(current)),
	Text.IsUpperCase = 
		// Is text all uppercase? returns false if any non-alpha characters are present
		Document(
			"Text.IsUpperCase", 
			"переключает текст",
			{ 
				[ Description = "sdfvsd",   Code = "Switch(1, {1, 2, 3}, {""A"", ""B"", ""C""})",       Result = "A"],
				[ Description = "sdvsdbv",  Code = "Switch(1, {{1, ""A""}, {2, ""B""}, {3, ""C""}})",   Result = "A"]
			},
		(text as text) => 
			List.AllTrue(
				List.Transform(
					Text.ToList(text), 
					(letter)=>Text.Contains(Text.Alphabet, letter) and letter = Text.Upper(letter)
				)
			) 
		),

	Text.IsAlpha = 
		(text as text) => List.MatchesAll(Text.ToList(text), each Text.Contains(Text.Alphabet, _)),
	Text.RemoveExtraWhitespace = 
		(text as text) => Text.Combine(Splitter.SplitTextByWhitespace()(text)," "), 
	Text.SplitCamelCase = 
		// Splits camelCased and PascalCased text and separates by a space. Ex: "thisIsAColumn" -> "this Is A Column"
		(text as nullable text) => if text is null then null else List.Accumulate(Text.ToList(text),"", (state, current) => 
		let
			PreviousLetter = Text.End(state, 1),
			Ignore = (text as text) => text = " " or text = "."
		in 
			state & 
			(if 
				not Text.IsUpperCase(PreviousLetter) and 
				not Ignore(PreviousLetter) and 
				not Ignore(current) and 
				Text.IsUpperCase(current) 
			then 
				" " else "" ) & 
			current),

	Text.SplitOnNotIn = 
		(line as nullable text, validCharacters as text) => Splitter.SplitTextByNotIn(validCharacters)(line),
	Text.SplitOnNonAlpha = 
		// 
		(line as nullable text) =>
		if line is null then null else List.Accumulate(Text.ToList(line), {null} , (state, current) => 
			let
			doSkip = not Text.Contains(Text.Alphabet, current),
			lastItem = List.Last(state),
			appendLast = lastItem<>null
			in
			if doSkip then 
				if lastItem is null then 
				state 
				else 
				List.Combine({state, {null}})
			else
				if appendLast then
				List.Combine({List.RemoveLastN(state, 1), {lastItem & current}})
				else  
				List.Combine({List.RemoveLastN(state, 1), {current}})),


	Text.Substring = 
		// 
		(text as text, start as number, optional count as number) => 
		let 
			start = if start >= 0 then start else error "start index should be >= 0",
			end = if count = null then Text.Length(text) else if count <= Text.Length(text) then count else error "count should be <= text length",
			textList = Text.ToList(text),
			substr = Text.FromList(List.FirstN(List.Skip(textList, start), end - start))
		in 
			substr,

	Text.IsNumber = 
		// 
		(text as text) => try Number.FromText(text) is number otherwise false,

	Text.PositionAfter = 
		// 
		(text as nullable text, substring as text) => 
		let 
			firstIndex = Text.PositionOf(text, substring),
			indexAfter = if firstIndex >=0 then firstIndex + Text.Length(substring) else -1
		in
			if text is null then -1 else if indexAfter >= 0 and indexAfter < Text.Length(text) then indexAfter else -1,
	
	Text.Until = 
		//
		(text as text, endDelimiter as text, optional startIndex as number) => 
		let
			start = if startIndex = null then 0 else startIndex,
			textFromStart = Text.Substring(text, start),
			delimPosition = if Text.PositionOf(textFromStart, endDelimiter) >= 0 then Text.PositionOf(textFromStart, endDelimiter) else Text.Length(textFromStart)
		in
			if text is null then null else Text.Range(textFromStart, 0, delimPosition),


	//===========================
	// Value
	//===========================
	Value.TypeText = 
		//
		(value as any) => 
		if value is binary then "binary" else
		if value is date then "date" else
		if value is datetime then "datetime" else
		if value is datetimezone then "datetimezone" else
		if value is duration then "duration" else
		if value is function then "function" else
		if value is list then "list" else
		if value is logical then "logical" else
		if value is none then "none" else
		if value is null then "null" else
		if value is number then "number" else
		if value is record then "record" else
		if value is table then "table" else
		if value is text then "text" else
		if value is time then "time" else
		if value is type then "type" else
		if value is any then "any"
		else error "unknown -- not a primitive type!"
], // конец библиотеки extensionLibrary

master =[
	readme.master = "БИБЛИОТЕКА master",
	Text.ContainsAny = 
		/*
			Check if a string contains any of the keywords from a given list
			Usage:
				Text.ContainsAny = Load("Text.ContainsAny"),
				Text.ContainsAny("the cat sat on the mat", {"cat", "apple"})
			Result: true
		*/
		(str, needles) as logical =>
		let
			count = List.Count(needles)
		in
			List.AnyTrue(
				List.Generate(
					()=>[i=0],
					each [i] < count,
					each [i=[i]+1],
					each Text.Contains(str,needles{[i]})
				)
			),

	Text.PowerTrim = 
		/*
			Function removes double presents of specified characters.
			By default remove double spaces and leading + ending spaces.
			Like TRIM function in Excel
			
			Original is taked from Ken Puls's blog
			http://www.excelguru.ca/blog/2015/10/08/clean-whitespace-in-powerquery/
		*/
		(text as text, optional char_to_trim as text) =>
			let
				char 		 = if char_to_trim = null then " " else char_to_trim,
				split 		 = Text.Split(text, char),
				removeblanks = List.Select(split, each _ <> ""),
				result		 = Text.Combine(removeblanks, char)
			in
				result,
	
	Text.RemoveSymbols = 
		/*
			Remove all uicode symbols from text
			Originally written by Chris Webb: https://cwebbbi.wordpress.com/2014/08/18/removing-punctuation-from-text-in-power-query/
			Usage:
				Text.RemoveSymbols = Load("Text.RemoveSymbols"),
				newText = Text.RemoveSymbols("a,b,c")
				newText	
			Result: newText = "abc"
		*/
		(inputtext as text) as text =>
		let
			//get a list of lists containing the numbers of Unicode punctuation characters
			numberlists = {{0..31},{33..47},{58..64},{91..96},{123..191}},
			//turn this into a single list
			combinedlist = List.Combine(numberlists),
			//get a list of all the punctuation characters that these numbers represent
			punctuationlist = List.Transform(combinedlist, each Character.FromNumber(_)),
			//some text to test this on
			//inputtext = "Hello! My name is Chris, and I'm hoping that this *cool* post will help you!",
			//the text with punctuation removed
			outputtext = Text.Remove(inputtext, punctuationlist)
		in
			outputtext,
		
	Record.Transform = 
		/*
			Transform a record using a given lambda (passed values k,v)
			Usage:
				let
					Record.Transform = Load("Record.Transform"),
					/* alt: F[Record.Transform]
					Rec = [A=1, B=2]
				in
					Record.Transform(Rec, (k,v) => k & Text.From(v))
				Result: [A="A1", B="B2"]
		*/
		(Rec as record, Lambda as function) as record =>
		let
			Keys 		= Record.FieldNames(Rec),
			Transformed = List.Transform(Keys, each Lambda(_, Record.Field(Rec, _))),
			Recorded 	= Record.FromList(Transformed, Keys),
			Return 		= Recorded
		in
			Return,
		
	Record.Rename = 
		/*
			Rename a record using a given lambda (passed values k,v)
			Usage:
				let
					Record.Rename = Load("Record.Rename"),
					Rec = [A=1, B=2]
				in
					Record.Rename(Rec, (k,v) => k & Text.From(v))
				Result: [A1=1, B2=2]
		*/
		(Rec as record, Lambda as function) as record =>
		let
			Keys = Record.FieldNames(Rec),
			Values = Record.FieldValues(Rec),
			Renamed = List.Transform(Keys, each Lambda(_, Record.Field(Rec, _))),
			Recorded = Record.FromList(Values, Renamed),
			Return = Recorded
		in
			Return			

], // конец библиотеки master

masterPQ = [ // https://github.com/camwally/Power-Query/blob/master/CreateTable2
	// FunctionAlternatives
	List.TransformSaveType = // сохраняет типы на выходе. Надо на примере попробовать
		/*	Эти функции в основном представляют собой версии функций, созданных Microsoft. 
			Однако эти функции помогают сохранить возвращаемые типы функций, которые так легко теряются. 
			Одним из практических результатов использования этих функций в ваших запросах является то, 
			что столбцы вашей таблицы с меньшей вероятностью изменятся на какой-либо тип 
		*/
		(list as list, transform as function)=>
		let
			listTransform = List.Transform(list, transform),
			listType      = type {Type.FunctionReturn(Value.Type(transform))}
		in
			Value.ReplaceType(listTransform,listType),

	//
	ReplaceUsingTable = // 
		(Table as table)=>
		let 
			Rec = Record.FromTable(Table.FromColumns(List.FirstN(Table.ToColumns(Table),2),{"Name","Value"}))
		in 
			(Text as text)=> if Text = null then null else Record.FieldOrDefault(Rec,Text,Text),
	
	// TimeZone
	DateTimeZoneToDateTime_SwitchZone = //
		(TimeZone as number) as function =>
			(DateTimeZone as nullable datetimezone) as nullable datetime=>
			let
				SwitchedZone = DateTimeZone.SwitchZone(DateTimeZone,TimeZone),
				DateTime 	 = DateTime.Date(SwitchedZone) & DateTime.Time(SwitchedZone)
			in
				if DateTimeZone is null then null else DateTime,

	DateTimeZoneToDate_SwitchTimeZone = // ...…te_SwitchTimeZone
		(TimeZone as number) as function =>
			(DateTimeZone as nullable datetimezone) as nullable date=>
			let
				SwitchedZone = DateTimeZone.SwitchZone(DateTimeZone,TimeZone),
				Date = DateTime.Date(SwitchedZone) // & DateTime.Time(DateTimeZone)
			in
				if DateTimeZone is null then null else Date,
	// Table
	Table.FromRecordOfLists = // 
		(rec as record)=>
		let
			tab2col 	  = Record.ToTable(rec),
			valuess_names = List.Reverse(Table.ToColumns(tab2col)),
			tab 		  = Function.Invoke(Table.FromColumns, valuess_names)
		in
			tab,
	
	Table.ToRecordOfLists = //
		(table as table, optional buffer as nullable logical)=>
		let
			tab 	= if buffer = true then Table.Buffer(table) else table,
			Name 	= Table.ColumnNames(tab),
			Value 	= Table.ToColumns(tab),
			Record 	= Record.FromTable(Table.FromColumns({Name, Value},{"Name","Value"}))
		in
			Record,

	Table.Create3 = // из таблицы на входе делает эту же таблицу в коде M
		(inputTable as table, optional #"table type is vertical?" as nullable logical) as text =>
		let
			vert = #"table type is vertical?" ?? false,
			source = Table.Schema(inputTable),
			sortRows = Table.Sort(source,{{"Position", Order.Ascending}}),
			simplifyTypeNameRec = 
				[#"Any.Type" = "any",
				#"Binary.Type" = "binary",
				#"Date.Type" = "date",
				#"DateTime.Type" = "datetime",
				#"DateTimeZone.Type" = "datetimezone",
				#"Duration.Type" = "duration",
				#"Function.Type" = "function",
				#"List.Type" = "list",
				#"Logical.Type" = "logical",
				#"None.Type" = "none",
				#"Null.Type" = "null",
				#"Number.Type" = "number",
				#"Record.Type" = "record",
				#"Table.Type" = "table",
				#"Text.Type" = "text",
				#"Time.Type" = "time",
				#"Type.Type" = "type"],
			simplifyTypeNames = Table.TransformColumns(sortRows,{{"TypeName", each Record.FieldOrDefault(simplifyTypeNameRec,_,_)}}),
			selectColumns = Table.SelectColumns(simplifyTypeNames,{"Name", "TypeName","IsNullable"}),
			addColOfTypeNames = 
				Table.AddColumn(
						selectColumns, 
					"TypeNames", 
					each 
							Expression.Identifier([Name]) 
							& " = " 
							& (if [IsNullable] then "nullable " else "") 
							& [TypeName]
				),
			tableTypeRec = "[" & Text.Combine(addColOfTypeNames[TypeNames], if vert then "#(lf)," else ", ") & "]",

			//Code above is based on a function by Chris Webb

			tableOfExpressions = Table.TransformColumns(inputTable,{},Expression.Constant),
			listOfCols = Table.ToColumns(tableOfExpressions),
			listOfColsOfPaddedText = 
				List.Transform(
						listOfCols,
					(col as list)=>
					let 
							buf = List.Buffer(col),
							maxTextLength = List.Max(List.Transform(buf,Text.Length)),
							paddedText = List.Transform(buf,each Text.PadEnd(_,maxTextLength))
						in
							paddedText
				),
				listOfRowsOfPaddedText = List.Zip(listOfColsOfPaddedText),
				bodyAsList = List.Transform(listOfRowsOfPaddedText, (row as list)=> "{"&Text.Combine(row,",")&"}"),
				bodyAsText = "{"&Text.Combine(bodyAsList,"#(lf),")&"#(lf)}",
				together =  
						"#table(type table"
					& (if vert then "#(lf)" else "")
					& tableTypeRec
					& ",#(lf)"
					& bodyAsText
					& ")"
		in
			together,
	
	Table.RunningTotal = // накопительный итог в новом столбце
		(table as table, colToSum as text, colAsNew as text)=>
		let
			tableB 			= Table.Buffer(table),
			listToSum 		= List.Buffer(Table.Column(tableB,colToSum)),
			listToSumCount 	= List.Count(listToSum),
			runningTotalB 	= 
				List.Buffer(
					List.Skip(
						List.Generate(
							()=> [item = 0, counter = 0],
							each [counter] <= listToSumCount,
							each [item = [item]+listToSum{[counter]}, counter = [counter]+1],
							each [item]
						),
					1)
				),
			addIndex 		= Table.AddIndexColumn(tableB,"IndexUniqueName",0,1),
			addRunningTotal = Table.AddColumn(addIndex,colAsNew,each runningTotalB{[IndexUniqueName]}, type number),
			removeIndex 	= Table.RemoveColumns(addRunningTotal,{"IndexUniqueName"})
		in
			removeIndex,
	
	Table.TransformOtherColumns = // 
		(	Table as table, 
			ColumnsToIgnore as list, 
			Function as function, 
			optional TypeForColumns as nullable type
		) as table =>
		let
			FunctionType 	= Value.Type(Function),
			TypeForFunction = 
				Type.ForFunction(
					[ReturnType = TypeForColumns, Parameters = Type.FunctionParameters(FunctionType)],
				Type.FunctionRequiredParameters(FunctionType)),
		
			FunctionWithNewReturnType = 
				if TypeForColumns is null 
				then Function 
				else Value.ReplaceType(Function,TypeForFunction),  
		
			TransformOtherCols = 
				Table.TransformColumns(
					Table,
					List.Transform(
						ColumnsToIgnore,
						(ColName)=> {ColName, each _}),
				FunctionWithNewReturnType)
		in
			TransformOtherCols,

	Table.AddRandomColumn = // создает дополнительный столбец со случайными числами
		/*	1. 	Add column of random numbers between 0 and 1:
				fAddRandomColumn(Source,"Random Number",0,1)
			
			2. 	Add column of random percentages between 0% and 100%:
				fAddRandomColumn(Source,"Random Percentage",0,1,Percentage.Type)
			
			3. 	Retrieve random item from list:
				fAddRandomColumn(Source,"Random Item",0,List.Count(ListOfItems),{each ListOfItems{Number.RoundDown(_)}, type text})
				(ListOfItems would need to be buffered into memory for it to run quickly, e.g. ListOfItems = List.Buffer(SomeList))
		*/
		(Table as table, 
			RandColName as text, 
			Min as number, 
			Max as number, 
			optional TransformAndTypeCriteria as any) =>
		let
			List = List.Buffer( 
					if not (TransformAndTypeCriteria is list) 
					then {TransformAndTypeCriteria} 
					else	if List.Count(TransformAndTypeCriteria)<=2 
							then TransformAndTypeCriteria 
							else error "Invalid Criteria (Too many items)"
					),
		
			TransformFunc = if List{0} is null or List{0} is type 
							then null 
							else	if List{0} is function 
									then List{0} 
									else error "Invalid Criteria (Tranform)", 

			ListLast = List.Last(List),
		
			Type = 	if	ListLast is null or ListLast is function 
					then type number 
					else  	if ListLast is type 
							then ListLast 
							else error "Invalid Criteria (Type)", 
		
			FirstName = Table.ColumnNames(Table){0},
		
			AddColumn = Table.AddColumn(Table,
										RandColName,
										each Number.RandomBetween(if Record.Field(_,FirstName)=null then Min else Min,Max),
										Type),
		
			Transform = if TransformFunc = null 
						then AddColumn 
						else Table.TransformColumns(AddColumn,{RandColName, TransformFunc})
		in
			Transform,

	Table.AddClusteredIndex = // добавляет столбец с индексом по выбранному столбцу по кластерам
		/* to = Table.AddClusteredIndex(tbl, {"Year"}) // если годы + месяцы, то нумерует месяцы по порядку

		*/
		(Table as table, GroupOnCols as list) => 
		let
			Group   = Table.Group(Table, GroupOnCols, {"ColumnOfTables", each Table.AddIndexColumn(_, "ClustIndex", 0, 1)} ),
			Combine = Table.Combine(Group[ColumnOfTables])
		in
			Combine,
	
	Table.CombineJoins = // объединяет столбцы
		(	LeftTable as table,  LeftKey as list, 
			RightTable as table, RightKey as list, 
			ListOfJoins as list) =>
		let
			Func 		 = each Table.NestedJoin(LeftTable, LeftKey, RightTable, RightKey,"Right",_), 
			ListOfTables = List.Transform(ListOfJoins, Func),
			Combine 	 = Table.Combine(ListOfTables)
		in
			Combine,

	Table.CombineJoinsExpandCoalesce = 
		/* 
		Similar to combine joins, but also: 
			1) Preserves the order
			2) Expands right hand side table
			3) Right key is transferred to left key when left key is null, right key is then deleted 
		*/

		(	LeftTable as table, LeftKeys as list, 
			RightTable as table ,RightKeys as list,
			ListOfJoins as list
		) =>
		let
			Func = each Table.NestedJoin(LeftTable, LeftKeys, RightTable, RightKeys,"Right",_), 
			ListOfTables 	 = List.Transform(ListOfJoins, Func),
			Combine 		 = Table.Combine(ListOfTables),
			OldRightColNames = List.Buffer(Table.ColumnNames(RightTable)),
			NewRightKeys 	 = List.Buffer(List.Transform(RightKeys, each "Right." & _)),
			NewRightColNames = List.ReplaceMatchingItems(OldRightColNames,List.Zip({RightKeys,NewRightKeys})),
			//Below: Using Table.Buffer, so that order is preserved on expand
			ExpandRight 	 = Table.ExpandTableColumn(Table.Buffer(Combine), "Right", OldRightColNames, NewRightColNames),
			Coalesce = 
				List.Accumulate(
					List.Zip({LeftKeys,NewRightKeys}),
					ExpandRight, 
					(Table, Keys) =>Table.ReplaceValue(Table, null, each Record.Field(_,Keys{1}), Replacer.ReplaceValue, {Keys{0}})
				),
			
			RemoveRightKeys 	 = Table.RemoveColumns(Coalesce,NewRightKeys),
			OriginalLeftKeyTypes = List.Transform(LeftKeys, each Type.TableColumn(Value.Type(LeftTable),_)),
			ChangeTypes 		 = Table.TransformColumnTypes(RemoveRightKeys,List.Zip({LeftKeys, OriginalLeftKeyTypes}))
		in
			ChangeTypes,

	Table.NonAggPivotMultRows = // супер. Прокачанный вариант Сведено. Переворачивает строки в столбцы 
		/*	https://www.dingbatdata.com/non-aggregate-pivot-with-multiple-rows-in-powerquery/
			to = PivotMultRows(FlatTable, "Vehicle", "Person")
		*/
		(Source as table,
		ColToPivot as text,
		ColForValues as text)=> 

		let
			PivotColNames 	  = List.Buffer(List.Distinct(Table.Column(Source,ColToPivot))), 
			#"Pivoted Column" = Table.Pivot(Source, PivotColNames, ColToPivot, ColForValues, each _),

			TableFromRecordOfLists = (rec as record, fieldnames as list) =>
				let
					PartialRecord = Record.SelectFields(rec,fieldnames),
					RecordToList = Record.ToList(PartialRecord),
					Table = Table.FromColumns(RecordToList,fieldnames)
				in
					Table,

			#"Added Custom" = Table.AddColumn(#"Pivoted Column", "Values", each TableFromRecordOfLists(_,PivotColNames)),
			#"Removed Other Columns" = Table.RemoveColumns(#"Added Custom",PivotColNames),
			#"Expanded Values" = Table.ExpandTableColumn(#"Removed Other Columns", "Values", PivotColNames)
		in
			#"Expanded Values",

	Table.NonAggPivotMultRows2 = // прокачанный вариант Сведено не проверял
		/*

		*/
		(	Source as table,
			PivotCol as text,
			ValueCol as text
		)=> 
		let
			Source = Table.Buffer(Source), //As source table is referenced 3 times
			GroupClustIndex = 
				Table.Group(Source,
					List.RemoveItems(Table.ColumnNames(Source),{ValueCol}),
					{"ColOfTables", each Table.AddIndexColumn(_,"idx")}
				),

			CombineTables = Table.Combine(GroupClustIndex[ColOfTables]),
			Pivot 		  = Table.Pivot(CombineTables, List.Distinct(Table.Column(Source,PivotCol)), PivotCol, ValueCol),
			RemoveIndex   = Table.RemoveColumns(Pivot,{"idx"})
		in
			RemoveIndex,

	Table.OrderFirstOccurrence = // группировка множественных вхождений элементов массива, упорядоченных по первому вхождению
		/*	Учитывая несортированный массив с повторениями, задача состоит в том, чтобы сгруппировать 
			многократное вхождение отдельных элементов. Группировка должна происходить таким образом, 
			чтобы сохраняться порядок первого появления всех элементов.
		
			Orders table by grouping on a column(s).
			The result is that you order by first occurence.
			https://www.geeksforgeeks.org/group-multiple-occurrence-of-array-elements-ordered-by-first-occurrence 
		*/ 
		(Source as table, GroupBy as list) =>
		let
			#"Grouped Rows" 	= Table.Group(Source, GroupBy, {{"Tables", each _, type table}}),
			#"Removed Columns" 	= Table.RemoveColumns(#"Grouped Rows", GroupBy),
			#"Expanded Tables" 	= Table.ExpandTableColumn(#"Removed Columns", "Tables", Table.ColumnNames(Source)),
			RevertTypes 		= Value.ReplaceType(#"Expanded Tables", Value.Type(Source))
		in
			RevertTypes,
	
	Table.StableSortAndOrderFirstOccurrence = // 
		/*
			It's a bit like Table.Sort, but:
			1) Sorts are stable
			2) If you neglect to put a Order.Ascending or a Order.Descending then ordered by first occurrence 
		*/
		(Table as table, List as list) =>
		let
			fMakeEfficientList = (List as list) =>
				let 
					ColName 	= List.Transform(List, each _{0}),
					Order 		= List.Transform(List, each _{1}?),
					TableForm 	= Table.FromColumns({ColName,Order},{"ColName","Order"}),
					Comparer 	= (a as record, b as record) => Number.From(a[Order] is null or b[Order] is null) ,
					PartionedTable = Table.Group(TableForm, {"Order"}, {{"NestedLists", each Table.ToRows(_), type list}},GroupKind.Local, Comparer),
					PartionedList  = PartionedTable[NestedLists]
				in
					PartionedList,

			fReorder = (Table as table, List as list) =>
				let
					ListItem = List.Buffer(List{0}),
					ListItemColNames = List.Buffer(List.Zip(ListItem){0}),
					GroupedTable = Table.Group(Table, ListItemColNames, {{"NestedTables", each Table.RemoveColumns(_,ListItemColNames), type table}}),
					SortedTable  = if ListItem{0}{1} = null then GroupedTable else Table.Sort(GroupedTable,ListItem),
					fRecursion 	 = each if List.Count(List) > 1 then @fReorder(_, List.Skip(List)) else _,
					RecurseOnNestedTables = Table.Buffer(Table.TransformColumns(SortedTable, {"NestedTables", fRecursion})),
					Expand = Table.ExpandTableColumn(RecurseOnNestedTables, "NestedTables", Table.ColumnNames(RecurseOnNestedTables[NestedTables]{0}))
				in
					Expand,
			
			EfficientList 	= List.Buffer(fMakeEfficientList(List)),
			ReorderRows 	= fReorder(Table,EfficientList),
			ReclaimColOrder = Table.ReorderColumns(ReorderRows,Table.ColumnNames(Table)),
			ReclaimColTypes = Value.ReplaceType(ReclaimColOrder,Value.Type(Table))
		in
			ReclaimColTypes,
	
	Table.SyncTableTypeOfCol = // приписывает тип таблицы столбцу на основе первой таблицы в столбце
		/*
			Приписывает тип таблицы столбцу на основе первой таблицы в столбце как
			Table.TransformColumnTypes не работает со сложными типами
			
			https://social.technet.microsoft.com/Forums/en-US/636e9b44-6820-4ff2-ab60-5dd6a5307bd2/type-conversion-mysteries
		*/
		(PreviousStep as table,ColName as text) =>
			Table.TransformColumns(
				PreviousStep, 
				{ColName, each {_}{0}, Value.Type(Table.Column(PreviousStep,ColName){0})}
			),
	
	Table.SelectRowsWeb = // не пробовал применить
		/*

		*/
		
		(	Table as table,
			fnFilter as function,
			optional Range as nullable list,
			optional Horizontal as nullable logical)=>

		let
			Range      	= if Range      = null then {0}   else Range,
			Horizontal 	= if Horizontal = null then false else Horizontal,
			AddIndex 	= Table.AddIndexColumn(Table, "Index", 0, 1),
			Filter 		= Table.SelectRows(AddIndex, fnFilter),
			Positions 	= List.Buffer(Filter[Index]),
			Values 		= List.Buffer(Table[Column1]),
			
			ListOfTables = 
				List.Transform(Positions, (Pos)=>  
					let  
						Len = List.Count(Range),
						Gen = List.Generate(()=>[i= 0      , Rng = Range{i}, Position = Pos+Rng,  Value = Values{Position}],
											each _[i] < Len,
											each [i= _[i]+1, Rng = Range{i}, Position = Pos+Rng,  Value = Values{Position}],
											each _[[Value],[Rng],[Position]]),
						
						Table = Table.FromRecords(Gen),
						RenameAndChangeTypes = Value.ReplaceType(Table, type table[Value = text, Range = Int64.Type, Position = Int64.Type])
					in
						RenameAndChangeTypes),
			
			Combine 	 = Table.Combine(ListOfTables),
			ValuesToRows = Table.FromRows(List.Transform(ListOfTables, each _[Value]))
		in
			if Horizontal then ValuesToRows else Combine,
	
	Table.TakeTypesAndTryApply = // приписывает типы из одной таблицы и применяет их к одноименным столбцам в другой таблице
		/*
			
		*/
		(tblToAlter as table, tblGet as table) =>
			let
				typetblGet  = Value.Type(tblGet),
				NameCols 	= List.Intersect({Table.ColumnNames(tblGet), Table.ColumnNames(tblToAlter)}),
				NameFuncTypes = 
					List.Transform(
						NameCols, 
						(x) => {x, 
								each {_}{0}, 
								Type.TableColumn(typetblGet, x)}
					),
				Transform = Table.TransformColumns(tblToAlter, NameFuncTypes)
			in
				Transform,
	
	Table.SplitColsDateTime = // наверное правильно соединяет столбцы с датой и временем
		/*
		
		*/
		(	table as table,
			ColName as any, // Can be as list of column names or a single column name 
			optional ColNameDate as nullable text,
			optional ColNameTime as nullable text,
			optional Delimiter   as nullable text
		) as table=> 
		let
			ColNameDateD  = if ColNameDate is null then "Date" else ColNameDate, // Default values can be changed to fit  
			ColNameTimeD  = if ColNameTime is null then "Time" else ColNameTime, // your own language or table naming standards
			DelimiterD    = if Delimiter   is null then "."    else Delimiter,   // (only applicable when ColName is list)
			ColNameIsText = if ColName is text then true else
							if ColName is list then false else
							error "ColName needs to be a text value or a list of text values", 
			fnSplitColumn = 
				(State as table, CurrentColName as text) as table=>
					let
						CurrentColNamesNew = 
							if ColNameIsText 
							then {ColNameDateD, ColNameTimeD}
							else {CurrentColName & DelimiterD & ColNameDateD, CurrentColName & DelimiterD & ColNameTimeD},

						SplitColumn =  Table.SplitColumn(State, CurrentColName, each {_,_}, CurrentColNamesNew),
						
						ColumnTransType = {
							{DateTime.Date, type nullable date},
							{DateTime.Time, type nullable time}
						}, 
						
						TransformColumns = 
							Table.TransformColumns(
								SplitColumn,
								List.Zip({CurrentColNamesNew} & List.Zip(ColumnTransType))
							)
					in
						TransformColumns
		in
			List.Accumulate(
				if ColNameIsText then {ColName} else ColName,
				table,
				fnSplitColumn
			)

],

buch = [
    readme="Небольшая библиотека для получения данных из файлов пакета MS Office. Пользуйтесь! @buchlotnik",
    fxUnzip = 
		[func=(ZIP, optional options)=>
			let encod = [ a = options[Encoding]?, b = if a = null then 866 else a][b],
				compr = (x)=> if x = 0 then Compression.None else Compression.Deflate,

				u16 = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16,ByteOrder.LittleEndian),
				u32 = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32,ByteOrder.LittleEndian),

				get = (x) => BinaryFormat.Record(
					if Binary.Range(x,0,4) = #binary({0x50,0x4b,0x03,0x04})
					then [  Name  = BinaryFormat.Text(u16(Binary.Range(x,26,2)),encod),
							Extr  = BinaryFormat.Binary(u16(Binary.Range(x,28,2))),
							Value = BinaryFormat.Transform(BinaryFormat.Binary(u32(Binary.Range(x,18,4))),(y)=>Binary.Decompress(y,compr(u16(Binary.Range(x,8,2)))))]
					else []
				),

				lst = BinaryFormat.List(BinaryFormat.Choice(BinaryFormat.Binary(30),get),(x)=>x<>[])(ZIP),
				to  = Table.FromRecords(List.RemoveLastN(lst,1),type table [Name=text,Value=binary])
			in 	to,

		typ = type function (
				ZIP as (type binary meta [ Documentation.FieldCaption="ZIP - бинарное содержимое (zip, xlsx, docx, pptx и т.д.)"]), 
				optional options as record) as table meta [ Documentation.Name = "fxUnzip (@buchlotnik)",
						   Documentation.LongDescription ="функция считывает содержимое архива и возвращает таблицу с полями:<p> <b>Name</b> - имя файла (с путём к файлу) <p><b>Value</b> - бинарное содержимое файла. <p>Необязательный аргумент <b>options</b> на текущий момент поддерживает следующие поля: <p>1) <b>Encoding</b> - требуется для корректного чтения не латинских имён файлов (по умолчанию <b>866 - кириллица</b>)"],
		result = Value.ReplaceType(func,typ)
		]  [result],

    
	fxUnzip2016 = 
		[ func = (ZIP, optional cp)=>
			[   ui16 = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger16, ByteOrder.LittleEndian),
				ui32 = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32, ByteOrder.LittleEndian),
				b = (x) => BinaryFormat.Binary(x),
				c = if cp is null then 866 else cp,
				f = (x) => try Binary.Decompress(x, Compression.Deflate) otherwise null,
				g = (x) => BinaryFormat.Transform(b(x),f),
				h = (x) => [head = BinaryFormat.Record([a = b(14), size=ui32, c=b(4), name=ui16, extra=ui16])(x),
							body = BinaryFormat.Record( [ Name   = BinaryFormat.Text(head[name],c),
														Extras = b(head[extra]),
														Value  = g(head[size]),
														flag   = true])][body],
				iftrue   = BinaryFormat.Choice(b(26),h),
				iffalse  = BinaryFormat.Record([flag = false]),
				getfile  = BinaryFormat.Choice(ui32, (x)=> if x = 67324752 then iftrue else iffalse, type binary),
				filelist = BinaryFormat.List(getfile, each [flag] = true)(ZIP),
				remove   = List.RemoveLastN(filelist,1),
				to       = Table.FromRecords(remove, type table [Name = text, Value = binary])
			]  [to],
    
		typ = type function 
				(
					ZIP as (type binary meta [Documentation.FieldCaption = "бинарное содержимое (zip,xlsx,docx,pptx и т.п.)"]),
					optional cp as (type number meta [Documentation.FieldCaption = "кодировка", Documentation.SampleValues = {866}])
                )
                as table meta [ Documentation.Name = "UnZip (@buchlotnik)",
                                Documentation.LongDescription = "Функция считывает содержимое zip-архива и возвращает таблицу с полями FileName - имя файла (с путём, если он в подпапке) и Content - бинарное содержимое конкретного файла. Кодировка требуется для корректного распознавания <b>не латинских</b> имён файлов (для офисных документов указывать необязательно), по умолчанию 866 - кодировка DOS <b>с поддержкой кириллицы</b>"],
		result = Value.ReplaceType(func,typ)
		]  [result],

    
	fxExcelGetRowsAttributes= 
		[func=(filebin,optional options)=>
			let sheetsonly = if options[SheetsOnly]?=true then true else false,
				fxUnzip    = if options[Excel2016]?=true then fxUnzip2016 else fxUnzip,
				bin = Binary.Buffer(filebin),
				xml = List.Buffer(Table.SelectRows(fxUnzip(bin),(r)=>Text.StartsWith(r[Name],"xl/worksheets/sheet"))[Value]),
				wb  = Table.Group(Excel.Workbook(bin,false),"Kind",{"tmp",(x)=>x}),
				xl  = wb{[Kind="Sheet"]}[tmp],
				add = Table.AddIndexColumn(xl,"xml"),
				tr  = Table.TransformColumns(add,{"xml",(x)=>xml{x}}),
				cmb = Table.CombineColumns(tr,{"xml","Data"},fxExcelSheetAddXmlInformation,"Data"),
				to  = if sheetsonly then cmb else cmb & wb{[Kind="DefinedName"]}[tmp]
			in	to,
				
		typ = 	type function(filebin as binary, optional options as record) as table 
				meta [Documentation.Name="fxExcelGetRowsAttributes (@buchlotnik)",
					  Documentation.LongDescription="функция повторят действия функции Excel.Workbook, но к таблице каждого листа добавляет информацию об атрибутах строк"],
		result = Value.ReplaceType(func,typ)
		]  [result],

	
	fxExcelSheetAddXmlInformation= 
		[ func=(x,optional options)=>
			let xml  = Xml.Document(x{0}){0}[Value]{[Name="sheetData"]}[Value][Attributes],
				tr   = List.Transform(xml,Record.FromTable),
				nms  = List.Transform(xml,(x)=>x{[Name="r"]}?[Value]?),
				dict = Record.FromList(tr,nms),
				add  = Table.AddIndexColumn(x{1},"Attributes",Number.From(nms{0})),
				tr1  = Table.TransformColumns(add,{"Attributes",(x)=>Record.FieldOrDefault(dict,Text.From(x))}),
				to   = Table.ReorderColumns(tr1,{"Attributes"}&Table.ColumnNames(x{1}))
			in  to,
		
		  typ  = type function(x as list, optional options as record) as table 
				 meta [ Documentation.Name="fxExcelSheetAddXmlInformation (@buchlotnik)",
						Documentation.LongDescription="вспомогательная функция, добавляющая к таблице листа дополнительную информацию из xml-разметки"],
		  result = Value.ReplaceType(func,typ)
		] [result],

	fxGetMCode=(file)=>
		[ 	u32  = BinaryFormat.ByteOrder(BinaryFormat.UnsignedInteger32,ByteOrder.LittleEndian),
			f=(x)=>[a=Text.Split(x," = "),b={Text.Trim(Text.Replace(a{0},"shared","")),Text.Combine(List.Skip(a)," = ")}][b],
			from = fxUnzip(file){[Name="customXml/item1.xml"]}?[Value]?,
			bin  = Binary.FromText(Xml.Document(from){0}[Value],BinaryEncoding.Base64),
			data = BinaryFormat.Choice(BinaryFormat.Binary(8),(x)=>BinaryFormat.Binary(u32(Binary.Range(x,4,4))))(bin),
			unz  = fxUnzip(data){[Name="Formulas/Section1.m"]}[Value],
			txt  = Text.FromBinary(unz)&"#(cr,lf)",
			splt = Text.Split(txt,";#(cr,lf)"),
			lst  = List.Range(splt,1,List.Count(splt)-2),
			tbl  = Table.FromList(lst,f,{"Name","Value"}),
			to   = if from=null then #table({"Name","Value"},{{null,null}}) else tbl
		][to],


	fxTableRemoveEmptyColumns= 
		[ func = (table, optional options) =>
			let	tbl = Table.Buffer(Table.Profile(table)),
				lst = Table.SelectRows(tbl,(x)=>x[Count]=x[NullCount])[Column],
				to  = Table.RemoveColumns(table,lst)
			in  to,
		  
		  typ = type function (table as table, optional options as record) as table 
			    meta [ Documentation.Name            = "fxTableRemoveEmptyColumns (@buchlotnik)",
				  	   Documentation.LongDescription = "функция, удаляющая пустые столбцы из таблицы"],
		  result = Value.ReplaceType(func,typ)
		][result],	

	TableRenameColumnsByPositions = 
		[func =(table,list) =>
			[   a = List.Buffer(Table.ColumnNames(table)),
				b = List.Count(a),
				c = (x) =>{a{if x{0}>0 then x{0}-1 else b+x{0}},x{1}},
				d = if list{0} is list then List.Transform(list,c) else c(list),
				e = Table.RenameColumns(table,d)
			]  [e],
		 
		 typ = type function (
							table as (type table meta [Documentation.FieldCaption = "исходная таблица"]),
							list as (type list meta [Documentation.FieldCaption = "параметры переименования"])
						)           
						as table 
							meta [ 	Documentation.Name = "TableRenameColumnsByPositions> (@buchlotnik)",
									Documentation.LongDescription = "Функция переименовывает столбцы таблицы  по их номеру. Параметры переименования задаются в виде списка {номер, новое название} либо списка списков {{номер1, новое название1},{номер2, новое название2}}. Положительное значение номера задает нумерацию с начала таблицы, отрицательное - с конца.",
									Documentation.Examples = 
									{
										[Description = "переименовать первый столбец", Code = "=TableRenameColumnsByPositions(#table({""a"",""b"",""c""},{}),{1,""первый""})",Result="#table({""первый"",""b"",""c""},{})"], 
										[Description = "перименовать первый с конца столбец", Code =  "=TableRenameColumnsByPositions(#table({""a"",""b"",""c""},{}),{-1,""первый с конца""})",Result="#table({""a"",""b"",""первый с конца""},{})"], 
										[Description = "переименовать первый и последний столбцы", Code = "=TableRenameColumnsByPositions(#table({""a"",""b"",""c""},{}),{{1,""первый""},{-1,""последний""}})",Result="#table({""первый"",""b"",""последний""},{})"]
									}       
								],
		 result = Value.ReplaceType(func,typ)
		][result]

],
	// выводим только список функций, без других утилит
	result = 
		UTIL & 
		libPQ & 
		extensionLibrary & 
		master &
		buch &
		TABLE
 in 
 	result 