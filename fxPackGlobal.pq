let
	//====================================================================================================================================
	Biblio = 
		[ readme = "Библиотека функций для проекта Анализ Отчетности v 1.1",
		
		/*
			from – на вход прилетела таблица, здесь не очень большая, просто чтоб пост не загромождать, но и дат может быть больше и всякие цены с НДС – короче включите фантазию
			nms  – получили список имён всех столбцов
			dlst – из общего списка получаем только имена столбцов с датами
			nlst – также из общего списка получаем имена столбцов с числовыми значениями (сам приём обсуждался тут (https://t.me/pbi_pq_from_tank/140)  и основываясь на этом принципе вы можете адаптировать код под ваши ключевые слова)
			tlst – ну и все столбцы не с датами и не с числами мы хотим сделать текстовыми – тут нам в помощь List.Difference - из всех имён оставили только те, которые не вошли в dlst и nlst.
			Функция f – она  принимает два аргумента: x – список, y – тип, и преобразует наш список в список списков вида {название,тип} (это мы тоже на самом деле уже обсуждали (https://t.me/pbi_pq_from_tank/98) ) 
			transflst – на основе наших списков с помощью вспомогательной функции собираем списки списков преобразований и объединяем их (как работает амперсанд  со списком списков обсуждали тут (https://t.me/pbi_pq_from_tank/150)) 
			to – применили наш список преобразований типов к таблице – и всё!
		*/
		
		// ---------------------------------------------------------------------------------------	
		// fxParseHead - сворачивание многоуровневых заголовков
		fxParseHead = (
			table,				//
			num,				//
			optional delim,		//
			optional fill		//
		) =>
				let lst   = List.Buffer(
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
				in  out,


		// ---------------------------------------------------------------------------------------	
		// Загрузка функций с Гитхаба / локальной папки
		Load = (fnName as text, optional BasePath as text) as function => // as function  к чему обязывает?!
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
				Function,


		// ---------------------------------------------------------------------------------------
		// Справка по функциям. txt = часть названия. Например =fxFAQ("list") 
		fxFAQ = (txt as text) =>
			let	list_1 = Record.ToTable(#shared), 
				rows_select = Table.SelectRows( list_1, each Text.Contains(Text.Lower([Name]), txt) and Text.Contains([Name], ".") ), 
				tab_add_col = Table.AddColumn( rows_select, "Meta", each Value.Metadata(Value.Type([Value])) ), 
				cols_select = Table.SelectColumns(tab_add_col, {"Meta"}), 
				col_expand_1 = Table.ExpandRecordColumn(
					cols_select, 
					"Meta", 
					{
						"Documentation.Name", 
						"Documentation.Description", 
						"Documentation.LongDescription", 
						"Documentation.Examples"
					}, 
					{"Name", "Description", "LongDescription", "Examples"}
				), 
				col_expand_2 = Table.ExpandListColumn(col_expand_1, "Examples"), 
				
				col_expand_3 = Table.ExpandRecordColumn(
					col_expand_2, 
					"Examples", 
					{"Description", "Code", "Result"}, 
					{"Example Description", "Example Code", " Example Result"}
				),
					Sort = Table.Sort(col_expand_3,{{"Name", Order.Ascending}}),
					Filtr = Table.SelectRows(Sort, each ([Name] <> null))
			in  Filtr,
		

		// ---------------------------------------------------------------------------------------
		// Создание нескольких столбцов
		fxAddMultiColumns = 
			[func = (table as table, addedColumns as list) as table =>
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

			
		// ---------------------------------------------------------------------------------------
		/*	ФУНКЦИЯ: fxClear - Чистка текста */
		fxClear = 
		[ 	func = (
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
					meta [ Documentation.Name            = "fxClear",
						Documentation.LongDescription = "функция, создающая сразу несколько столбцов"],

			result = Value.ReplaceType(func,typ)
		] [result],
		
		
		// ---------------------------------------------------------------------------------------
		/*  FnReorderSubsetOfColumns - Переупорядочиваение столбцов без привязки к названию других
			Функция получает:
				tbl - таблицу, 
				reorderedColumns - подмножество переупорядоченных имен столбцов в виде списка и 
				offset - индекс (номер колонки куда вставлять полонки по списку), начинающийся с нуля. 
				Затем выполняется переупорядочение очень надежным способом, без ссылок на имена других столбцов.

			Пример = FnReorderSubsetOfColumns(Буфер, {"Строка", "Код"}, 0)
		*/
		fxReorderSubsetOfColumns = (tbl as table, 
									reorderedColumns as list, 
									offset as number) as table =>
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
		

		// ---------------------------------------------------------------------------------------
		// Удаление пустых столбцов
		fxTableRemoveEmptyColumns= [ 
			func = (table, optional options) =>
				let	tbl = Table.Buffer(Table.Profile(table)),
					lst = Table.SelectRows(tbl,(x)=>x[Count]=x[NullCount])[Column],
					to  = Table.RemoveColumns(table,lst)
				in  to,
			
			typ = type function (table as table, optional options as record) as table 
					meta [ Documentation.Name         = "fxTableRemoveEmptyColumns (@buchlotnik)",
						Documentation.LongDescription = "функция, удаляющая пустые столбцы из таблицы"],
			
			result = Value.ReplaceType(func,typ)
		][result],


		// ---------------------------------------------------------------------------------------
		// Удаление пустых строк
		fxTableRemoveEmptyRows= 1,

		

		// ---------------------------------------------------------------------------------------
		/* 	fnCompareGroups - проверяет, если ли тексте символы из списка
			
			Например: 
			1. 	Создаем нужный список (группа) ключей и проверяем вот таким образом и отбираем строки:
				to = Table.SelectRows(Таблица, each fnCompareGroups(Ключи1, [Текст]) 
			2.  
		*/
		fxCompareGroups=(keys as list, text as text) as logical =>
			let Source = List.AnyTrue( List.Transform( keys, (key)=> Text.Contains(text, key, Comparer.OrdinalIgnoreCase) ) )
			in 	Source,
		
		
		// ---------------------------------------------------------------------------------------
		// ФУНКЦИЯ: TableRenameColumnsByPositions - пример?
		TableRenameColumnsByPositions = 
			(table,list) =>
				[   a = List.Buffer(Table.ColumnNames(table)),
					b = List.Count(a),
					c = (x) =>{a{if x{0}>0 then x{0}-1 else b+x{0}},x{1}},
					d = if list{0} is list then List.Transform(list,c) else c(list),
					e = Table.RenameColumns(table,d)
				]  [e],

		
		// ---------------------------------------------------------------------------------------
		// Переименовать все столбцы по новому списку
		fxRenameColumns = 
			(table,listOld,listNew) => 
				Table.RenameColumns(table, List.Zip({
					List.Buffer( listOld ), 
					listNew})
				),

		// Номер региона
		fxРегион = (КодРегиона as text) => 
			Record.FieldOrDefault(
				[1 = "Брест", 2 = "Витебск", 3 = "Гомель", 4 = "Гродно", 5 = "Минск", 6 = "Могилев", 0 = "БКС"], 
				КодРегиона, 
				null
			),



	// =============================================================================================================================================
	// ШАБЛОН ОФОРМЛЕНИЯ ФУНКЦИИ
	// =============================================================================================================================================
		Readme.Template = "Шаблоны написания функций",
		
		// --------------------------- Function ------------------------------
		Template = 
		let
			// --------------------------- Fucntion segment -----------------------------------
			output =
				(/*parameter as text, optional opt_parameter as text*/) /*as text*/ =>      	// Input definition + Function output type definition
					let                                                                         // Inner function steps declaration
						initStep = "",
						lastStep = ""
					in
						lastStep,                                                               // Output from inner steps
			// --------------------------- Documentation segment ------------------------------
			documentation = [
				Documentation.Name = " NAME OF FUNCTION ",                                      // Name of the function
				Documentation.Description = " DESCRIPTION ",                                    // Decription of the function
				Documentation.Source = " URL / SOURCE DESCRIPTION ",                            // Source of the function
				Documentation.Version = " VERSION ",                                            // Version of the function
				Documentation.Author = " AUTHOR ",                                              // Author of the function
				Documentation.Examples =                                                        // Examples of the functions
				{
					[
						Description = " EXAMPLE DESCRIPTION ",                                  // Description of the example
						Code = " EXAMPLE CODE ",                                                // Code of the example
						Result = " EXAMPLE RESULT "                                             // Result of the example
					]
				}
			]
			// --------------------------- Output --------------------------------------------
		in
			Value.ReplaceType(                                                                  // Replace type of the value
				output,                                                                         // Function caller
				Value.ReplaceMetadata(                                                          // Replace metadata of the function
					Value.Type(output),                                                         // Return output type of function
					documentation                                                               // Documentation assigment
				)
			), 

		

		// ---------- ВАРИАНТ 1. Простая сложность ----------
		// Важно!	Тип функции, типы и названия аргументов в func только для самой функциии
		//			Тип функции, типы и названия аргументов в fnType для переопределения. На выход имет именно это, благодаря Value.ReplaceType
		// 			meta - добавляется к ЗНАЧЕНИЮ и является Записью/Record
		function1 = 
			[ 	func = 
					(txt as text, optional options) =>
						let	from = Text.Length(txt)
						in  from,
				
				fnType = type function (txt as text, optional options as text) as any
			//  fnType = type function (#"Число Символов" as text, optional Параметр as text) as any
					meta [ 	Documentation.Name            = "function1 (@Medvedev)",
							Documentation.LongDescription = "Функция, определяющая количество символов в строке <br>
															Имя аргумента в fnType = имя аргумента в заголовке поля для ввода: <b> tbl, options</b><br>
															Можно менять тут, а можно в <b>meta</b> по каждому аргументу/значению <br>
															В последнем случае можно добавить много чего еще, например пример отображения"],
						/*  Documentation.LongDescription = "<b> Грузим из папки код запросов </b> <br>" & Text.Combine(from[Local],"<br>") */ 

				result = Value.ReplaceType(func, fnType) // меняем тип исходной функции и ее аргументов
			]  [result],
		


		// ---------- ВАРИАНТ 2. Средняя сложность ----------
		/* добавляем описание каждого аргумента для вывода в окно ввода значений аргументов функции */
		function2 =	 let 
			// ШАГ 1. формируем тело функции
			func = 
				(message as text, optional count as number) as list =>
					let from 	 = Text.Split(message," ") meta [Поле1 = "раз", Поле2 = "два"],
						metadata = Value.Metadata(from),
						to 		 = from & {metadata[Поле1]}
					in  to, 	
				/* in  to meta metadata,*/
			
			// ШАГ 2. формируем тип функции, типы аргументов и добавляем для них метаданные
			fnType = 
				type function (
					message as (
						type text meta [
							Documentation.FieldCaption 		= "Сообщение",
							Documentation.FieldDescription 	= "Text to display", 					// не отображается нигде
							Documentation.SampleValues 		= {"Привет мир", "Мама мыла раму"}
						]
					),
					optional count as (
						type number meta [
							Documentation.FieldCaption 		= "Количество",
							Documentation.FieldDescription 	= "Количество пропускаемых пробелов",	// не отображается нигде
							Documentation.AllowedValues 	= {1, 2, 3}
						]
					)
				) as list
					meta [
						Documentation.Name 					= "Name - Text.SplitSepar",
						Documentation.LongDescription 		= "Long Description - разделитель текста по пробелу",
						Documentation.Examples = {
							[
								Description = "Возвращяет список слов исходной строки разделенной по раздлелителю",
								Code 		= "Text.SplitSepar(""Привет мир"", 2)",
								Result 		= "#table({""Column1""}, {{""Привет мир""}, {""Привет мир""}})"
							],
							[
								Description = "Другой пример, новое сообщение, новое количество!",
								Code 		= "Text.SplitSepar(""Goodbye"", 1)",
								Result 		= "#table({""Column1""}, {{""Goodbye""}})"
							]
						}
				],
			
				// ШАГ 3. Собираем функцию
				result = Value.ReplaceType(func, fnType)
			in 
				result,
		


		// ---------- ВАРИАНТ 3. Сложный ----------
		/* Разнообразим типы аргументов + добавляем аргументы в группу */
		function3 =
			let 
				// ШАГ1. Тело функции
				func = (
					tbl as table, 		// ...	
					optional path,		// ...
					optional endDate	// ...
				)=>
			
				let from 	 = Table.FromColumns({{1..31}},{"День месяца"}),
					to 		 = Table.ToList(tbl)
				in  to ,
						
			// ШАГ 2. Документирование аргументов функции
			fnType = type function (
				tbl as (type table 
					meta [
						Documentation.FieldCaption = "бинарное содержимое (zip,xlsx,docx,pptx и т.п.)",
						Documentation.SampleValues = {866}
					]
				),
				
				optional path as (type text 
					meta [
						Documentation.FieldCaption = "Название отчета:",
						Documentation.FieldDescription = "Придумайте название отчёта, чтобы потом загрузить его по имени",
						Documentation.SampleValues = {"Платон Щукин"}
					]
				),
				optional fields as (type {text}
					meta [
						Documentation.FieldCaption = "Поля отчета (нужные отметьте галочками):",
						Documentation.FieldDescription = "Выберите допустимые поля для отчета #(cr) Можно передавать в виде списка или в строку с разделителями",
						Documentation.AllowedValues = List.Sort( Table.FromColumns({{1..31}},{"ДеньМесяца"})[ ДеньМесяца ] )
					]
				) 

			// возвращяемый тип функции и ее документирование
			) as number 
					meta [ 
						Documentation.Name = "Функция JuPQ",
						Documentation.LongDescription = "Функция считывает и возвращает таблицу с полями FileName - имя файла (с путём, если он в подпапке) и Content - бинарное содержимое конкретного файла. Кодировка требуется для корректного распознавания <b>не латинских</b> имён файлов, по умолчанию 866 - кодировка DOS <b>с поддержкой кириллицы</b>"
					],
			
			// ШАГ 3. Собираем функцию
			result = Value.ReplaceType(func, fnType)
		in 
			result,



		// ---------- ВАРИАНТ 4. Сложный ----------
		/* Разнообразим типы аргументов + добавляем аргументы в группу + выбираем вид аргументов
		То есть, можно сделать исходнуюфункцию с одним любым аргументом, а потом переназначить аргументы на нужные (можно переключать)*/
		function4 =
			let
				func = (input as record) => Record.ToTable(input),
			
				fnType = type function(
					beginDate as (type date meta [
							Documentation.FieldCaption     = "Дата начала интервала:++",
							Documentation.FieldDescription = "Дата начала интервала не может быть позднее чем сегодня",
							Documentation.SampleValues     = {DateTime.Date(DateTime.LocalNow() - 1)}
						]
					)
				) as text,

				NewType = type function (
					inputSection1 as
					[
						server = (type text 
							meta [
								Documentation.FieldCaption	= "Введите дату начала интервала:",
								Documentation.SampleValues 	= {"поле для записи"}
							]
						), 
						optional Timeout = (type number 
							meta [
								Documentation.FieldCaption 	= "Укажите время ожидания в секеундах",
								Documentation.AllowedValues = { 10, 60, 360 }
							]
						),
						optional CutOff = (type date
							meta [
								Documentation.FieldCaption 	= "Введите дату",
								Documentation.SampleValues 	= {"01/01/2024"}
							]
						),
						optional fields = (type {text}
							meta [
								Documentation.FieldCaption 		= "Поля отчета (нужные отметьте галочками):",
								Documentation.FieldDescription 	= "Выберите допустимые поля ",
								Documentation.AllowedValues 	= List.Sort( Table.FromColumns({{1..7}},{"ДеньНедели"})[ДеньНедели] )
							]
						),
						optional endDate = (type date 
							meta [
								Documentation.FieldCaption 		= "Дата окончания интервала:",
								Documentation.FieldDescription 	= "Дата окончания должна совпадать с датой начала интервала или быть раньше нее",
								Documentation.SampleValues 		= {DateTime.Date(DateTime.LocalNow() - 1)}
							]
						)
					]
				) as any
					meta [ 	Documentation.Name            = "function2 (@Medvedev)",
							Documentation.LongDescription = "Сложный типа функцииФункция, в которой на входе <b> record </b> <br>
															а на выходе значение введенных аргументов в виде таблицы со списком значений введенных аргументов"
					]
				,
			
				result = Value.ReplaceType(func, NewType)
			in
				result,

			
			Readme.Tempate2= "Функция простого документирования других функций",
			
			// Функция документирования других функций
			// по простому, буз именения типов, поле и прочего. Только Имя, Описание и Пример
			Document = 
				(	name as text, 
					description as text, 
					valueOrExample as any, 
					optional valueIfExample as any) =>
				
					let
						value 	 = if valueIfExample is null then valueOrExample else valueIfExample,
						examples = if valueIfExample is null then {} else valueOrExample
					in
						Value.ReplaceType(value, Value.Type(value) 
							meta [
								Documentation.Name = name,
								Documentation.Description = description,
								// [Description = "", Code="", Result =""]
								Documentation.Examples = examples
							]
				),
			
			// пример использования
			Text.Lenght =
				Document(
					"Text.Lenght",
					"Возвращает число символов в строке",
					{[  Description = "Считаем число символов, не цифр", 
						Code 		= "func[Text.Lenght])(""мама мыла раму"")", 
						Result 		= "{""14""}"
					]},
					(txt as text) => Text.Length(txt)
			)
	],
	

	// =============================================================================================================================================
	// БЛОК тестирования функций 
	// =============================================================================================================================================

	Readme.Testint = "Функции для тестирования функций",
	MissingTests         = List.Select(Record.FieldNames(Biblio), each not List.MatchesAny(Tests, (testRecord as record) => Text.Contains(testRecord[CaseName], _))),
	MissingDocumentation = List.Select(Record.FieldNames(Biblio), each not Record.HasFields(Value.Metadata(Value.Type(Record.Field(Biblio , _))), "Documentation.Examples")),

	// утилита проверки функций
	TestUtils = [ 
		AssertEqual = (expected as any, actual as any, description as text) => 
			let
				expectedType = Value.Type(expected),
				actualType   = Value.Type(actual),
				listAsText   = (list as list) => "{" & Text.Range(List.Accumulate(expected, "", (state, current) => state & ", " & Text.From(current)), 1) & " }",
				expectedAsText = if expected is list then listAsText(expected) else Text.From(expected),
				actualAsText = if expected is list then listAsText(actual) else Text.From(actual),
				typeAsText   = (value as any) => 
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
					else "unknown -- not a primitive type!"
			in
				if 		not Value.Is(actualType, Value.Type(expectedType)) 
				then  	error "Expected type " & typeAsText(expectedType) & " does not match actual type " & typeAsText(actualType) 
				else 	if expected = actual 
					 	then true
						else if expectedAsText = actualAsText 
							then error "Cannot provide accurate failure message! Text versions of expected and actual are identical, but they failed the equality test!"
							else error "Expected: " & expectedAsText & "; Actual: " & actualAsText & "; Reason: " & description,
		SimpleTable = Table.FromRecords({[TextCol = "A", NumberCol = "1"], [TextCol = "B", NumberCol = 2], [TextCol = "C", NumberCol = 3]})
	],
		
	// утилита выворда результатов проверки функций
	TestResults = 
		let failedTests = List.Select(
				List.Transform(Tests, (suite as record) => 
					let testResult = try suite[Test](Biblio)
					in	if testResult[HasError] then Error.Record(suite[CaseName], testResult[Error][Message], null) else true), 
				each _ <> true) 
		in 	if List.IsEmpty(failedTests) 
			then "All " & Text.From(List.Count(Tests)) & " tests passed! :)" 
			else failedTests,
		
	// сама проверка функций
	Tests = {
		[ 	CaseName = "List.From", 
			Test = (Biblio as record) => 
				TestUtils[AssertEqual]
				(
					{"A", "B", "C"}, 
					Biblio[Text.Lenght]("мама мыла раму"),
					"Text should be split on ,"
				)
		]/*,
		[	CaseName = "List.From - no braces", 
			Test = "цвм" ]*/
	},

	// выводим только список функций, без других утилит
	result = Biblio 
 in result 