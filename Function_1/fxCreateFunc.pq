// ШАБЛОН ОФОРМЛЕНИЯ ФУНКЦИИ
let
	
	// есть ссылка в других функциях, из за этого ошибка вылазит
	Biblio = 
		[ Readme.Template = "Шаблоны написания функций",
		
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
				Documentation.Source   = " URL / SOURCE DESCRIPTION ",                          // Source of the function
				Documentation.Version  = " VERSION ",                                           // Version of the function
				Documentation.Author   = " AUTHOR ",                                            // Author of the function
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
		/* 	Разнообразим типы аргументов + добавляем аргументы в группу + выбираем вид аргументов
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
				result
	],
			
		Readme.Tempate2= "Функция простого документирования других функций",
			
		// =============================================================================================================================================
		// Функция документирования других функций
		// по простому, буз именения типов, поле и прочего. Только Имя, Описание и Пример
		// =============================================================================================================================================
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
				),
	

	// =============================================================================================================================================
	// БЛОК тестирования функций Должен идти отдельно от библиотеки
	// =============================================================================================================================================

	Readme.Testint = "Функции для тестирования функций",
	// MissingTests         = List.Select(Record.FieldNames(Biblio), each not List.MatchesAny(Tests, (testRecord as record) => Text.Contains(testRecord[CaseName], _))),
	// MissingDocumentation = List.Select(Record.FieldNames(Biblio), each not Record.HasFields(Value.Metadata(Value.Type(Record.Field(Biblio , _))), "Documentation.Examples")),

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
		
	// утилита вывода результата проверки функций
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
in 
	result