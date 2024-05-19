let 
func = [
	// меняет первую колонку на ключ из нескольких колонок ID
	key = 	
		(tbl as table) as table=>
			Table.FromList( // пересобираем всю таблицу с новой колонкой
        		Table.ToRows(tbl), // данные
        		(x)=>{Text.Combine(List.Transform(F[LIST][List.Items](x,{9,10,11,12}),Text.From),".")} & List.Skip(x), // функция преобразования
        		{"key"} & List.Skip(Table.ColumnNames(tbl))
			), // колонки
			 
	fxSelectRowFromKod = // фильтрация строк по коду строки. Надо документировать функцию с описанием аргументов
		/* Пример.
			= fxPackForm[fxSelectRowFromKod](STEP1_Connect,
				{"Column1","Column2","Column3","Column4"},
				{"Статья", "Строка_ID", "На начало", "На конец"},
				"Строка_ID", 10, 1000
				)
		*/
		(	tbl as table, 
			nmsOld as list,		//
			nmsNew as list,		//
			columnName as text, //
			kodStart as number, //	
			kodEnd as number	//
		) as table=>
			let a =	Table.AddColumn(tbl, "Step2", (x)=>
						let selCol  = Table.SelectColumns(x[Content], nmsOld), 
							rename  = Table.RenameColumns(selCol, List.Zip({nmsOld, nmsNew} )),
							clear   = F[Table.ClearAll](rename," #(lf)#(00A0)",true,true),
							tr      = Table.TransformColumns(clear, {columnName, try Number.From otherwise null, Int64.Type}),
							err     = Table.RemoveRowsWithErrors(tr, {columnName}), // удаляем errors, который иногда появляется, когда там текст или еще чего
							lst     = {null, kodStart, kodEnd},
							f=(x)=> x>=lst{1} and x<=lst{2},
							selRow  = Table.SelectRows(err, (row)=> f(Record.Field(row, columnName) ) )
						in	selRow
					),
				to	= Table.RemoveColumns(a, {"Content"})
			in 	to
],

Справочники = [
	readme = "Справочники проекта v 1.0",
	BKS = [
		Спр_Предприятия =
			let
				Источник = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("jVPbasJAFPyVkGcFL7HFb7H+iwbEFqQl9qElFYMR2qeCRoO3Jv7C7B91zkbNJljqk2F3ZnZmzrHTset2xYaneohVX7mqjz222FmYIsCMd3W7W+nYDUGNsVUuYixLqIr8evjg9/TCaArjFSl+yDio0S0cR3PoJcUaCXmJZqWIr6Fbgp5QM7lF+y5Dp1gReaB29GfWmkDnCKwHNiPutYcUR8miBg827xsZVDv+OoN9LAimIU2IWNcwD7BBVOVZalHTx8xUuXg76OSxGpbTk8xjFqMGPNioF+xkXKKXy9wXzASXZv57vp3xgjxzyqe25KZmSz5CTPR1kd9sFt59q2JtlXo76pmuTlvWP4vkEo5ROY2rJxkQKXx2r3doqWtdKLckLSbVM9MtTEdtQ+6dLbi6vOC0WaH4wA5rg+LUCiGK9iMZh3qkjBSzMmn1jBbm7QUaxBmqkfRebssprswYn7ckck6JpsK5+q/y6OEbc4PTqhkcj931uJ5JYTE1M4R/ZnV/AQ==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [ID_Предприятие = _t, Название_исх = _t, Группа_1 = _t]),
				#"Измененный тип" = Table.TransformColumnTypes(Источник,{{"ID_Предприятие", Int64.Type}, {"Группа_1", Int64.Type}}),
				#"Переименованные столбцы" = Table.RenameColumns(#"Измененный тип",{{"Название_исх", "Предприятие_НазваниеИсх"}, {"Группа_1", "Предприятие_Группа1"}})
			in #"Переименованные столбцы",
		
		Спр_ГруппаПредприятий =
			let
				Источник = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WMlTSUbqw5MKOC7uBeMOF3Re7lWJ1opWMQMKLL8xXio0FAA==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Group_ID = _t, НазваниеГруппы = _t]),
				#"Измененный тип" = Table.TransformColumnTypes(Источник,{{"Group_ID", Int64.Type}, {"НазваниеГруппы", type text}})
			in
				#"Измененный тип",
		
		Спр_Предприятия2 = // полный список, копируется их файла Excel
			let
				Источник = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("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", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Код = _t, Название_исходное = _t, Группа = _t, Область = _t]),
				#"Измененный тип" = Table.TransformColumnTypes(Источник,{{"Код", Int64.Type}, {"Название_исходное", type text}, {"Группа", Int64.Type}, {"Область", Int64.Type}})
			in
				#"Измененный тип"

	],
		
	Общие  = [ // Формы отчетности и др.
		Спр_ФормыОтчетности = [ // Справочники форм отчетности 
			Источник = #table(type table 
				[   КодБКС          = text, 
					Report_ID       = text, 
					Название        = text, 
					КраткоеНазвание = text, 
					Примечание      = text   ],
				{   {"0112",    "balance",  "Баланс",                                               "ББ",   ""}, 
					{"0212",    "opu",      "Отчет и прибылях и убытках",                           "ОПУ",  ""},
					{"0312",    "ved_odr",  "Отчет о доходах, расходах и финансовых результатах",   "ОДРФР",""},
					{"0412",    "odds",     "Отчет о движении денежных средств",                    "ОДДС", ""},
					{"0512",    "pril",     "Приложение к бухгалтерскому балансу",                  "ПББ",  ""},
					{"3012",    "oik",      "Отчет об изменении капитала",                          "ОИК",  ""},
					{"0612",    "osr1",     "Отчет о состоянии расчетов",                           "ОСР",  ""},
					{"06R1",    "osr2",     "Отчет о состоянии расчетов new",                       "ОСР",  ""},
					{"1212",    "ofr1",     "Отчет о финансовых результатах",                       "ОФР",  ""},
					{"12f1",    "ofr2",     "Отчет о финансовых результатах",                       "ОФР",  ""}  
				}
			)
			] [Источник],

		Спр_Год =  // Справочники периодов
			Table.FromList({2017..2024}, Splitter.SplitByNothing(), type table [Год = number]),
				
		Спр_Область = // Справочники областных городов
			let Источник = Table.Buffer(#table(type table [
						Область_ID 	= number, 
						Область 	= text	],
					{
						{1,"Брест"}, 
						{2,"Витебск"},
						{3,"Гомель"},
						{4,"Гродно"},
						{5,"Минск"},
						{6,"Могилев"} 
					}))
			in Источник
		],
		
	БАЛАНС = [ // Бухгалтерский баланс
		Спр_ББ_Часть   = // Справочники частей 
			let Источник = #table(type table 
					[   
						Часть_ID = number, 
						ЧастьИмяПолное= text, 
						ЧастьИмяКраткое = text, 
						ЧастьАббревиутура = text
					],
					{  
						{1,	"Активы",								"Актив",	"А"}, 
						{2,	"Собственный капитал и обязательства", 	"Пассив",	"П"}
					}
				) 
			in 	Источник,
			
		Спр_ББ_Раздел  = // Справочники разделов 
			let	Источник = #table(type table 
					[
						Раздел_ID = number, 
						РазделПрефикс= text, 
						РазделИмя = text, 
						РазделИмяПолное = text,
						РазделИмяСокращенное = text, 
						РазделИмяАббревиатура = text
					],
					{
						{1, "Раздел 1", "Долгосрочные активы",         "Раздел 1 Долгосрочные активы",         "ДолгАктивы",  "ДА" }, 
						{2, "Раздел 2", "Краткосрочные активы",        "Раздел 2 Краткосрочные активы",        "КраткАктивы", "КА" },
						{3, "Раздел 3", "Собственный капитал",         "Раздел 3 Собственный капитал",         "СобКапитал",  "СК" },
						{4, "Раздел 4", "Долгосрочные обязательства",  "Раздел 4 Долгосрочные обязательства",  "ДолгОбяз",    "ДО" },
						{5, "Раздел 5", "Краткосрочные обязательства", "Раздел 5 Краткосрочные обязательства", "КраткОбяз",   "КО" },
						{0, "Раздел 6", "Забалансовые счета",          "Раздел 0 Забалансовые счета",          "ЗабСчета",    "ЗС" }
					} ),
				selectRow = Table.SelectRows(Источник, each ([Раздел_ID] <> 0))
			in 	selectRow,
					
		Спр_ББ_Группа  = // Справочники групп статей 
			let	Источник = #table(type table 
					[
						Группа_ID = number, 
						ГруппаИмя = text
					],
					{
						{1, "Группа"}, 
						{2, "Подгруппа"}
					} 
				)
			in 	Источник,
			
		Спр_ББ_Уровень = // Справочники уровней 
				let Источник = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WMlTSUbqw+GLDhX0XNl3YemHvxR4FsNDyCxsuNl5sutijFKsTrWSEqQostODChgvbL2wBCu0GqzPGVAcWWnixCWge0LSL/WB1JpjqwELzgQJbQPYiVMcCAA==", BinaryEncoding.Base64), Compression.Deflate)), 
					let _t = ((type nullable text) meta [Serialized.Text = true]) 
					in type table [Уровень_ID = _t, УровеньПрефикс = _t, УровеньИмя = _t]),
					tr = Table.TransformColumnTypes(Источник,{{"Уровень_ID", Int64.Type}, {"УровеньИмя", type text}})
				in  tr
		],
	
	ОПУ    = [ // Отчет о прибылях и убытках
		Спр_ОПУ_ВидДеятельности = 
			let Источник = #table(type table
					[   
						ВидДеятельности_ID  = Int64.Type,
						ДиапазонСтрок_ID    = text,
						ВидДеятельностиИмя  = text
					],
					{
						{1, "10..90",   "Текущая деятельность"}, 
						{2, "100..140", "Инвестиционная деятельность"},
						{3, "80..100",  "Другие показатели"}
					}
				)
			in Источник,
		
		Спр_ОПУ_ВидПрибыли 		= // Справочники видов показателей прибыли
			let Источник = #table(type table [ 
				Прибыль_ID = Int64.Type, 
				ПрибыльВид = text ],
				{
					{1, "Показатели прибыли"}, 
					{2, "Другие показатели"}
				} )
			in Источник,
		
		Спр_ОПУ_Группа 			= // Справочники групп статей
			let Источник = #table(type table [
				Группа_ID = number, 
				ГруппаИмя = text    ],
				{
					{1,"Группа"}, 
					{2,"Подгруппа"}
				} )
			in Источник
		], 
	
	ОДДС   = [ // Отчет о движении денежных средств
		Спр_ОДДС_ВидДеятельности = // 
			let Источник = #table(type table 
					[
						ВидДеятельности_ID   = Int64.Type,
						ДиапазонСтрокДеят_ID = text,
						ВидДеятельностиИмя 	 = text
					],
					{
						{1, "20..40",	"Текущая деятельность"}, 
						{2, "50..70",	"Инвестиционная деятельность"},
						{3, "80..100",	"Финансовая деятельность"}, 
						{4, "110..140",	" Другое"}
					}
					)
			in  Table.SelectColumns(Источник, {"ВидДеятельности_ID","ВидДеятельностиИмя"}),

		Спр_ОДДС_ВидДвижения 	 = // Справочники видов движения
			let Источник = #table(type table 
					[
						ВидДвижения_ID  	 = Int64.Type,
						ДиапазонСтрокДвиж_ID = text, 
						ВидДвиженияИмя 		 = text
					],
					{
						{1, "20..24",	"Поступило"},
						{1, "50..55",	"Поступило"}, 
						{1, "80..84",	"Поступило"}, 
						{2, "30..34",	"Выбыло"},
						{2, "60..64",	"Выбыло"},
						{2, "90..95",	"Выбыло"},
						{3, "40",		"РезультатТек"}, 
						{4, "70",		"РезультатИнв"}, 
						{5, "100",		"РезультатФин"}, 
						{6, "110",		"РезультатОбщий"}, 
						{7, "120..140",	"Другие"}
					}
				)
			in  Table.SelectColumns(Источник, {"ВидДвижения_ID","ВидДвиженияИмя"})
		
		],

	ОСК    = [ // Отчет об изменении собственного капитала
		Спр_ОСК_Раздел = 
			let	from = #table(type table 
				[	
					Раздел_ID 	  = number, 
					РазделПрефикс = text, 
					РазделИмя 	  = text, 
					СтрокаНач     = number, 
					СтрокаКон     = number
				],
				{    
					{1, "Раздел 1", "Позапрошлый год", 10, 40}, 
					{2, "Раздел 2", "Прошлый год", 50, 130},
					{3, "Раздел 3", "Текущий год", 140, 200}
				}),
				to = Table.AddColumn(from, "РазделПолноеИмя", each [РазделПрефикс] & " " & [РазделИмя])
			in	Table.SelectColumns(to, {"Раздел_ID","РазделПрефикс","РазделИмя"})
		],
	
	ПРИЛОЖЕНИЕ = [ // ПРИЛОЖЕНИЕ К БАЛАНСУ
		Спр_Разделы 	= // Названия разделов отчета Приложение в Балансу
				let Источник = #table(type table 
					[   Раздел_ID     = number, 
						РазделПрефикс = text, 
						РазделИмя     = text, 
						Str1          = number,
						Str2          = number],
					{   {1, "Раздел 1", "Основные средства и нематериальные активы", 		1, 36}, 
						// {2, "Раздел 1", "Справочно. ", 1, 5}, 
						// {3, "Раздел 1", "Источники финансирования сложений ", 30, 36}, 
						{2, "Раздел 2", "Незавершенные вложения во внеоборотные активы", 	50, 65},
						{3, "Раздел 3", "Резервы по сомнительным долгам", 					70, 70},
						{4, "Раздел 4", "Финансовые вложения", 								80, 95},
						{5, "Раздел 5", "Полученные кредиты и займы", 						140, 171},
						{6, "Раздел 6", "Справка о платежах в бюджет и ФСЗН", 				1010, 1320},
						{7, "Раздел 7", "Забалансовые счета", 								1401, 1411}
					}),

					AddCol  = Table.AddColumn(Источник, "РазделИмяПрефикс", each [РазделПрефикс] & " " & [РазделИмя]),
					Reorder = Table.ReorderColumns(AddCol,
							let nms = Table.ColumnNames(AddCol), lst = {"РазделИмяПрефикс"}, col = 3  // 0 - если в начало
							in  List.InsertRange(List.Difference(nms, lst ), col, lst) ),
					SelCol = Table.SelectColumns(Reorder, List.FirstN(Table.ColumnNames(Reorder), 4) )
				in  SelCol,
			
		Спр_Заголовки 	= // заголовки разделов
				let Источник = #table(type table 
					[   
						IndexIn = text,      // порядковый номер в заголовка пределах раздела отчета
						ЗаголовокИмя = text, // имя заголовка 
						НомерРаздела = text  // номера разделов, в которых есть этот заголовок
					],
					{   
						{"1", "На начало года","1"}, 
						{"2", "Поступило","1"},
						{"3", "Выбыло","1"},
						{"4", "На конец отчетного периода всего","1"},
						{"5", "На конец отчетного периода всего, и.т.ч. полностью самортизированных","1"},
						
						{"1", "На начало года","2"},
						{"2", "Поступило","2"},
						{"3", "Выбыло","2"},
						{"4", "На конец периода","2"}, 
						
						{"1", "На начало года","3"},
						{"2", "Создано","3"},
						{"3", "Списано","3"},
						{"4", "На конец периода","3"}, 

						{"1", "На начало года","4,7"},
						{"2", "На конец периода","4,7"}, 

						{"1", "На начало года","5"},
						{"2", "Получено","5"},
						{"3", "Погашено","5"},
						{"4", "На конец периода","5"}, 

						{"1", "Пpичитается по расчету","6"},
						{"2", "Изpасходовано","6"},
						{"3", "Фактически уплачено","6"}
					}),
	
					// разворачиваем по строкам через запятую
						ExpandList = Table.ExpandListColumn(
							Table.TransformColumns(Источник, 
									{{"НомерРаздела", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), 
									let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}}), 
									"НомерРаздела"),
					
					// 	объединяем два столбца в новый используя функцию
						CombineCol = 
							Table.CombineColumns(ExpandList, 
									{"НомерРаздела", "IndexIn"},
									Combiner.CombineTextByDelimiter("", QuoteStyle.None),
									"Заголовок_ID"
							),
					// 	Написать надо что получилось в итоге!!!
						TransCol   = Table.TransformColumnTypes(CombineCol,{{"Заголовок_ID", Int64.Type}})
				in  TransCol
		]	
],

Report = [ 
	readme = "Библиотека функций для обогащения форм отчетности новыми колонками",
	IsReportViewGroup = true,	
	Connect = ВсеФормыВпапке2, // СДЕЛАТЬ НА КАЖДЫЙ ИСТОЧНИК Гомель, Марта, БКС
	/*БалансОднойСТрокой = 
		let
			Source = ВсеФормыВпапке2,
			from =  fxPackForm[Balans][Balance.StepEnd](
						fxPackForm[Balans][Balance.AddColumnsSprav](
							fxPackForm[Balans][Balance.AddColumnsLiquid](
								fxPackForm[Balans][Balance.AddColumnsLevels](
									fxPackForm[Balans][Balance.Cleaning](
										fxPackForm[Balans][Balance.ExpandColumns](
											fxPackForm[Balans][Balance.Load](
												fxPackForm[Balans][Balance.Connect](Source)
											)
										)
									)
								)
							)
						)
					)
		in
			from,*/
	Balans = [ // БАЛАНС = на входе 4 общие колонки и 4 колонки в подтаблице
		// ДОБАВИТЬ: 1. Обработку ошибок при добавлении в папку с отчетами других файлов
		Balance.Connect 		=  // подключение к источнику СДЕЛАТЬ НА КАЖДЫЙ ИСТОЧНИК Гомель, Марта, БКС
			(tbl as table) =>
				let from = 	Table.SelectColumns(
								Table.SelectRows(tbl, each [Report_ID] = "balance"),
								{"Company","DataSet","Report_ID","Date"} & {"Content"}
							)
				in  from,
			
		Balance.Load 			=  // загрузка СДЕЛАТЬ НА КАЖДЫЙ ИСТОЧНИК Гомель, Марта, БКС 
			(tbl as table) =>
				func[fxSelectRowFromKod](tbl,
					{"Column1","Column2","Column3","Column4"},
					{"Статья", "Строка_ID", "На начало", "На конец"},
					"Строка_ID", 10, 1000),

		Balance.ExpandColumns    = // Шаг 2. Разворачиваем баланс и типизируем 4 + 4 колонки
			/* 	На входе:	
				- 4 общие колонки: 			Название организации + Номер датасета + Вид ответа + Дата отчета
				- 4 обязательные колонки:   Статья + Строка + На начало + На конец
			*/
			(tbl as table) =>
				let 
				// 	список начальных колонок итоговой таблицы: общие + обязательные (форму еще не добавили)
					lst1 = {"Company", "DataSet", "Report_ID", "Date"},
					lst2 = {"Статья", "Строка_ID", "На начало", "На конец"},

				/* ТРАНСФОРМАЦИЯ --------------------------------------------------------------------------------------*/
					// Разворачиваем подтаблицы. Список столбцов без первой колонки "Date", которой внутри нет
					// надежней будет колонки определять сразу до разворота  ColNames = Table.ColumnNames(from2[Step2]{0})
					EXPAND = Table.ExpandTableColumn(tbl, "Step2", lst2, lst2),
					
					// меняем тип колонок с трансформацией. Можно делать это и в подтаблицах, до разворота
					
					Transform = Table.TransformColumns(
							EXPAND,
							{ 
								{ lst1{0}, (x)=> x, type text},               	// организация
								{ lst1{1}, (x)=> Number.From(x), Int64.Type},   // номер датасета
								{ lst1{2}, (x)=> "balance", type text},         // отчет - Сделать join со Справочникиом отчетов
								{ lst1{3}, (x)=> Date.From(x), Date.Type},      // дата отчета
								
								{ lst2{0}, (x)=> x, type text}, 													// статья
								{ lst2{1}, (x)=> try Number.From(x) otherwise null, Int64.Type}, 					// строка
								{ lst2{2}, (x)=> if Number.From(x) = 0 then null else Number.From(x), Int64.Type}, 	// на начало
								{ lst2{3}, (x)=> if Number.From(x) = 0 then null else Number.From(x), Int64.Type}  	// на конец
							} )
				in 
					Transform,

		Balance.Cleaning         = // Шаг 3. Трансформируем колонки (чистим, меняем, обогащаем)
			(tbl)=> // добавить в аргументы имя колонки или список колонок типа optional? ColNamesOrList + через если
			let 
				// удаляем лишний текст в колонке
				removeText = let 
					lst    = {"в том числе: ", " (лизинга)"},
					remove = F[REPLACE][Text.RemoveMulti](tbl, "Статья", lst)
				in  remove,

				// добавляем отступ в колонке "Статья"
				offsetText = F[fxOffsetTextInColumn](removeText, "Статья", "Строка_ID", 10)
			in 
				offsetText,
		
		Balance.AddColumnsLevels = // Шаг 4. Добавляем 4 колонки иерархии: Часть, Раздел, Статья, Подстатья (на основе колонки "Номер строки")
			(tbl as table)=>
			let 
				// добавляем колонки с цифровым ID атрибутов
				AddColumns = F[fxAddMultiColumns](tbl, {
					// 	статьи с номерами строк от 100 до 399 соответствуют Части 1 (Активы), другие Части 2 (Пассивы)
					{ 	"Часть_ID",  each 
							if List.Contains ({110..399}, [Строка_ID]) = true 
							then 1 
							else 2
						, Int64.Type},
					
					// 	номер раздела = первому символу номера строки. 0 - не вошедшие в разделы
					{ 	"Раздел_ID", each 
							let f=(x as text) => Record.FieldOrDefault([1=1, 2=2, 4=3, 5=4, 6=5, 0=0], x, null),
								Номер = Text.Start(Text.From([Строка_ID]),1)    
							in f(Номер)
						, Int64.Type},
					
					// 	если первая буква большая, то группа (1), если маленькая, то подгруппа (2)
					{ 	"Группа_ID", each 
							if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1)) = true 
							then 1 
							else 2
						, Int64.Type},
					
					// 	уровень иерархии: 1 - Часть, 2 - Раздел, 3 - Статья (= 0 в конце № строки), 4 - Подстрока (не равно 0 в конце № строки)
					{ 	"Уровень_ID", each 
								 if List.Contains ({300, 700},[Строка_ID]) = true then 1 
							else if List.Contains ({190, 290, 390, 490, 590, 690},[Строка_ID]) = true then 2 
							else if Number.Mod([Строка_ID], 10) <> 0  and 
									not List.Contains ({411},[Строка_ID]) = true then 4 // исключение для строки 411 "Паевые взносы"
							else 3
						, Int64.Type}
					})
			in 
				AddColumns,

		Balance.AddColumnsLiquid = // Шаг 4. Добавляем 3 колонки для Анализа ликвидности: Статус, Категория, Группа 
			(tbl as table)=>
			let 
				// добавляем колонки с цифровым ID атрибутов
				AddColumn = F[fxAddMultiColumns](tbl, {
				
					// выбираем строки, которые НЕ включаются в расчет = false
					{ "Ликвидность_Статус", each 
						if List.Contains ({190,290,390,490,590,690,300,700,630,211,212,213,214,215,216},[Строка_ID]) = true 
						then false 
						else true, type logical},
					
					// те что включаются делим по категориям 4 + 4 
					{"Ликвидность_Группа",
						each if List.Contains ({260, 270},[Строка_ID]) = true then "А1"
						else if List.Contains ({210, 250, 280},[Строка_ID]) = true then "А2"
						else if List.Contains ({220, 230, 240, 150, 170},[Строка_ID]) = true then "А3"
						else if List.Contains ({110, 120, 130, 140, 160, 180},[Строка_ID]) = true then "А4"
						else if List.Contains ({632, 633, 634, 635, 636, 637, 638, 639},[Строка_ID]) = true then "П1" 
						else if List.Contains ({610, 620, 640, 650, 660, 670, 631},[Строка_ID]) = true then "П2" 
						else if List.Contains ({510, 520, 530, 540, 550, 560},[Строка_ID]) = true then "П3"
						else if List.Contains ({410, 420, 430, 440, 450, 460, 470, 480},[Строка_ID]) = true then "П4" 
						else null, type text}, // мб сделать null ? было "0"

					// КатегорияЛиквидности
					{"Ликвидность_Категория",
						each if List.Contains ({260, 270} & {632, 633, 634, 635, 636, 637, 638, 639},[Строка_ID]) = true then "АП1"
						else if List.Contains ({210, 250, 280} & {610, 620, 640, 650, 660, 670, 631},[Строка_ID]) = true then "АП2"
						else if List.Contains ({220, 230, 240, 150, 170} & {510, 520, 530, 540, 550, 560},[Строка_ID]) = true then "АП3"
						else if List.Contains ({110, 120, 130, 140, 160, 180} & {410, 420, 430, 440, 450, 460, 470, 480} ,[Строка_ID]) = true then "АП4"
						else "0", type text}
				})
			in 
				AddColumn,
		
		Balance.AddColumnsSprav  = // Шаг 4. Добавляем разные колонки из Справочникиов через Join 
			(tbl as table) =>
				let
					from = tbl,
					
					// создаем Справочники статей из 2 + 1 колонок с рабочими названиями статей
					Спр_ББ_Статья = 
						let
							nms      = {"Строка_ID"} & {"Статья", "СтатьяИмя"} & {"СтатьяСтрока"},
							distinct = Table.Distinct(Table.SelectColumns(from, {nms{0}, nms{1}}), {nms{0}} ),
							rename   = Table.RenameColumns(distinct, {{nms{1}, nms{2}}} ),
																					
							// МЕНЯЕМ текст в столбце "СтатьяИмя" для удобства
							lst  = { 
									{"Налог на добавленную стоимость",                      "НДС"},
									{"НДС по приобретенным товарам, работам, услугам",      "НДС по приобретенным Т,Р,У"},
									{"Долгосрочные активы, предназначенные для реализации", "Долгосрочные активы для реализации"},
									{"Долгосрочные обязательства по лизинговым платежам",   "Долгосрочные обязательства по лизингу"},
									{"Денежные средства и эквиваленты денежных средств",    "Денежные средства и их эквиваленты"}, 
									{"кредиторская задолженность",                          "КЗ"},                
									{"Вложения в долгосрочные активы",                      "Вложения в ДА"},
									{"дебиторская задолженность",                           "ДЗ"} 
								},
							
							// меняем текст по словарю
							replace = F[REPLACE][Text.ReplaceMulti](rename, nms{2}, lst),

							addColumn = Table.AddColumn(replace, nms{3}, each Text.From([Строка_ID]) & " " & [СтатьяИмя])
							
							//  addColOther = Table.InsertRows(addColumn, 61, {[ID_Строка = 800, СтатьяИмя = "Прочие", СтрокаСтатья = "800 Прочие"]} )
						in  addColumn,

					// добавляем значения колонок с ID (если переключатель IsReportViewGroup = True ?!)
					Join = [ 
						Join_Часть   = Table.Join(from,    		"Часть_ID",     Справочники[БАЛАНС][Спр_ББ_Часть],   "Часть_ID"),
						Join_Раздел  = Table.Join(Join_Часть ,  "Раздел_ID",    Справочники[БАЛАНС][Спр_ББ_Раздел],  "Раздел_ID"),
						Join_Группа  = Table.Join(Join_Раздел,  "Группа_ID",    Справочники[БАЛАНС][Спр_ББ_Группа],  "Группа_ID"),
						Join_Уровень = Table.Join(Join_Группа,  "Уровень_ID",   Справочники[БАЛАНС][Спр_ББ_Уровень], "Уровень_ID"),
						Join_Строка  = Table.Join(Join_Уровень, "Строка_ID",    Спр_ББ_Статья, 						"Строка_ID") // ? зачем это
					]  [Join_Строка]
				in  Join,

		Balance.StepEnd =          // Шаг 5. Наводим красоту перед выгрузкой 
			(tbl as table) =>
			let
				// перегруппировка колонок согласно кастомного списка в Excel для БД
				/*ReorderColumns = F[REORDER][Reorder.AllColumns](
					OffsetChar, 
					Value.Metadata(СправочникиСтолбцов)[balance_listcol][Имя] 
				),*/

				/* РЕЗУЛЬТАТ ------------------------------------------------------------------------------------------*/

				// выводим или узкую таблицу или с допколонками со значениями групп ID "все в одном"
				_Result_Step2 = tbl,
					/*if Parameters[IsReportViewGroup]
					then ReorderColumns 
					else Balance.AddColumnsSprav(tbl),*/
					
				// ставим метки на шагах запроса (потом использовать to = Value.Metadata(А_Баланс_Загрузка)[Expand_]_
				Result = _Result_Step2 //meta [Expand_ = EXPAND]
			in 
				Result
	],

	OPU  = [ // ОТЧЕТ О ПРИБЫЛЯХ И УБЫТКАХ
		OPU.Load	= 
			(tbl as table) =>
				func[fxSelectRowFromKod](tbl,
					{"Column1","Column2","Column3","Column4"},
					{"Статья", "Строка_ID", "Текущий период", "Прошлый период"},
					"Строка_ID", 9, 1000),
		OPU.Expand 	 = 
			(tbl as table) => 
			[	// разворачиваем подтаблицы
				lstColumnNew = {"Статья", "Строка_ID", "Текущий период", "Прошлый период"},
				Expand = Table.ExpandTableColumn(tbl, "Step2", lstColumnNew, lstColumnNew),
				
				// трансформируем столбцы с указанием типа
				lst1 = {"Company", "DataSet", "Report_ID", "Date"},
				
				tr = Table.TransformColumns(Expand,{ 
					{ lst1{0}, (x)=> x, type text},               	// организация
					{ lst1{1}, (x)=> Number.From(x), Int64.Type},   // датасет
					{ lst1{2}, (x)=> "opu", type text},         	// отчет Сделать join со справончиком отчетов
					{ lst1{3}, (x)=> Date.From(x), Date.Type},      // дата
					
					{"Статья",         (x)=> x, Text.Type},
					{"Строка_ID",      each try Number.From(_) otherwise null, type number},
					{"Текущий период", each try Number.From(_) otherwise _, Number.Type},
					{"Прошлый период", each try Number.From(_) otherwise _, Number.Type}
					} 
				)
			] [tr],

		OPU.Clean 	 = 
			(tbl as table) =>
			[
				// 	меняем 0 на null, чтобы 0 в визуализации не появлялся
					replace0ToNull = Table.ReplaceValue(tbl, 0, null, Replacer.ReplaceValue, {"Текущий период", "Прошлый период"}),

				// 	убираем текст внутри скобок + сами скобки + текст после "(" по условию
					removeText1 = 
						let f=(txt)=> 
							if   F[Text.CompareText]({"010","100","±"}, txt) = true // проверяем строку на наличие символов 
							then Text.BeforeDelimiter(txt, "(", {0, RelativePosition.FromEnd}) // оставляем текст ДО, а "(" и ПОСЛЕ убираем
							else txt
						in Table.TransformColumns(replace0ToNull, {"Статья", f, type text} ),

				// 	удаляем текст
					removeText2 = F[REPLACE][Text.RemoveMulti](removeText1, "Статья", 
						{"В том числе: ", ""}
					),

				// 	делаем отступ, если номер строки не четный
					offsetText = 
						F[fxOffsetTextInColumn](removeText2, "Статья", "Строка_ID", 10),
				
				to = offsetText
			] [to],

		OPU.AddCol1	 = // добавляем 3 колонки Группа_ID + Прибыль_ID + ВитдДеят_ID 
			(tbl as table) =>
			[	add = F[fxAddMultiColumns](tbl, {
					// если первая буква большая, то это 1-Группа, иначе 2-Подгруппа
					{ "Группа_ID", each 
						if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1)) = true 
						then 1 
						else 2, Int64.Type},

					// 1 - статьи со словом "прибыль" в названии
					{ "Прибыль_ID", each 
						if List.Contains ({30,60,90,140,150,201,210,240},[Строка_ID]) = true 
						then 1 
						else 2, Int64.Type},

					// 1-Текущая, 2-Инвестиционная, 3-финансовая деятельность
					{ "ВидДеят_ID", each 
						if List.Contains ({10..90},[Строка_ID]) = true then 1
						else if List.Contains ({100..140},[Строка_ID]) = true then 2
						else 3, Int64.Type},
					
					// 	уровень иерархии
					{ 	"Уровень_ID", each 
							if Number.Mod([Строка_ID], 10) <> 0 = true 
							then 1 
							else 2
						, Int64.Type},
					
					// Номер года
					{"Год", each Date.Year([Date]), Int64.Type}
				} ),
				to = add
			][to],

		OPU.AddCol2  = // добавляем колонки из Справочникиов
			(tbl as table) =>
			[	// Справочники статей
				Спр_ОПУ_Статья = 
					let
						Источник        = Table.Distinct(tbl[[Строка_ID],[Статья]], {"Строка_ID"}),
						AddCol_СтрокаID = Table.AddColumn(Источник, "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья]),
						RenameColumns   = Table.RenameColumns(AddCol_СтрокаID,{{"Статья", "СтатьяИмя"}}),
						
						// Делаем множественную замену в подстроке столбца
						МультиЗамена = [ 
							dict = List.Buffer(Table.FromValue({ 
									{"товаров, работ, услуг",    ""},
									{"основных средств",                    "ОС"},
									{"нематериальных активов",              "НА"},
									{"долгосрочных активов",                "ДА"},
									{"инвестиционной деятельности",         "ИнвДеят"},
									{"финансовой деятельности",             "ФинДеят"},
									{"отложенных налоговых активов",        "ОНА"},
									{"отложенных налоговых обязательств",   "ОНО"}
								}) [Value]),
							func = (x)=> List.Accumulate(dict, x, (state, current) => Text.Replace(state, current{0}, current{1} )),
							b1   = Table.TransformColumns(RenameColumns, {"СтатьяИмя", func}),
							b2   = Table.TransformColumns(b1, {"СтатьяСтрока", func})
							][b2],
						
						// зачем?
						AddIndexColumn = Table.AddIndexColumn(МультиЗамена, "СтрокаИндекс", 1, 1, Int64.Type)
					in  AddIndexColumn,
				
				// обединяем таблицу фактов со Справочникиами с формированием дополнительных колонок-срезов
				Join = 
					let Join_ВидДеят    = Table.Join(tbl,  	  		  "ВидДеят_ID", Справочники[ОПУ][Спр_ОПУ_ВидДеятельности],  "ВидДеятельности_ID"),
						Join_ВидПрибыли = Table.Join(Join_ВидДеят,    "Прибыль_ID", Справочники[ОПУ][Спр_ОПУ_ВидПрибыли],       "Прибыль_ID"),
						Join_Группа     = Table.Join(Join_ВидПрибыли, "Группа_ID",  Справочники[ОПУ][Спр_ОПУ_Группа],           "Группа_ID"),
						Join_Строка     = Table.Join(Join_Группа,     "Строка_ID",  Спр_ОПУ_Статья,								"Строка_ID")
					in  Join_Строка,
				
				to = Join
			][to],

		OPU.StepEnd	 = 
			(tbl as table) => [
				// 	перегруппировка колонок: (DataSet, Date  + Статья, Строка, Текущий, Прошлый) + Остальные по алфавиту
				// 		если так, Reorder = Table.ReorderColumns(Join, {"DataSet", "Статья", "ID_Строка", "Текущий период", "Прошлый период"} )
				// 		,то  эти вперед, остальные как есть
				/*ReorderColumns = 
					F[REORDER][Reorder.AllColumns](
						ДобавляемОтступ, 
						Value.Metadata(СправочникиСтолбцов)[opu_listcol][Имя] 
					),*/
				
				// выводим: true - узкую таблицу, fase - таблицу ""все в одном" с допколонками со значениями групп
				result = if IsReportViewGroup 
						then tbl 
						else tbl, // AddColumnY
				
				to = result
			] [to]
		
	],
		
	ODDS = [ // ОТЧЕТ О ДВИЖЕНИИ ДЕНЕЖНЫХ СРЕДСТВ
		ODDS.Load	 = 
			(tbl as table) =>
				func[fxSelectRowFromKod](tbl,
					{"Column1","Column2","Column3","Column4"},
					{"Статья", "Строка_ID", "Текущий период", "Прошлый период"},
					"Строка_ID", 9, 1000),

		ODDS.Expand  = 
			(tbl as table) =>
			[	// разворачиваем подтаблицы
				lstColumnNew = {"Статья", "Строка_ID", "Текущий период", "Прошлый период"},
				Expand = Table.ExpandTableColumn(tbl, "Step2", lstColumnNew, lstColumnNew ),
				
				// трансформируем столбцы с указанием типа
				TransformAllColumns = Table.TransformColumns(Expand,
					{ 
						{ "Company", 	(x)=> x, type text},               	// организация
						{ "DataSet", 	(x)=> Number.From(x), Int64.Type},  // датасет
						{ "Report_ID", 	(x)=> "odds", type text},         	// отчет Сделать join со справончиком отчетов
						{ "Date", 		(x)=> Date.From(x), Date.Type},     // дата
						
						{"Статья",         (x)=> x, Text.Type},
						{"Строка_ID",      each try Number.From(_) otherwise null, type number},
						{"Текущий период", each try Number.From(_) otherwise _, Number.Type},
						{"Прошлый период", each try Number.From(_) otherwise _, Number.Type}
					} 
				),
				
				to = TransformAllColumns
			] [to],

		ODDS.Clean   = // 
			(tbl as table) =>
			[
				// меняем 0 на null (чтобы в визуализации не появлялся)
				Replace0ToNull = Table.ReplaceValue(tbl, 
					0, null, Replacer.ReplaceValue, 
					{"Текущий период", "Прошлый период"}
				),
				
				// удаляем лишний текст
				removeList  = {"В том числе: ", "В том числе:", " - всего"},
				multiRemove = F[REPLACE][Text.RemoveMulti](Replace0ToNull,"Статья", removeList),
				
				// трансформируем текст
				transform   = Table.TransformColumns(multiRemove, 
					{"Статья", (x)=> [
							txt = F[TEXT][Text.ReplaceBetweenDelimiters](x, "(",")"), 	// убираем текст в скобках
							to =  	if   Text.Contains(txt,"Остаток денежных средств") 	// удаляем даты "... на 01.01.2015" и т.д.
									then Text.BeforeDelimiter(txt," на") 
									else txt
						] [to]
					}),
				
				// делаем отступ для подстатей
				addOffset = F[fxOffsetTextInColumn](transform, "Статья", "Строка_ID", 10),
				
				to = addOffset
			][to],

		ODDS.AddCol1 = // 
			(tbl as table) =>
			[	// добавляем колонки
				AddColumns = F[fxAddMultiColumns](tbl, {

					{ "Группа_ID",  each 
							if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1)) = true 
							then 1 
							else 2 , Int64.Type}, // 1- Группа, 2 - Подгруппа
					
					{ "ВидДеятельности_ID", 
							each if List.Contains ({20..40},[Строка_ID]) = true then 1 // "Текущая"
							else if List.Contains ({50..70},[Строка_ID]) = true then 2 // "Инвестиционная"
							else if List.Contains ({80..100},[Строка_ID]) = true then 3 // "Финансовая"
							else 4, Int64.Type}, // "Другое"
					
					{ "ВидДвижения_ID", 
							each if List.Contains ({20..24, 50..55, 80..84},[Строка_ID]) = true then 1 // "Поступление" 
							else if List.Contains ({30..34, 60..64, 90..95},[Строка_ID]) = true then 2 // "Выбытие" 
							else if List.Contains ({40},[Строка_ID]) = true then 3 // "РезультатТек" 
							else if List.Contains ({70},[Строка_ID]) = true then 4 // "РезультатИнв" 
							else if List.Contains ({100},[Строка_ID]) = true then 5 // "РезультатФин" 
							else if List.Contains ({110},[Строка_ID]) = true then 6 // "РезультатОбщий" 
							else 7 /* "Другое" */, type text},
					
					{"Год", each Date.Year([Date]), Int64.Type}                        
				}),
				
				to = AddColumns
			] [to],
									
		ODDS.AddCol2 = // 
			(tbl as table) =>
			[	Спр_ОДДС_Статья = 
					let 
						// берем две колонки "Строка_ID","Статья"
						from = 	Table.Distinct(Table.SelectColumns(tbl, {"Строка_ID","Статья"}), {"Строка_ID"} ),
					
						// УДАЛЯЕМ лишний текст
						lst    = {" по отношению к белорусскому рублю", ""},
						REMOVE = F[REPLACE][Text.RemoveMulti](from, "Статья", lst),

						// ДОБАВЛЯЕМ колонку "СтатьяСтрока" соединяя имеющиеся
						AddColumn = Table.AddColumn(REMOVE, "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья]),
				
						// меняем имя колонки "Статья", чтобы сделать в ней рабочие сокращенные названия статей
						RenameCol = Table.RenameColumns(AddColumn, {{"Статья", "СтатьяИмя"}} ),

						// МЕНЯЕМ текст в подстроке столбца "СтатьяИмя" для удобства работы
						replaceList  = { 
								{"текущей, инвестиционной и финансовой деятельности", 	"ТИФ деят"},
								{"от покупателей продукции, товаpов, заказчиков pабот, услуг", 	"покупателей товаров"},
								{"от покупателей материалов и других запасов", 					"покупателей других запасов"},
								{"доходы от участия в уставном капитале других организаций", 	"доходы от дивидендов"},
								{"основных средств, нематериальных активов и других долгосрочных активов", "ОС, НА и других ДА"},
								{"вклады собственника имущества (учредителей, участников)", 	"вклады собственника"},
								{"на выплаты дивидендов и других доходов от участия в уставном капитале организации", "на выплаты дивидендов"},
								{"текущей деятельности",		"ТекДеят"},
								{"инвестиционной деятельности", "ИнвДеят"},
								{"финансовой деятельности", 	"ФинДеят"}/*,
								{"продукции, товаpов, заказчиков pабот, услуг","товаров"}*/
							},
						
						// заменяем содержимое колонок
						REPLACE1 = F[REPLACE][Text.ReplaceMulti](RenameCol, "СтатьяИмя", replaceList),
						REPLACE2 = F[REPLACE][Text.ReplaceMulti](REPLACE1, "СтатьяСтрока", replaceList)

					in 
						REPLACE2,

				Join = 
					let Join_Вид1   = Table.Join(tbl,  	  	"ВидДеятельности_ID", Справочники[ОДДС][Спр_ОДДС_ВидДеятельности],  	"ВидДеятельности_ID"),
						Join_Вид2   = Table.Join(Join_Вид1, "ВидДвижения_ID", 	  Справочники[ОДДС][Спр_ОДДС_ВидДвижения],      	"ВидДвижения_ID"),
						Join_Строка = Table.Join(Join_Вид2, "Строка_ID",  		  Спр_ОДДС_Статья,						 		"Строка_ID")
					in  Join_Строка,

				to = Join
			][to],
			
		ODDS.StepEnd = // 
			(tbl as table) =>
			[					
				// перегруппировка колонок согласно кастомного списка в Excel (для копирования в БД!)
				/*ReorderColumns = 
					F[REORDER][Reorder.AllColumns](
						OffsetChar, 
						Value.Metadata(СправочникиСтолбцов)[odds_listcol][Имя] 
					),*/
			
				// выводим: true - узкую таблицу (шаг AddColumn_Year), fase - таблицу ""все в одном" с допколонками (шаг Отступ)
				RESULT = 
					if IsReportViewGroup 
					then tbl 
					else tbl	//Replace0ToNull
			] [RESULT]
	],

	OSK = [ // ОТЧЕТ ОБ ИЗМЕНЕНИИ СОБСТВЕННОГО КАПИТАЛА
		OSK.Load	= 
			// let a = func[fxSelectRowFromKod](tbl,
			// 				{"Column1","Column2","Column3","Column4"},
			// 				{"Статья", "Строка_ID", "Текущий период", "Прошлый период"},
			// 				"Строка_ID", 9, 1000),
			(tbl as table) =>
		    [  	addCol = 
					Table.AddColumn(tbl, "Content2", (x)=>
						let z = x[Content],
							a = Table.Skip(z,2), // пропускаем 2 строки (из за кода организации выдает ошибку, обходим)
							b = Table.TransformColumnTypes(a,{{"Column2", Int64.Type}}), // делаем тип колонки КодСтроки ЦелоеЧисло
							c = Table.RemoveRowsWithErrors(b, {"Column2"}), // удаляем из колонки КодСтроки ошибки, т.к. в ней есть текстовые данные
							d = Table.SelectRows(c, each [Column2] <> null and [Column2] > 2), // удаляем все строки кроме строк, с номерами статей
							e = Table.AddIndexColumn(d, "Index", 1, 1),  // добавляем колонку индекс
							k=(x)=> try Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " ") otherwise x,
							ClearData = Table.TransformColumns(e, {}, k)
						in  ClearData
					),
            	
				// убираем столбец с исходными подтаблицами
				to = Table.RemoveColumns(addCol, {"Content"} )
        	][to],

		OSK.Expand = 
			(tbl as table) =>
			[   // список заголовков в подтаблицах отчета - 10 шт.
				lstHeader = 
					{   
						"Статья", "Строка_ID",
						"Уставный капитал", "Неоплаченная часть уставного капитала",
						"Собственные акции (доли в уставном капитале)", "Резервный капитал", "Добавочный капитал",
						"Нераспределенная прибыль (непокрытый убыток)", "Чистая прибыль (убыток)",
						"Итого"
					},

				// разворачиваем таблицу
				Expand = 
					Table.ExpandTableColumn(tbl,
						"Content2", 
						Table.ColumnNames(tbl[Content2]{0}), // берем список имен колонок с первой таблицы, 10 штук
						lstHeader & {"Index"}
					) // + 1 колонка
			][Expand],
    
		OSK.Clean =
			(tbl as table) => // на входе 4 технические колонки + 10 колонок отчета + 1 колонка с индексом
			[
				// 	Делаем заголовки, кроме первого "Статья", числовыми (всего 1+10, последние 10)
					TransformColumnTypes = 
						Table.TransformColumnTypes(tbl, 
							List.Transform(
								List.LastN(Table.ColumnNames(tbl), 10), // считая по номеру с конца списка
								each {_, type number} 
							)
						),
				
				// 	Сортируем колонку "Строка_ID", 4 + 1 = 5-я в списке
					Sort = Table.Sort(TransformColumnTypes, {Table.ColumnNames(tbl){5}, Order.Ascending} ),

				// 	Делаем множественную замену в подстроке столбца
					МультиЗамена =  
						let func = (txt)=> List.Accumulate(dict, txt, (state, x) => Text.Replace(state, x{0}, x{1} )),
							dict = List.Buffer(Table.FromValue({ 
									{"В том числе: ", ""},
									{"в том числе: ", ""}
								}) [Value]),
							trans = Table.TransformColumns(Sort, {
									{"Строка_ID", (x)=>x, Int64.Type}, 
									{"Статья",    func, type text}
								})
						in trans,
				
				//  Меняем названия некоторых статей
					Replace = 
						Table.ReplaceValue(МультиЗамена, 
							(x)=>x[Статья], // в каком столбце менять. И в конце. А если несколько тут надо?!
							each Record.Field([
									10 = "Остаток на начало", 
									40 = "Скорректированный остаток",
									50 = "Увеличение собственного капитала - всего",
									100 = "Остаток на начало 1",
									110 = "Остаток на начало 2",
									140 = "Остаток на начало скорректированный",
									150 = "Увеличение собственного капиатала - всего",
									151 = "чистая прибыль",
									161 = "чистая убыток",
									200 = "Остаток на конец"
								], Text.From([Строка_ID])),
							Replacer.ReplaceValue,
							{"Статья"}
						)
			] [Replace],
		
		OSK.AddCol = 
			(tbl as table) => 
			[	
				// делаем отступ, если номер строки не четный
				offsetText = 
					F[fxOffsetTextInColumn](tbl, "Статья", "Строка_ID", 10),

				// убираем строки с позапрошлым годом + 2 колонки
					Remove  = Table.RemoveColumns(offsetText,{"Итого", "Index"}),
					AddCol  = Table.AddColumn(Remove, "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья]),

				// переворачиваем таблицу. Надо ли?!
					Unpivot = 	Table.UnpivotOtherColumns(AddCol, 
									{"Company", "DataSet", "Report_ID", "Date"} & {"Статья", "Строка_ID", "СтатьяСтрока"}, 
									"Показатель", "Значение"
								),
				
				// 	меняем 0 на null (чтобы в визуализации не появлялся) - можно убрать и оставить фильтр идущий следом + <> 0
					Replace0ToNull = Table.ReplaceValue(Unpivot, 
						0, null, Replacer.ReplaceValue, 
						{"Значение"}
					),

				//  удаляем пустые строки
					Select  = 	Table.SelectRows(Replace0ToNull, each ([Значение] <> null)),
				
				//  Зачем тут менять тип колонки?!
					Typed   = 	Table.TransformColumnTypes(Select, {"Статья", type text} ),
				
				// добавляем новые колонки
					AddColumns = 
						F[fxAddMultiColumns](Typed, {
							{ 	"Раздел_ID", 
									each if List.Contains ({10},[Строка_ID]) = true then 2 // "Позапрошлый год" ?!
									else if List.Contains ({20..110},[Строка_ID]) = true then 2 // "Прошлый год"
									else if List.Contains ({120..200},[Строка_ID]) = true then 3 // Текущий год"
									else 4 // "Другое"
								, Int64.Type
							},
							
							// 	уровень иерархии: 1 - группа, 2 - подгруппа
							{ 	"Уровень_ID", each 
									if Number.Mod([Строка_ID], 10) = 0 = true 
									then 1 
									else 2
								, Int64.Type
							}
						})
			] [AddColumns],
    
		OSK.StepEnd = 
			(tbl as table) =>
			[
				/* 	
					По нормальному этим отчетом не занимался еще. А надо ли?!
					непонятно в каком видже надо представить в конце.
					нужно ди делать анпивот или две колонки очстаток на начало + остаток на конец
					таблица сложная 
				*/
				Join = 
					let Join_Раздел  = Table.Join(tbl, "Раздел_ID", Справочники[ОСК][Спр_ОСК_Раздел], "Раздел_ID")
					in  Join_Раздел,

				Reorder = 1
			] [Join]
	],

	OFR = [ // ОТЧЕТ О ФИНАНСОВЫХ РЕЗУЛЬТАТАХ (Форма №12) - был 12ф, стал 12.
		OFR.Load    = 
			(tbl as table) => // in Table.SelectColumns(g, {"Company","DataSet","Report_ID","Date","Step2"})
				func[fxSelectRowFromKod](tbl,
					{"Column1","Column2","Column3","Column4"},
					{"Статья", "Строка_ID", "Текущий период", "Прошлый период"},
					"Строка_ID", 0, 1000),
		
		OFR.Expand  = 
			(tbl as table) =>
			/*  Разворачиваем подтаблицы. Список столбцов без первой колонки "Date", которой внутри нет
				надежней будет колонки определять сразу до разворота  ColNames = Table.ColumnNames(from2[Step2]{0}) */
			[   
				//  список заголовков в подтаблицах отчета -  шт.
					lstHeader = Table.ColumnNames(tbl[Step2]{0}),
				
				// разворачиваем таблицу
					expand = Table.ExpandTableColumn(tbl, "Step2", lstHeader,lstHeader)

			][expand],
    
		OFR.Clean   = 
			(tbl as table) => // на входе 4 технические колонки + 4 колонок отчета
			[
				//	список столбцов (пропускаем первые 4 технические столбца)
					nms = List.Skip(Table.ColumnNames(tbl), 4),

				// 	типизируем таблицу
					transform = Table.TransformColumns(tbl,
						{ 
							{ nms{0}, (x)=> F[TEXT][Text.ReplaceBetweenDelimiters](x," (",")"), type text},
							{ nms{1}, (x)=> try Number.From(x) otherwise null, Int64.Type },
							{ nms{2}, (x)=> if  Number.From(x) = 0 then null else Number.From(x), Int64.Type},
							{ nms{3}, (x)=> if  Number.From(x) = 0 then null else Number.From(x), Int64.Type}
						} 
					),
				
				//  меняем названия некоторых статей	
					Replace = 
						Table.ReplaceValue(transform, (x)=>x[Статья],
							each Record.Field([
									29 = "Средства из бюджета на покрытие убытков, на возмещение затрат",
									32 = "   на погашение основного долга",
									33 = "   на погашение процентов по кредиту, займу"
								], Text.From([Строка_ID])),
							Replacer.ReplaceValue, {"Статья"}
						),

					Replace2 = F[REPLACE][Text.ReplaceMulti](Replace, "Статья", {
						{"налог на добавленную стоимость", "НДС"},
						{"реализации продукции, товаров, работ, услуг", "реализации"}
					}),
				
				to = Replace2
			][to],
		
		OFR.AddCol  = 
			(tbl as table) => 
			[	
				AddCol  = Table.AddColumn(tbl, "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья], type text),
				to = AddCol
			] [to],
    
		OFR.StepEnd = 
			(tbl as table) =>
			[
				to = tbl
			] [to]
	],

	PRIL = [ // ПРИЛОЖЕНИЕ К БАЛАНСУ
		PRIL.Load    = 
			(tbl as table) =>
				func[fxSelectRowFromKod](tbl,
					{"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7"},
					{"Статья", "Строка_ID", "1", "2", "3", "4", "5"},
					"Строка_ID", 0, 2000),
				
		PRIL.Expand  = 
			(tbl as table) =>
				[			
					// 	список колонок итоговой таблицы, сквозной, от 2 до 5 в каждом разделе
						// lstNewColNаmes = {"Статья", "Строка_ID", "1", "2", "3", "4", "5"},
					
					//  список заголовков в подтаблицах отчета -  шт.
						lstHeader = Table.ColumnNames(tbl[Step2]{0}), // 7 колонок

					// 	разворачиваем подтаблицы (список столбцов без первой колонки "Date", которой внутри нет)
						EXPAND = Table.ExpandTableColumn(tbl, "Step2", lstHeader,  lstHeader),
			
					//  список всех колонок после разворота
						nms = List.Skip(Table.ColumnNames(EXPAND), 4),

					// меняем тип колонок с трансформацией. Можно делать это и в подтаблицах, до разворота
					Transform = Table.TransformColumns(EXPAND,
						{ 
							{ nms{0}, (x)=> x, type text},  // Статья
							{ nms{1}, (x)=> x, Int64.Type}, // Строка
							{ nms{2}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
							{ nms{3}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
							{ nms{4}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
							{ nms{5}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
							{ nms{6}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type}
						} 
					),
					to = Transform
			][to],

		PRIL.Clean   = 
			(tbl as table) =>
			[
				// 	удаляем лишнее
					replace = F[REPLACE][Text.ReplaceMulti](tbl, "Статья", {
							{"в т.ч.: ", ""},
							{" (сумма стр.1010-1130)", ""},
							{" (сумма стр.1010-1150)", ""},
							{" (сумма стр. 1210, 1220)", ""},
							{" - всего", ":"},
							{" - ВСЕГО", ":"},
							{" (авансы выданные, предоплата)", ""}
						}),

				// 	делаем отступ для подстатей
					addOffset = F[fxOffsetTextInColumn](replace, "Статья", "Строка_ID", 10),

				//  меняем названия некоторых статей	
				//  	надо что то другое придумать, чтобы разные названия полей в разные года свести 
				//  	к одному значени. Например было "Здания" и "Соружения", стало "Здания и сооружения"
				//  	в начале может поменять/объединить/удалить
					Replace = 
						Table.ReplaceValue(addOffset, (x)=>x[Статья],
							each 
								if Date.Year([Date]) = 2020 or Date.Year([Date]) = 2021 
								then Record.Field([11 = "    здания "], Text.From([Строка_ID])) // и сооружения
								else 
									if Date.Year([Date]) = 2016
									then Record.Field([11 = "    здания и сооружения"], Text.From([Строка_ID]))
										else Record.Field([11 = _[Статья] ], Text.From([Строка_ID]) )
								,
							Replacer.ReplaceValue, {"Статья"}
						),

				// 	разворачиваем	
					Unpivot = 
						Table.UnpivotOtherColumns(Replace, 
							{"DataSet","Company","Report_ID", "Date", "Статья","Строка_ID"}, 
							"Заголовок", "Значение"
						),
				
				to = Unpivot
			][to],
		
		PRIL.AddCol  =
			(tbl as table) =>
			[	
				// 	добавляем колонки
					AddedColumns = F[fxAddMultiColumns](tbl, {
						{ "Раздел_ID",  
							each if List.Contains ({1..36},      [Строка_ID]) = true then 1	// Раздел 1
							else if List.Contains ({50..65},     [Строка_ID]) = true then 2	// Раздел 2
							else if List.Contains ({70..70},     [Строка_ID]) = true then 3	// Раздел 3
							else if List.Contains ({80..95},     [Строка_ID]) = true then 4	// Раздел 4
							else if List.Contains ({140..171},   [Строка_ID]) = true then 5	// Раздел 5
							else if List.Contains ({1010..1320}, [Строка_ID]) = true then 6	// Раздел 6
							else if List.Contains ({1401..1411}, [Строка_ID]) = true then 7	// Раздел 7
							else 9
						, Int64.Type},
					
						{ "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья] , type text},
						
						//
						{ "Заголовок_ID", 
							each Number.From(Text.From([Раздел_ID]) & Text.From([Заголовок]))
						, Int64.Type}
					}),
			
				// 	добавляем значения колонок с ID, если переключатель IsReportViewGroup = True
					Join = 
						let Join_Заголо  = Table.Join(AddedColumns, "Заголовок_ID", Справочники[ПРИЛОЖЕНИЕ][Спр_Заголовки],   "Заголовок_ID"),
							Join_Раздел  = Table.Join(Join_Заголо,  "Раздел_ID",    Справочники[ПРИЛОЖЕНИЕ][Спр_Разделы],     "Раздел_ID")
						in  Join_Раздел,
				
				//	возвращаем
					to = Join 
			][to],

		PRIL.StepEnd = 
			(tbl as table) =>
			[ 
				to = tbl 
			][to]
	],

	RAS = [ // ОТЧЕТ О СОСТОЯНИИ РАСЧЕТОВ
		RAS.Connect =
			(tbl as table) =>
				let from = 	Table.SelectColumns(
									Table.SelectRows(tbl, each [Report_ID] = "osr1"),
									{"Company","DataSet","Report_ID","Date"} & {"Content"}
								)
				in  from,
			
		RAS.Load =
				(tbl as table) =>	
				let load = // Чистим подтаблицы = лист с отчетом. Для Гомеля "Step 1" нет, так как каждый отчет в отдельном файле. from = заменяет его
					Table.AddColumn(tbl, 
						"Step2", // создаем колонку с подтаблицами
							(x)=> 
								let 
								// фильтруем колонки
									select = Table.SelectColumns(x[Content], {"Column1", "Column2", "Column3", "Column4"} ), 
								
								// чистим текст во всей таблице (" " двойные пробелы, табуляцию,)
									clear   = F[Table.ClearAll](select," #(lf)#(00A0)",true,true),
									
								// список имен колонок подтаблицы
									ColNames = Table.ColumnNames(clear),

								// меняем тип данных второй колонки - "НомерСтроки" для дальнейшей фильтрации
									tr = Table.TransformColumns(clear,{ 
											{ ColNames{1}, (x)=> try Number.From(x) otherwise null, Int64.Type}
										} ),
							
								// фильтруем строки в колонке ПО ИМЕНИ колонки, не по [полю]
								// Table.SelectRows(RenameColumns, each {ColNames{1}} <> null), не работает!
									f=(x)=> (x <> null) and (x > 0) and (x < 1000),
									SelectRow = Table.SelectRows (tr, 
											(row) => f (Record.Field (row, ColNames{1} ) ) 
										),

								// список колонок итоговой таблицы
									lstNewColNаmes = {"Статья", "Строка_ID", "Всего", "Просроченная"},

								// меняем названия колонок кроме первой (первая Date)
									rename = Table.RenameColumns(SelectRow, List.Zip({ColNames, lstNewColNаmes}) )

							in rename
						),
					remove   = Table.RemoveColumns(load, {"Content"})
				in 	remove,
		
		RAS.Expand	=
			(tbl as table) =>
				[ 
						lst =Table.ColumnNames(tbl[Step2]{0}), // {"Статья", "Строка_ID", "Всего", "Просроченная"},

					// Разворачиваем подтаблицы. Список столбцов без первой колонки "Date", которой внутри нет
					// надежней будет колонки определять сразу до разворота  ColNames = Table.ColumnNames(from2[Step2]{0})
						EXPAND = Table.ExpandTableColumn(tbl,  "Step2", lst, lst),
					
					// меняем тип колонок с трансформацией. Можно делать это и в подтаблицах, до разворота
						Transform = Table.TransformColumns(EXPAND,
							{ 
								{ lst{0}, (x)=> x, type text},
								{ lst{1}, (x)=> try Number.From(x) otherwise null, Int64.Type},
								{ lst{2}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
								{ lst{3}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type}
							} 
						)
				] [Transform],
		
		RAS.Clean	=
			(tbl as table) =>
				[
					// Убираем текст внутри скобок и сами скобки.
					// Проверяем колонку x на наличие символ из списка. Если есть - оставляем текст ДО, а "(" и ПОСЛЕ нее все убираем
						УбираемТекстПослеСкобки = 
							Table.TransformColumns ( tbl, // можно просто убрать все в скобках!
								{ "Статья", (x) => 
									if   F[Text.CompareText]({"50","66","70","09","76","60","69","68","202"}, x) = true
									then Text.BeforeDelimiter(x, "(", {0, RelativePosition.FromEnd})    
									else x, type text } 
							),
					
					sel = Table.SelectRows(УбираемТекстПослеСкобки, each [Строка_ID] < 200),
					
					Remove = F[REPLACE][Text.RemoveMulti](sel, "Статья", 
						{" - всего", "Из стр. 101 - ", "Из стр. 110 - ", "Из стр. 101 графы 2 - ",
						 " (в отпускных ценах)"," (договоров)","из нее ","из них "})
						
				] [Remove],

		RAS.AddCol	=
			(tbl as table) =>
				let addOffset = Table.ReplaceValue(tbl, // сделать функцию. Ждобавить тип функций - проверка типа is ...
						each [Статья], 
						each if not List.ContainsAny ({Text.Start([Статья],1) },{"А".."Я"},Text.Contains) = true 
							then "    " & [Статья] 
							else [Статья], 
						Replacer.ReplaceValue,{"Статья"}
					),
					addColumn = Table.AddColumn(addOffset, "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья])
				in  addColumn,
		RAS.StepEnd	=
			(tbl as table) =>
				[ // можно метки на шагах запроса - AddDataSet meta [Expand_ = EXPAND] - Value.Metadata(А_Баланс_Загрузка)[Expand_]
					Result = tbl
				][Result]
	]

] // конец Report
in	
	func & Справочники & Report