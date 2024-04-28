let Report = 
	[ readme = "Библиотека функций для обогащения форм отчетности новыми колонками",
	
	// ---------------------------------------------------------------------------------------	
	// Добавляем колонку с "Form_ID" отчета в самое начало таблицы (ColumnName as text) 
	fxAddColumnOffset = [   
		func = (tbl as table, Value as text, Offset as number) =>
			let ColumnName  = "Form_ID",
				AddColumns  = Table.AddColumn(tbl, ColumnName, each Value),
				nms         = Table.ColumnNames(AddColumns),
				Reorder     = Table.ReorderColumns(AddColumns,
								let lst = {ColumnName}, 
									col = Offset
								in  List.InsertRange(List.Difference(nms, lst), col, lst))
			in 	Reorder,

			typ = type function (tbl as table, Value as text, Offset as number) as table 
				meta [ 	Documentation.Name            = "fxAddColumnOffset",
				  	   	Documentation.LongDescription = "добавляем колонку с 'Form_ID' отчета в самое начало таблицы"],
		  	result = Value.ReplaceType(func,typ) 
	][result],
	
	
	// ---------------------------------------------------------------------------------------	
	// Перегруппировка всех столбцов таблицы согласно нового списка lst
	fxReorderAllColumns = 
		[	func = (tbl as table, lst as list) => 
		 		Table.ReorderColumns(tbl,
					List.InsertRange(
						List.Difference(
							Table.ColumnNames(tbl),
						lst),
					0, lst)
				),
			typ = type function (tbl as table, lst as list) as table 
				meta [ 	Documentation.Name            = "fxReorderAllColumns",
				  	   	Documentation.LongDescription = "Перегруппировка всех столбцов таблицы согласно нового списка lst.
							 							 Список новых заголовков, число и названия колонок должны совпадать. <br>
														 Пример: lst = Value.Metadata(СправочникСтолбцов)[balance][Имя]" ],
		  	result = Value.ReplaceType(func,typ) 
		][result],


	// ---------------------------------------------------------------------------------------	
	// делаем отступ из двух пробелом перед первым символом если номер строки не кратный 10 (последний символ <> 0)
	// могут быть исключения. Например 411 - статья "Паевой фонд". Это доп колонка в отчетах ПК
	fxOffsetColumn = 
		[   func = (tbl as table) => 
				Table.ReplaceValue(tbl, 
					each [Статья], 
					each if Number.Mod([Строка_ID], 10) <> 0  
						then "    " & [Статья] 
						else [Статья], 
					Replacer.ReplaceValue, {"Статья"}
				),

			typ = type function (tbl as table) as table 
				meta [ 	Documentation.Name            = "fxOffsetColumn",
				  	   	Documentation.LongDescription = "Делаем отступ из двух пробелом перед первым символом если номер строки не кратный 10 (последний символ <> 0)" ],
		  	result = Value.ReplaceType(func,typ) 
		][result],

	// -----------------------------------------------------------------------------------------------------------------
    // fxJoinSpr
	// добавляем колонки в таблицу через объединение таблиц по ID_Форма
    // to = fxJoinSpr(addColumns1, Справочники[Спр_ФормыОтчетности], "КодБКС", {"Название", "Report_ID"}),
	// -----------------------------------------------------------------------------------------------------------------
	fxJoinSpr =
		[   func = (tblIn as table, 		//
					tblSpr as table, 		//
					columnName as text, 	//
					lstLoadColumn as list) 	//
			as table => 
		
			let Sprav  = Справочники[Спр_ФормыОтчетности],
				Join   = Table.NestedJoin(tblIn, columnName, tblSpr, columnName,"tmp"),
				Expand = Table.ExpandTableColumn(Join, "tmp", lstLoadColumn)
			in  Expand,

		typ = type function (tblIn as table, tblSpr as table, columnName as text, lstLoadColumn as list) as table 
				meta [ 	Documentation.Name            = "",
				  	   	Documentation.LongDescription = "добавляем колонки в таблицу через объединение таблиц по ID_Форма" ],
		  	result = Value.ReplaceType(func,typ) 
		][result],
    
		// -----------------------------------------------------------------------------------------------------------------
		// # fxReplaceTextBetweenDelimiters - удаляет текст между разделителями (скобки), вместе с ними
		// работает если вхождений несколько. Не проверял!
	    // -----------------------------------------------------------------------------------------------------------------
		fxReplaceTextBetweenDelimiters = (txt as text) =>
        [ 	fnRemoveFirstTag = (DELIM as text)=>
				let	OpeningTag = Text.PositionOf(DELIM,"("),
					ClosingTag = Text.PositionOf(DELIM,")"),
					Output 	   = if OpeningTag = -1 
								 then DELIM 
								 else Text.RemoveRange(DELIM,OpeningTag,ClosingTag-OpeningTag+1)
				in Output,

			fnRemoveDELIM = (y as text)=>
				if fnRemoveFirstTag(y) = y
				then y 
				else @fnRemoveDELIM(fnRemoveFirstTag(y)),
			
			Output = @fnRemoveDELIM(txt),
			
			typ = type function (txt as text) as table 
				meta [ 	Documentation.Name            = "",
				  	   	Documentation.LongDescription = "удалает текст межлу разделителями" ],
		  	
			result = Value.ReplaceType(fnRemoveDELIM,typ) // не работает. К какому шагу присваивать тип?!
        ][result],



// =====================================================================================================================
// 		БУХГАЛТЕРСАКИЙ БАЛАНС
// =====================================================================================================================

    // -----------------------------------------------------------------------------------------------------------------
	// обогащаем баланес дополнительными колонками по коду строки
	// fx_Balance_AddColID
	fxAddColBalanceID = (tbl as table)=>
		let 
			// добавляем колонки с цифровым ID атрибутов
			AddColumns = fxPackGlobal[fxAddMultiColumns](tbl, {
				// статьи с номерами строк от 100 до 399 соответствуют Части 1 (Активы), другие Части 2 (Пассивы)
				{ "Часть_ID",  each 
						if List.Contains ({110..399}, [Строка_ID]) = true 
						then 1 
						else 2, Int64.Type},
				
				// номер раздела соответствует первому символу номера строки. 0 - не вошедшие в разделы
				{ "Раздел_ID", each 
							let fxRazdel = (x as text) => Record.FieldOrDefault([1 = 1, 2 = 2, 4 = 3, 5 = 4, 6 = 5, 0 = 0], x, null),
								Номер = Text.Start(Text.From([Строка_ID]),1)    
							in fxRazdel(Номер)
					, Int64.Type},
				
				// если первая буква большая, то группа (1), если маленькая, то подгруппа (2)
				{ "Группа_ID", each 
						if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1)) = true 
						then 1 
						else 2, Int64.Type},
				
				// уровень иерархии: 1 - Часть, 2 - Раздел, 3 - Статья (равно 0 в конце № строки), 4 - Подстрока (не равно 0 в конце № строки)
				{ "Уровень_ID", each if List.Contains ({300, 700},[Строка_ID]) = true then 1 
								else if List.Contains ({190, 290, 390, 490, 590, 690},[Строка_ID]) = true then 2 
								else if Number.Mod([Строка_ID], 10) <> 0  and 
										not List.Contains ({411},[Строка_ID]) = true then 4 // исключение для строки 411 "Паевые взносы"
								else 3, Int64.Type}
				})
		in AddColumns,

    // -----------------------------------------------------------------------------------------------------------------
	// Добавляем столбцы для Анализа ликвидности
	// fx_Balance_AddColLiq
	fxAddColBalanceLiq = (tbl as table)=>
		let 
			// добавляем колонки с цифровым ID атрибутов
			AddColumn = fxPackGlobal[fxAddMultiColumns](tbl, {
			
				// выбираем строки, которые НЕ включаются в расчет = false
				{ "Ликвидность_Статус", each 
					if List.Contains ({190,290,390,490,590,690,300,700,630,211,212,213,214,215,216},[Строка_ID]) = true 
					then false 
					else true, type logical},
				
				// те что включаются делим по категориям 4 + 4 
				{"Ликвидность_Группа",each if List.Contains ({260, 270},[Строка_ID]) = true then "А1"
								else if List.Contains ({210, 250, 280},[Строка_ID]) = true then "А2"
								else if List.Contains ({220, 230, 240, 150, 170},[Строка_ID]) = true then "А3"
								else if List.Contains ({110, 120, 130, 140, 160, 180},[Строка_ID]) = true then "А4"
								else if List.Contains ({632, 633, 634, 635, 636, 637, 638, 639},[Строка_ID]) = true then "П1" 
								else if List.Contains ({610, 620, 640, 650, 660, 670, 631},[Строка_ID]) = true then "П2" 
								else if List.Contains ({510, 520, 530, 540, 550, 560},[Строка_ID]) = true then "П3"
								else if List.Contains ({410, 420, 430, 440, 450, 460, 470, 480},[Строка_ID]) = true then "П4" 
								else "0", type text}, // мб сделать null ?

				// КатегорияЛиквидности
				{"Ликвидность_Категория",
								each if List.Contains ({260, 270} & {632, 633, 634, 635, 636, 637, 638, 639},[Строка_ID]) = true then "АП1"
								else if List.Contains ({210, 250, 280} & {610, 620, 640, 650, 660, 670, 631},[Строка_ID]) = true then "АП2"
								else if List.Contains ({220, 230, 240, 150, 170} & {510, 520, 530, 540, 550, 560},[Строка_ID]) = true then "АП3"
								else if List.Contains ({110, 120, 130, 140, 160, 180} & {410, 420, 430, 440, 450, 460, 470, 480} ,[Строка_ID]) = true then "АП4"
								else "0", type text}
			})
		in AddColumn,


    // -----------------------------------------------------------------------------------------------------------------
	// ОБОГАЩЕНИЕ БАЛАНСА, после того, как привели его с нормальный вид, сделав один массив данных и 4 столбца
	// fx_Balance_StepEnd
	fxBalance_StepEnd = (tbl as table) =>
		let 
			// список начальных колонок итоговой таблицы: общие + обязательные (форму еще не добавили)
				lst1 = {"Company", "DataSet", "Report_ID", "Date"},
				lst2 = {"Статья", "Строка_ID", "На начало", "На конец"},

			/*  ТРАНСФОРМАЦИЯ ---------------------------------------------------------------------------------------*/
				// Разворачиваем подтаблицы. Список столбцов без первой колонки "Date", которой внутри нет
				// надежней будет колонки определять сразу до разворота  ColNames = Table.ColumnNames(from2[Step2]{0})
				EXPAND = Table.ExpandTableColumn(tbl, "Step2", lst2, lst2),
				
				// меняем тип колонок с трансформацией. Можно делать это и в подтаблицах, до разворота
				//  - 3 дополнительные колонки: Название организации + Номер датасета + Вид ответа + Дата отчета)
				//  - 4 обязательные колонки:   Статья + Строка + На начало + На конец)
				Transform = Table.TransformColumns(
						EXPAND,{ 
							{ lst1{0}, (x)=> x, type text},               	// организация
							{ lst1{1}, (x)=> Number.From(x), Int64.Type},   // номер датасета
							{ lst1{2}, (x)=> "balance", type text},         // отчет - Сделать join со справочником отчетов
							{ lst1{3}, (x)=> Date.From(x), Date.Type},      // дата отчета
							
							{ lst2{0}, (x)=> x, type text}, 													// статья
							{ lst2{1}, (x)=> try Number.From(x) otherwise null, Int64.Type}, 					// строка
							{ lst2{2}, (x)=> if Number.From(x) = 0 then null else Number.From(x), Int64.Type}, 	// на начало
							{ lst2{3}, (x)=> if Number.From(x) = 0 then null else Number.From(x), Int64.Type}  	// на конец
						} ),
		
			/* ОБОГАЩЕНИЕ -----------------------------------------------------------------------------------------*/

				// Обязательно должны быть 4 стандартные колонки: Статья, Строка, Сумма на начало, Сумма на конец
				//AddColumns0 = fxPackForm[fxAddColumnOffset] (Transform, "balance", 0),   // добавляем колонку с номером отчета в начало таблицы
				AddColumns1 = fxPackForm[fxAddColBalanceID] (Transform),               // добавляем колонки с ID (Часть, Раздел, Группа, Статья, Уровень)
				AddColumns2 = fxPackForm[fxAddColBalanceLiq] (AddColumns1),            // добавляем колонки для анализа ликвидности (Статус, Категория, Группа)

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
				Join = [ Join_Часть   = Table.Join(AddColumns2,    "Часть_ID",     Справочники[БАЛАНС][Спр_ББ_Часть],   "Часть_ID"),
						Join_Раздел  = Table.Join(Join_Часть ,     "Раздел_ID",    Справочники[БАЛАНС][Спр_ББ_Раздел],  "Раздел_ID"),
						Join_Группа  = Table.Join(Join_Раздел,     "Группа_ID",    Справочники[БАЛАНС][Спр_ББ_Группа],  "Группа_ID"),
						Join_Уровень = Table.Join(Join_Группа,     "Уровень_ID",   Справочники[БАЛАНС][Спр_ББ_Уровень], "Уровень_ID"),
						Join_Строка  = Table.Join(Join_Уровень,    "Строка_ID",    Спр_ББ_Статья,                       "Строка_ID")
					][Join_Строка],
				
			/* НАВОДИМ КРАСОТУ ---------------------------------------------------------------------------------*/ 

				// добавляет отступ в колонке Статья
					OffsetChar = fxPackForm[fxOffsetColumn](Join),

				// перегруппировка колонок согласно кастомного списка в Excel
					ReorderColumns = fxPackForm[fxReorderAllColumns]( 
						OffsetChar, 
						Value.Metadata(СправочникСтолбцов)[balance_listcol][Имя] 
					),

			/* РЕЗУЛЬТАТ ---------------------------------------------------------------------------------------*/ 

				// выводим или узкую таблицу или с допколонками со значениями групп ID "все в одном"
				_Result_Step2 = if IsReportViewGroup 
								then ReorderColumns 
								else AddColumns2,
				
				// ставим метки на шагах запроса (потом использовать to = Value.Metadata(А_Баланс_Загрузка)[Expand_]_
				Result = _Result_Step2  meta [Expand_ = EXPAND]
		in 
			Result,


	
// ==========================================================================================================
//		ОТЧЕТ О ПРИБЫЛЯХ И УБЫТКАХ
// ==========================================================================================================
    // -----------------------------------------------------------------------------------------------------------------
	// добавляем дополнительные колонки
	// fx_OPU_AddCol
	fxAddCol_OPU = 
		[   func = (tbl as table) =>
			fxPackGlobal[fxAddMultiColumns](tbl, {
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
					else 3, Int64.Type}
				}),

			typ = type function (tbl as table) as table 
				meta [ 	Documentation.Name            = "fxAddCol_OPU",
				  	   	Documentation.LongDescription = "добавляем колонки Группа_ID + Прибыль_ID + ВитдДеят_ID - тип данных Int64.Type" ],
		  	result = Value.ReplaceType(func,typ) 
		][result],

	
    // -----------------------------------------------------------------------------------------------------------------
	// 
	// fx_OPU_StepEnd
	fxStepEnd_OPU = 
		[   func = (tbl as table) =>
		let 
			/*  ШАГ 3 - ОБРАБОТКА загруженных данных  */ /* ОТСЮДА МОЖНО ЗАГОНЯТЬ В ГЛОБАЛЬНУЮ ФУНКЦИЮ */
			
			// разворачиваем подтаблицы
			lstColumnNew = {"Статья", "Строка_ID", "Текущий период", "Прошлый период"},
			Expand = Table.ExpandTableColumn(tbl, "Step2", lstColumnNew, lstColumnNew),
			
			// трансформируем столбцы с указанием типа
			lst1 = {"Company", "DataSet", "Report_ID", "Date"},
			TransformAllColumns = Table.TransformColumns(Expand,{ 
				{ lst1{0}, (x)=> x, type text},               	// организация
				{ lst1{1}, (x)=> Number.From(x), Int64.Type},   // датасет
				{ lst1{2}, (x)=> "opu", type text},         	// отчет Сделать join со справончиком отчетов
				{ lst1{3}, (x)=> Date.From(x), Date.Type},      // дата
				
				{"Статья",         (x)=> x, Text.Type},
				{"Строка_ID",      each try Number.From(_) otherwise null, type number},
				{"Текущий период", each try Number.From(_) otherwise _, Number.Type},
				{"Прошлый период", each try Number.From(_) otherwise _, Number.Type}
			} ),
			
			// Меняем 0 на null. Чтобы в визуализации не появлялся.
			Replace0ToNull = Table.ReplaceValue(TransformAllColumns, 
				0, null, Replacer.ReplaceValue, 
				{"Текущий период", "Прошлый период"}
			),

			// Убираем текст внутри скобок и сами скобки! убираем и после скобок, а это не надо исправить надо
    		УбираемТекстПослеСкобки = 
				let f=(x)=> 	    				// if Text.Contains (x, "(") = true
					[	lst = {"010","100","±"},	// на наличие каких символов проверяем строку
						chr = "(",				    // после какого символа влючая его убираем текст
						to = if  fxPackGlobal[fxCompareGroups](lst, x) = true     			 // проверяем колонку x на наличие символ из списка
							then Text.BeforeDelimiter(x, chr, {0, RelativePosition.FromEnd}) // если есть - оставляем текст ДО, а "(" и ПОСЛЕ нее все убираем
							else x
					] [to],
					to = Table.TransformColumns( Replace0ToNull, {"Статья", f, type text} )
				in  to,

			// делаем отступ, если номер строки не четный
			ДобавляемОтступ = fxPackForm[fxOffsetColumn](УбираемТекстПослеСкобки),
			
			// AddColumns1 = fxPackForm[fxAddColumnOffset] (ДобавляемОтступ, "opu", 0),   // добавляем колонку с номером отчета в начало таблицы
			AddColumns2 = fxPackForm[fxAddCol_OPU] (ДобавляемОтступ),                      // добавляем 3 колонки с ID 
			AddColumnY  = Table.AddColumn(AddColumns2, "Год", each Date.Year([Date]), Int64.Type), // добавляем колонку с номером года
			
		/* СПРАВОЧНИКИ */
		
			// справочник статей
			Спр_ОПУ_Статья = 
				let
					Источник        = Table.Distinct(Expand[[Строка_ID],[Статья]], {"Строка_ID"}),
					AddCol_СтрокаID = Table.AddColumn(Источник, "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья]),
					RenameColumns   = Table.RenameColumns(AddCol_СтрокаID,{{"Статья", "СтатьяИмя"}}),
					
					// Делаем множественную замену в подстроке столбца
					МультиЗамена = [ 
						dict = List.Buffer(Table.FromValue({ 
								{"продукции, товаров, работ, услуг",    ""},
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
					AddIndexColumn = Table.AddIndexColumn(МультиЗамена, "СтрокаИндекс", 1, 1, Int64.Type)
				in  AddIndexColumn,
			
			
			/* JOIN */
			
			// обединяем таблицу фактов со справочниками с формированием дополнительных колонок-срезов
			Join = 
				let Join_ВидДеят    = Table.Join(AddColumnY,  	  "ВидДеят_ID", Справочники[ОПУ][Спр_ОПУ_ВидДеятельности],  "ВидДеятельности_ID"),
					Join_ВидПрибыли = Table.Join(Join_ВидДеят,    "Прибыль_ID", Справочники[ОПУ][Спр_ОПУ_ВидПрибыли],       "Прибыль_ID"),
					Join_Группа     = Table.Join(Join_ВидПрибыли, "Группа_ID",  Справочники[ОПУ][Спр_ОПУ_Группа],           "Группа_ID"),
					Join_Строка     = Table.Join(Join_Группа,     "Строка_ID",  Спр_ОПУ_Статья,"Строка_ID")
				in  Join_Строка,
			
			/* ПЕРЕГРУППИРОВКА ---------------------------------------------------------------------------------*/ 
			
			// перегруппировка колонок: (DataSet, Date  + Статья, Строка, Текущий, Прошлый) + Остальные по алфавиту
			// если так, Reorder = Table.ReorderColumns(Join, {"DataSet", "Статья", "ID_Строка", "Текущий период", "Прошлый период"} ),то  эти вперед, остальные как есть
			ReorderColumns = 
				fxPackForm[fxReorderAllColumns](
					Join, 
					Value.Metadata(СправочникСтолбцов)[opu_listcol][Имя] 
				),

			// выводим: true - узкую таблицу, fase - таблицу ""все в одном" с допколонками со значениями групп
			result = if IsReportViewGroup then ReorderColumns else AddColumnY
		in	result,

		// документирование функции
			typ = type function (tbl as table) as table 
					meta [ 	Documentation.Name            = "///",
							Documentation.LongDescription = "///" ],
		  	result = Value.ReplaceType(func,typ) 
		][result],

		
// =====================================================================================================================
//	ОТЧЕТ О ДВИЖЕНИИ ДЕНЕЖНЫХ СРЕДСТВ
// =====================================================================================================================
    // -----------------------------------------------------------------------------------------------------------------
	//  
	// fx_ODDS_StepEnd
	fxStepEnd_ODDS = [ 
		func = (tbl as table) =>
		let 
		/*  РАЗВОРАЧИВАЕМ ----------------------------------------------------------------------------------------*/
			// разворачиваем подтаблицы
			lstColumnNew = {"Статья", "Строка_ID", "Текущий период", "Прошлый период"},
			Expand = Table.ExpandTableColumn(tbl, "Step2", lstColumnNew, lstColumnNew ),
			
			// трансформируем столбцы с указанием типа
			TransformAllColumns = Table.TransformColumns(Expand,{ 
				{ "Company", 	(x)=> x, type text},               	// организация
				{ "DataSet", 	(x)=> Number.From(x), Int64.Type},  // датасет
				{ "Report_ID", 	(x)=> "odds", type text},         	// отчет Сделать join со справончиком отчетов
				{ "Date", 		(x)=> Date.From(x), Date.Type},     // дата
				
				{"Статья",         (x)=> x, Text.Type},
				{"Строка_ID",      each try Number.From(_) otherwise null, type number},
				{"Текущий период", each try Number.From(_) otherwise _, Number.Type},
				{"Прошлый период", each try Number.From(_) otherwise _, Number.Type}
			} ),
			
			// меняем 0 на null (чтобы в визуализации не появлялся)
			Replace0ToNull = Table.ReplaceValue(TransformAllColumns, 
				0, null, Replacer.ReplaceValue, 
				{"Текущий период", "Прошлый период"}
			),
		
		/*  ОБОГАЩЕНИЕ ----------------------------------------------------------------------------------------*/
		
			// добавляем колонки
			AddColumns1 = fxPackGlobal[fxAddMultiColumns](Replace0ToNull, {
				{ "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья] },
				
				{ "Группа_ID",  each 
						if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1)) = true 
						then 1 
						else 2 , Int64.Type}, // 1- Группа, 2 - Подгруппа
				
				{ "ВидДеят_ID", each if List.Contains ({20..40},[Строка_ID]) = true then 1 // "Текущая"
								else if List.Contains ({50..70},[Строка_ID]) = true then 2 // "Инвестиционная"
								else if List.Contains ({80..100},[Строка_ID]) = true then 3 // "Финансовая"
								else 4, Int64.Type}, // "Другое"
				
				{ "ВидДвиж_ID", each if List.Contains ({20..24, 50..55, 80..84},[Строка_ID]) = true then "Поступление" 
								else if List.Contains ({30..34, 60..64, 90..95},[Строка_ID]) = true then "Выбытие" 
								else if List.Contains ({40},[Строка_ID]) = true then "РезультатТек" 
								else if List.Contains ({70},[Строка_ID]) = true then "РезультатИнв" 
								else if List.Contains ({100},[Строка_ID]) = true then "РезультатФин" 
								else if List.Contains ({110},[Строка_ID]) = true then "РезультатОбщий" 
								else "Другое", type text},
				
				{"Год", each Date.Year([Date]), Int64.Type}                        
			}),
			AddColumns2 = fxPackForm[fxAddColumnOffset] (AddColumns1, "odds", 0),
				
		/*  СПРАВОЧНИКИ -----------------------------------------------------------------------------------------*/

			// справочник
			Спр_ОДДС_Статья = 
				let
					Источник      = Table.Distinct(Replace0ToNull[[Строка_ID],[Статья]],{"Строка_ID"}),
					
					// добавляем отступ к строкам подгруппы
					ReplaceValue  = fxPackForm[fxOffsetColumn](Источник),
					
					ДобавляетмОтступ = Table.AddColumn(ReplaceValue, "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья]),

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

		/*  КОЛОНКИ ---------------------------------------------------------------------------------*/ 
			
		// добавляет отступ в колонке Статья
			OffsetChar = fxPackForm[fxOffsetColumn](AddColumns2),
		
		// перегруппировка колонок согласно кастомного списка в Excel
			ReorderColumns = 
				fxPackForm[fxReorderAllColumns]( 
					OffsetChar, 
					Value.Metadata(СправочникСтолбцов)[odds_listcol][Имя] 
				),

		/*  РЕЗУЛЬТАТ  ---------------------------------------------------------------------------------*/ 
			
			// выводим: true - узкую таблицу (шаг AddColumn_Year), fase - таблицу ""все в одном" с допколонками (шаг Отступ)
			RESULT = if IsReportViewGroup then ReorderColumns else Replace0ToNull
		in	RESULT,

			typ = type function (tbl as table) as table 
				meta [ 	Documentation.Name            = "///",
				  	   	Documentation.LongDescription = "///" ],
		  	result = Value.ReplaceType(func,typ) 
		][result]

    // -----------------------------------------------------------------------------------------------------------------
	// 

    // -----------------------------------------------------------------------------------------------------------------
	// 

// =====================================================================================================================
//	ОТЧЕТ О ИЗМЕНЕНИИ СОБСТВЕННОГО КАПИТАЛА
// =====================================================================================================================
    // -----------------------------------------------------------------------------------------------------------------
	// 

    // -----------------------------------------------------------------------------------------------------------------
	// 

    // -----------------------------------------------------------------------------------------------------------------
	// 
    


] in Report
