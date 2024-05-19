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
			)
	],

	// выводим только список функций, без других утилит
	result = Biblio 
 in result 