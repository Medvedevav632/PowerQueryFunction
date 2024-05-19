// ! -------------------------------------------------------------------------------------
a = Record.FieldValues(Source2{x}), //*получаем значения текущей строки
     b = Table.ColumnNames(Source2), //* получаем названия столбцов таблицы
    c = #table(2,List.Zip({b,a})), //* создаем таблицу

= Table.FromList(
                    Table.ToRows(КолонкуВначало),
                    (x)=>{Text.Combine(List.Transform(List.LastN(x,3),Text.From),".")} & List.Skip(x), 
                    // лучше отдельным шагом через f=(x)=>{Text.Combine...
                    {"Номер"} & List.Skip(Table.ColumnNames(КолонкуВначало))
            )

= Table.ReplaceValue(#"Измененный тип", 
                    null, // b
                    null, // c
                    (a, b, c) => if List.Contains("в т.ч.", a) = true then "" else a,
                    {"Товар"})

= Table.TransformColumns(
    #"Измененный тип", 
        {{"Товар", 
                   each List.ReplaceMatchingItems(
                          {_}, 
                          List.Buffer(List.Zip({ {"в т.ч","130"},{"+++","190"} }))
                   ){0}}} 
    )
// ! -------------------------------------------------------------------------------------
fxCamelCaseHeaders =[
    // Функция fxCamelCaseHeaders - поставить пробел перед каждой большой буквой
    // https://datachant-com.translate.goog/2018/02/14/split-camelcase-headers-m/?_x_tr_sl=auto&_x_tr_tl=ru&_x_tr_hl=ru
    // модно использовать для заголовков или самого текста
    (Source)=>
        let to_list = List.Buffer(Text.ToList(Source)),
            Transform = List.Transform(to_list, each if Text.Lower(_) = _ then _ else " " & _),
            Reassemble = Text.Combine(Transform)
        in  Reassemble	
	]

fxSplitCamelCase
    // если не нужны пробелы когда аббревиатура, то функция другая
    // to = Table.TransformColumnNames(#"Previous Step", fxSplitCamelCase)
    // Функция - fxSplitCamelCase
    (Source) =>
    let    
        res = List.Accumulate(
            Text.ToList(Source),
            [result="", index=0, source=Source],
            fnAccumulator
        ),

    IsUpper = (txt) => txt <> "" and txt = Text.Upper(txt) and txt <> Text.Lower(txt),

    fnAccumulator =
        (state as record, current as text) as record =>
        let
            prevCharacter  = if state[index]=0 then "" else Text.At(state[source], state[index] - 1),
            prevCharacter2 = if state[index]<=1 then "" else Text.At(state[source], state[index] - 2),
            nextCharacter  = if state[index]= Text.Length(state[source]) - 1 then "" else Text.At(state[source], state[index] + 1),
          aggregatedResult = if state[index]=0 or current = "" then current
                else
                    if  IsUpper(current) and
                        (not IsUpper(prevCharacter) or 
                         (IsUpper(prevCharacter2) and not IsUpper(nextCharacter))) then
                    state[result] & " " & current
                else
                    state[result] & current,
            
            resultRecord = [result = aggregatedResult, index = state[index]+1, source = state[source]]

        in
            resultRecord
    in
        res[result]
	
// Функция - fxSplitCamelCase вариант2
    // Shared by Sergey Lossev
    (Source) =>
    let
        to_list     = Text.ToList(Source),
        accum_chars = List.Accumulate(to_list, "", (sum, x) => sum & (if x = Text.Upper(x) then " " & x else x)),
        split 		= Text.Split(Text.Trim(accum_chars), " "),
        accum_words = List.Accumulate(split, "", (sum, x) => sum & (if Text.Length(x) = 1 then x else " " & x & " ")),
        replace 	= Text.Trim(Text.Replace(accum_words, "  ", " "))
    in
        replace
	
//  Finally
    /*  FnSplitCamelCase is a very handly custom function. It can also help you to handle CamelCase values in 
	columns – not just headers. For example, if you have CamelCase values in Column1, 
	you can click fx after Previous Step in Applied Steps, and paste the following formula:*/
	to = Table.TransformColumns(#"Previous Step",{{"Column1", fxSplitCamelCase, type text}})


	// ---------------------------------------------------------------------------------------	
	// Table_TransformColumnTypes - изменение типа всех столбцов (Чисто, Дата, Текст все остальные)
		let
			from = Table.FromRecords({[Наименование="тапки",Категория="обувь",ЦенаОпт=1000,ЦенаРозн=2000,Количество=100,СуммаОпт=100000,СуммаРозн=200000,Продавец="Вася",ДатаПоступления=44562,ДатаПродажи=44835,ДатаОтгрузки=44866,Примечание="обычные тапки"]}),
			nms  = Table.ColumnNames(from),
			dlst = List.Select(nms,(i)=>Text.Contains(i,"Дата")),
			nlst = List.Select(nms,(i)=>List.Contains({"Цена","Количество","Сумма"},i,(x,y)=>Text.Contains(y,x))),
			tlst = List.Difference(nms,dlst&nlst),
			f=(x,y)=>List.Transform(x,(i)=>{i,y}),
			transflst = f(dlst,Date.Type)&f(nlst,Number.Type)&f(tlst,Text.Type),
			to = Table.TransformColumnTypes(from,transflst)
		in
			to

// Телеграмм. Чат по PQ
    // Файл Якушев (есть в папке PQ)
    // Ключевые слова: видимо не проще... типа того (доработка напильником на совести заказчика):
        let from = Excel.Workbook(File.Contents("C:\Users\muzyk\Desktop\якушев.xlsx"), null, true){0}[Data],
            nms1 = List.Buffer(List.RemoveMatchingItems(Record.FieldValues(from{1}),{"",null})),
            nms2 = List.Buffer(Record.FieldValues(from{2})),
            pos  = List.PositionOf(nms2,"Артикул",Occurrence.All),
            st   = pos{0},
            len  = pos{1}-pos{0},
            num  = len*List.Count(pos),
            nms3 = List.ReplaceRange(nms2,st,num,{"код"}&List.Transform(List.Range(nms2,pos{0},pos{1}-pos{0}),(x)=>"new_"&x)),
            g=(x)=>[a = List.ReplaceMatchingItems(List.Range(x,st,num),{{"",null}}),
                    d = Number.RoundDown(List.PositionOf(a,null,Occurrence.First,(x,y)=>x<>y)/len),
                    b = if d>=0 then {nms1{d}}&List.Range(a,d*len,len) else List.Repeat({null},len+1), 
                    c = List.ReplaceRange(x,st,num,b) ][c],
            lst = Table.ToList(Table.Range(from,3),g),
            to  = Table.FromList(lst,(x)=>x,nms3)
        in  to,
    
функцияЗамены = [
    let func =   
    (
    TableWithTextColumn as table, 
    ColumnNameTextToReplace as text, 
    ReplacementsTable as table, 
    ColumnNameWordToReplace as text, 
    ColumnNameReplaceWith as text
    ) => 
    let
        Source = ReplacementsTable,
        CreateListOfLists = Table.AddColumn(
            Source, 
            "Liste", 
            each ({Record.Field(_, ColumnNameWordToReplace), Record.Field(_, ColumnNameReplaceWith)})
        ),
        Text = TableWithTextColumn,
        TurnTextToList = Table.AddColumn(
            Text, 
            "Custom", 
            each Text.Split(Record.Field(_, ColumnNameTextToReplace), " ")
        ),
        Replacements = Table.AddColumn(
            TurnTextToList, 
            "Changed Text Expected", 
            each Text.Combine(List.ReplaceMatchingItems([Custom], CreateListOfLists[Liste]), " ")
        ),
        Cleanup = Table.RemoveColumns(Replacements, {"Custom"})
    in
        Cleanup ,
    documentation = [
    Documentation.Name =  " Table.ReplaceMultiple.pq ",
    Documentation.Description = " Replaces multiple values at a time in a table column without recursion. ",
    Documentation.LongDescription = "  Replaces multiple values at a time in a table column without recursion: http://www.thebiccountant.com/2016/05/22/multiple-replacements-in-power-bi-and-power-query/ The <code>ReplacementsTable </code> must have the values to be replaced in the 1st column and the new value in the 2nd. ",
    Documentation.Category = " Table.ColumnOperations ",
    Documentation.Source = " http://www.thebiccountant.com/2016/05/22/multiple-replacements-in-power-bi-and-power-query/ . ",
    Documentation.Version = " 2.0 ",
    Documentation.Author = " Imke Feldmann ",
    Documentation.Examples = {[Description =  "  ",
    Code = " let 
        TableWithText = #table( {""Text""}, List.Zip( { {""the cat sat on the mat"" ,""the cat sat next to the dog"" ,""the dog chased the cat"" ,""the dog sat on the mat"" ,""the catamaran sails through the air""} } ) ), 
        ReplacementsTable = #table( {""Word To Replace"", ""Replace With""}, List.Zip( { {""cat"" ,""mat"" ,""dog"" ,""the"" ,""air""}, {""bear"" ,""chair"" ,""dragon"" ,""THE"" ,""water""} } ) ),
    // The query containing the function has to be called ""fnTable_ReplaceMultiple_pq""
        Result = fnTable_ReplaceMultiple_pq ( TableWithText , ""Text"", ReplacementsTable, ""Word To Replace"", ""Replace With"" ) 
    in 
        Result ",
    Result = " Changed Text Expected
    ———————
    THE bear sat on THE chair
    THE bear sat next to THE dragon
    THE dragon chased THE bear
    THE dragon sat on THE chair
    THE catamaran sails through THE water "]}]
    
    in  
    Value.ReplaceType(func, Value.ReplaceMetadata(Value.Type(func), documentation)) 
 ],
test = [ // замены https://www.thebiccountant.com/2016/05/22/multiple-replacements-in-power-bi-and-power-query/
            let
            //Get table of word replacements
            Replacements = Excel.CurrentWorkbook(){[Name="Replacements"]}[Content],
            
            //Get table containing text to change
            TextToChange = Excel.CurrentWorkbook(){[Name="Text"]}[Content],
            
            //Get list of words to replace
            WordsToReplace = List.Buffer(Replacements[Word To Replace]),
            
            //Get list of words to replace them with
            WordsToReplaceWith = List.Buffer(Replacements[Replace With]),
            
            //A non-recursive function to do the replacements
            ReplacementFunction = (InputText)=> 
            let
            //Use List.Generate() to do the replacements
            DoReplacement = List.Generate(
                                ()=> [Counter=0, MyText=InputText], 
                                each [Counter]<=List.Count(WordsToReplaceWith), 
                                each [Counter=[Counter]+1, 
                                        MyText=Text.Replace(
                                                [MyText], 
                                                WordsToReplace{[Counter]}, 
                                                WordsToReplaceWith{[Counter]})], 
                                each [MyText]),
            //Return the last item in the list that
            //List.Generate() returns
            GetLastValue = List.Last(DoReplacement)
            in
            GetLastValue,
            //Add a calculated column to call the function on every row in the table
            //containing the text to change
            Output = Table.AddColumn(TextToChange, "Changed Text List.Generate", each ReplacementFunction([Text])),
            #"Removed Columns" = Table.RemoveColumns(Output,{"Text"})
        in
            #"Removed Columns"
]		



	(tbl as table) =>
			[
			    //	список столбцов (пропускаем первые 4 технические столбца)
					nms = List.Skip(Table.ColumnNames(tbl), 4),

				// 	типизируем таблицу
					transform = Table.TransformColumns(tbl,
						{ 
							{ nms{0}, (x)=> F[TEXT][Text.ReplaceBetweenDelimiters](x,"()",")"), type text},
							{ nms{1}, (x)=> try Number.From(x) otherwise null, Int64.Type },
							{ nms{2}, (x)=> if  Number.From(x) = 0 then null else Number.From(x), Int64.Type},
							{ nms{3}, (x)=> if  Number.From(x) = 0 then null else Number.From(x), Int64.Type}
						} 
					),
				
				// 	удаляем лишний текст
					// removeList  = {"В том числе: ", "В том числе:", " - всего"},
					// multiRemove = F[REPLACE][Text.RemoveMulti](Replace0ToNull,"Статья", removeList),
								
				// 	делаем отступ для подстатей
					// addOffset = F[fxOffsetTextInColumn](transform, "Статья", "Строка_ID", 10),
				
				to = nmstransform
			][to],



            let

/* ШАГ 0 - Модуль загрузки функций Load ---------------------------------------------------------------------*/

//  СПРАВОЧНИКИ
    // Названия разделов отчета Приложение в Балансу
    Спр_Разделы = 
        let Источник = #table(type table 
            [   Раздел_ID     = number, 
                РазделПрефикс = text, 
                РазделИмя     = text, 
                Str1          = number,
                Str2          = number],
            {   {1, "Раздел 1", "Раздел I. Основные средства и нематериальные активы", 1, 36}, 
                // {2, "Раздел 1", "Справочно. ", 1, 5}, 
                // {3, "Раздел 1", "Источники финансирования сложений ", 30, 36}, 
                {2, "Раздел 2", "Раздел II. Незавершенные вложения во внеоборотные активы", 50, 65},
                {3, "Раздел 3", "Раздел III. Резервы по сомнительным долгам", 70, 70},
                {4, "Раздел 4", "Раздел IV. Финансовые вложения", 80, 95},
                {5, "Раздел 5", "Раздел V. Полученные кредиты и займы", 140, 171},
                {6, "Раздел 6", "Раздел VI. Справка о платежах в бюджет и ФСЗН", 1010, 1320},
                {7, "Раздел 7", "Раздел VII. Забалансовые счета", 1401, 1411}
            }),

            AddCol  = Table.AddColumn(Источник, "РазделИмяПрефикс", each [РазделПрефикс] & " " & [РазделИмя]),
            Reorder = Table.ReorderColumns(AddCol,
					let nms = Table.ColumnNames(AddCol), lst = {"РазделИмяПрефикс"}, col = 3  // 0 - если в начало
					in  List.InsertRange(List.Difference(nms, lst ), col, lst) ),
            SelCol = Table.SelectColumns(Reorder, List.FirstN(Table.ColumnNames(Reorder), 4) )
    in  SelCol,

    // заголовки разделов
    Спр_Заголовки =
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
            ExpandList = Table.ExpandListColumn(Table.TransformColumns(Источник, 
                            {{"НомерРаздела", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), 
                            let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}}), 
                            "НомерРаздела"),
            
            // объединяем два столбца в новый используя функцию
            CombineCol = Table.CombineColumns(ExpandList, 
                            {"НомерРаздела", "IndexIn"},
                            Combiner.CombineTextByDelimiter("", QuoteStyle.None),
                            "Заголовок_ID"),
            TransCol   = Table.TransformColumnTypes(CombineCol,{{"Заголовок_ID", Int64.Type}})
        in  TransCol,

   
/* ШАГ 1 - ПОДКЛЮЧЕНИЕ к источнику данных -------------------------------------------------------------------*/

    from = Table.SelectRows(ВсеФормыВпапке2, each [Report_ID] = "pril" ) [[DataSet],[Company],[Date],[Content]],
   
/* ШАГ 2 - ЧИСТКА ДАННЫХ в подтаблицах -------------------------------------------------------*/
   
    // список колонок итоговой таблицы
    lstNewColNаmes = {"Date", "Статья"} & {"Строка_ID", "1", "2", "3", "4", "5"},

	// общая чистка в подтаблице
    f=(x)=>
        let 
			// фильтруем колонки
			Select = Table.SelectColumns(x, {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7"} ), 

			// чистим текст во всей таблице (" " двойные пробелы, табуляцию,)
			// ClearData = Load("fxClear")(Select, " #(lf)#(00A0)", true, true),
			k=(x)=> try Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " ") otherwise x,
            ClearData = Table.TransformColumns(Select, {}, k),

			// список имен колонок подтаблицы
			ColNames = Table.ColumnNames(ClearData),

			// меняем тип данных второй колонки - "НомерСтроки" для дальнейшей фильтрации
			Transform1 = Table.TransformColumns(ClearData,{ 
					{ ColNames{1}, (x)=> try Number.From(x) otherwise null, Int64.Type},
                    { ColNames{3}, (x)=> try Number.From(x) otherwise null, Int64.Type}
				} ),
		
			// фильтруем строки в колонке ПО ИМЕНИ колонки, не по [полю]
			// Table.SelectRows(RenameColumns, each {ColNames{1}} <> null), не работает!
			f=(x)=> (x <> null) and (x > 0) and (x < 2000),
            
			SelectRow = Table.SelectRows (Transform1, 
					(row) => f (Record.Field (row, ColNames{1} ) ) 
				),

			// меняем названия колонок кроме первой (первая Date)
			RenameColumns = Table.RenameColumns(SelectRow, List.Zip({
					List.Buffer(ColNames), 
					List.Skip(lstNewColNаmes,1)
				}) )
    
		in RenameColumns,

    /* ШАГ 2.2 - Добавление очищенной подтаблицы */
	
	// добавляем подтаблицу
    addColumn3 = Table.AddColumn(from, "Content2", each f([Content])),

    // конец шага 2.1 
    RESULT_STEP_1 = addColumn3  [[DataSet],[Company],[Date],[Content2]],

/* ШАГ 3 - ТРАНСФОРМАЦИЯ --------------------------------------------------------------------------------------*/

    // Разворачиваем подтаблицы. Список столбцов без первой колонки "Date", которой внутри нет
    // надежней будет колонки определять сразу до разворота  ColNames = Table.ColumnNames(from2[Step2]{0})
    EXPAND = Table.ExpandTableColumn(RESULT_STEP_1,
        "Content2", 
		List.Skip(lstNewColNаmes, 1), 
        List.Skip(lstNewColNаmes, 1)
    ),
    
    // меняем тип колонок с трансформацией. Можно делать это и в подтаблицах, до разворота
    Transform = Table.TransformColumns(EXPAND,
		{ 
			{ lstNewColNаmes{0}, (x)=> Date.From(x), Date.Type},

            { lstNewColNаmes{1}, (x)=> x, type text},
            { lstNewColNаmes{2}, (x)=> x, Int64.Type},
            { lstNewColNаmes{3}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
            { lstNewColNаmes{4}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
            { lstNewColNаmes{5}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
            { lstNewColNаmes{6}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
            { lstNewColNаmes{7}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type}
		} 
	),
    Unpivot = Table.UnpivotOtherColumns(Transform, {"Date", "Статья", "Строка_ID"}, "Заголовок", "Значение"),

  
    // добавляем колонки
    AddedColumns = F[fxAddMultiColumns](Unpivot, 
        {
            { "Раздел_ID",  each if List.Contains ({1..36},      [Строка_ID]) = true then 1
                            else if List.Contains ({50..65},     [Строка_ID]) = true then 2
                            else if List.Contains ({70..70},     [Строка_ID]) = true then 3
                            else if List.Contains ({80..95},     [Строка_ID]) = true then 4
                            else if List.Contains ({140..171},   [Строка_ID]) = true then 5
                            else if List.Contains ({1010..1320}, [Строка_ID]) = true then 6
                            else if List.Contains ({1401..1411}, [Строка_ID]) = true then 7
                            else 9, Int64.Type},
            
            { "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья], type text},
            
            { "Заголовок_ID", each Number.From(Text.From([Раздел_ID]) & Text.From([Заголовок])), Int64.Type}
        }),
    fltr = Table.SelectRows(AddedColumns, each ([Заголовок] <> "DataSet" and [Заголовок] <> "Company")),

	// добавляем значения колонок с ID, если переключатель IsReportViewGroup = True
    Join = 
        let Join_Заголо  = Table.Join(fltr,         "Заголовок_ID", Спр_Заголовки,   "Заголовок_ID"),
            Join_Раздел  = Table.Join(Join_Заголо,  "Раздел_ID",    Спр_Разделы,     "Раздел_ID")
        in  Join_Заголо,
            
    // ПОСЛЕДНИЙ ШАГ
    fxAddEndColumns = 
        let Add = F[fxAddMultiColumns](Join, {
                    {"Form_ID", each "pril0512", type text},
                    {"DataSet", each 2, Int64.Type},
                    {"Организация", each "Гомель", type text}
                }),
    
            Reorder = F[fxReorderSubsetOfColumns]
                    (Add, {"Form_ID", "DataSet", "Организация","Date"}, 0)
        in Reorder,

    RESULT= fxAddEndColumns
in
    RESULT