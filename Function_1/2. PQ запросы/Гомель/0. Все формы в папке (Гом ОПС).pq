let

/* ШАГ 0 - Модуль загрузки функций Load ---------------------------------------------------------------------*/

	Load = (fnName as text, optional BasePath as text) as function =>
		let DefaultPath = "G:\=EXCEL PQ\PQ-function-load\",
			GitHubPath  = "https://raw.githubusercontent.com/hohlick/pquery/master/",
				
			BasePath = if (BasePath <> null) then BasePath else DefaultPath,
			Path = BasePath & (if Text.End(BasePath, 1) <> "\" then "\" else ""),
			File = Path & fnName & ".pq",
			
			Function = try Expression.Evaluate(Text.Replace(fnName, ".", "_"), #shared)
			otherwise try  Expression.Evaluate(Text.FromBinary(Binary.Buffer(File.Contents(File))), #shared)
			otherwise      Expression.Evaluate(Text.FromBinary(Binary.Buffer(Web.Contents(GitHubPath & fnName & ".pq"))), #shared)
		in	Function,
	
    
/* ШАГ 1 - ПОДКЛЮЧЕНИЕ к источнику данных -------------------------------------------------------------------*/
    
    // подключаемся к папке
    path = Value.Metadata(PathFolder)[List]{1},
    load = 
		let from = Folder.Files(path),
            fltr = Table.SelectRows(from, each [Extension] = ".xls" or [Extension] = ".xlsx" and [Attributes][Hidden] = false),
            tbl  = Table.SelectColumns(fltr,{"Name","Content"})
        in  tbl,

    // выбираем нужные листы в файле 
    selectSheet = Table.TransformColumns(load,
        {"Content", (x)=>
            let a = Excel.Workbook(x), // соджержимое файла, если лист в файле один, то (x)=>Excel.Workbook(x){1}[Data]}
                b = Table.SelectRows(a, each Text.Length([Name]) = 15), // лист с кодом организации в названии листа = 15 знаков
                c = b{0}[Data]
            in  c
        }),

    // добавили колонки с допинформацией о файле/отчете
    addColumns1 = 
        let f = (tbl as table, addcol as list) as table => List.Accumulate(addcol, tbl, (t, d) => Table.AddColumn(t, d{0}, d{1}, d{2}?)),
            AddedColumns = f(selectSheet, { // т.к в файле xls дата = 01.16, в xlsx = 01.12.2016, то возникает ошибка => обрабатываем ее через try/otherwise
                { "Date",           each try       Date.From("01.01.20" & Text.End([Content][Column1]{0},2))             // если файл xls
                                         otherwise Date.From([Content][Column1]{0} + #duration(30, 0, 0, 0)), type date}, // если файл xlsx, то добаваляем 30 дней
                { "ID_Предприятие", each Text.From(  [Content][Column1]{1} ), type text },
                { "ID_Форма",       each Text.Start( [Content][Column2]{0}, 4 ), type text}
            })
        in  AddedColumns,
    
    // добавляем колонки через объединение таблиц по ID_Форма (нет необходимости, убрать надо будет)
    addColumns2 =
        let Спр_ФормыОтчетности = #table(type table 
				[ID_Форма = text, Название = text, КраткоеНазвание = text, Примечание = text],{
					{"0112","Баланс","ББ",""}, 
					{"0212","Отчет и прибылях и убытках", "ОПУ",""},
					{"0312","Отчет о доходах, расходах и финансовых результатах","ОДРФР",""},
					{"0412","Отчет о движении денежных средств","ОДДС",""},
					{"0512","Приложение к бухгалтерскому балансу","ПББ",""},
                    {"0612","Отчет о состоянии расчетов","ОСР",""},
                    {"06R1","Отчет о состоянии расчетов","ОСР",""},
                    {"12f1","Отчет о финансовых результатах","ОФР_12ф",""},
                    {"1212","Отчет о финансовых результатах ежемесячная","ОФР_мес",""},
					{"3012","Отчет об изменении капитала","ОИК",""} }),
			Join   = Table.NestedJoin(addColumns1,"ID_Форма",Спр_ФормыОтчетности,"ID_Форма","tmp"),
			Expand = Table.ExpandTableColumn(Join,"tmp", {"Название"})
		in  Expand
in
    addColumns2
