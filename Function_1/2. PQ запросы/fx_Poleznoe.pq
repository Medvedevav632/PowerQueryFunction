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

    
