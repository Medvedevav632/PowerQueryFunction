/*  Множественная замена текста по словарю на списках  ПОКА НЕ РАБОТАЕТ из за х
    Аргументы:
    dictionary  - список списков. Например: a = List.Buffer(Table.FromValue({ {"Что 1", "На что 1"}, {"Что 2", "На что 2"} }) [Value]) 
    x           - каждая строка. Потом так: b = Table.TransformColumns(tbl, {"Статья", func})

    Если в запросе использовать:
	МультиЗамена = 
        [ dict = List.Buffer(Table.FromValue({ {"Что 1", "На что 1"}, {"Что 2", "На что 2"} }) [Value]),
		  func = (x)=> List.Accumulate(dict, x, (state, current) => Text.Replace(state, current{0}, current{1} )),
		  b    = Table.TransformColumns(Фильтр, {"Статья", func})
		][b],  
    
    Пример использования:
    replace = Table.TransformColumns(Фильтр, {"Статья", func()})

*/
/*
let
    func = ( dictionary as list, x ) =>
        List.Accumulate(
            dictionary, 
            x, 
            (state, current) => Text.Replace( state, current{0}, current{1} )
        )
in  
    func 
    */
let
    func = ( dictionary as list, x ) =>
    let
        a = List.Buffer(Table.FromValue(dictionary) [Value]),
        b = List.Accumulate(
            a, 
            x, 
            (state, current) => Text.Replace( state, current{0}, current{1} )
        )
    in b
in  
    func