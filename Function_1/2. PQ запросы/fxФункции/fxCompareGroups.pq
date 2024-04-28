/*  - проверяет, если ли текстке символы из списка

Например: создаем нужный список (группа) ключей и проверяем вот таким образом и отбираем строки:
    = Table.SelectRows(Таблица, each fnCompareGroups(Ключи1, [Текст]) 
*/

let
    Источник = (keys as list, text as text) as logical =>
    let
        Source = List.AnyTrue(
                    List.Transform(
                        keys, 
                        ( key ) => Text.Contains(text, key, Comparer.OrdinalIgnoreCase)
                    )
                )
    in Source
in
    Источник