/*  FnReorderSubsetOfColumns - Переупорядочиваение столбцов без привязки к названию других

    Функция получает:
        tbl - таблицу, 
        reorderedColumns - подмножество переупорядоченных имен столбцов в виде списка и 
        offset - индекс (номер колонки куда вставлять полонки по списку), начинающийся с нуля. 
        Затем выполняется переупорядочение очень надежным способом, без ссылок на имена других столбцов.

    Пример = FnReorderSubsetOfColumns(Буфер, {"Строка", "Код"}, 0)
*/
 
let
    Источник = (tbl as table, reorderedColumns as list, offset as number)
   as table =>
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
   )
in
    Источник