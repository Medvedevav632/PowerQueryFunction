// Справочник периодов (таблица)
let
    from   = Table.FromList({2017..2022}, Splitter.SplitByNothing()),
    Rename = Table.RenameColumns(from,{{"Column1", "Год"}}),
    tr     = Table.TransformColumnTypes(Rename,{{"Год", Int64.Type}})
in
    tr