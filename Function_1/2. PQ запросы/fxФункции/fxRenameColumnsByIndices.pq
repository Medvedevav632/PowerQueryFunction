
let
    Источник = (Source as table, ColumnNamesNew as list, Indices as list) =>
let
    ColumnNamesOld = List.Transform( Indices, each Table.ColumnNames(Source){_} ), 
    ZippedList     = List.Zip( { ColumnNamesOld, ColumnNamesNew } ),
    RenamedColumns = Table.RenameColumns( Source, ZippedList )
in
    RenamedColumns
in
    Источник