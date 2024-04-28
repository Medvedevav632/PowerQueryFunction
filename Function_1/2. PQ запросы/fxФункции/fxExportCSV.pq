// Экспорт таблицы в файл CSV (https://excel.city/2018/06/export-powerquery-query-to-csv/)
(   table_to_export as nullable table, 
    folder as text, 
    filename as text) as nullable table => 
  let

    _tbl = table_to_export, 

    p = [
      folder_normalized       = if Text.EndsWith(folder,"\") then folder else (folder & "\"),
      folder_web_format       = Text.Replace(folder_normalized, "\", "/"),
      filename_with_extension = if Text.EndsWith(Text.Lower(filename),".csv") then filename else filename & ".csv",
      full_filename           = folder_web_format & filename_with_extension,
      tblexport = ()=> if _tbl = null then #table(null,{}) else _tbl
    ],

    json = 
      Text.FromBinary
        (Json.FromValue(
          Table.ToRows( p[tblexport]() )
        ), 
      1251
      ),

    json2 =   
      Text.Replace(
        Text.FromBinary(
          Json.FromValue(
            Table.ToRows( p[tblexport]() )
          ), 
          1251
        ),",",";"
      ),

    // objFile.WriteLine(arr.map(row => row.join(';')).join('\n'));
    
    // objFile.WriteLine(arr.join('\n'));
    // objFile.WriteLine(arr.join(';'));
    // var semicolonArr = arr.map(row => row.split(',').join(';'));
    // objFile.WriteLine(semicolonArr.join('\n'));
    // objFile.WriteLine(arr.join('\n'));
    export_csv = Web.Page("
      <script>
        var fso=new ActiveXObject('Scripting.FileSystemObject');
        var objFile=fso.OpenTextFile('" & p[full_filename] & "',2 , true);
        var arr=" & json & ";
        var arrWithSemicolon = arr.map(row => row.map(cell => cell.replace(/,/g, ';')).join(';')).join('\n');
        objFile.WriteBlankLines(0);
        objFile.Close(); 
      </script>"
    ),
    
    to = 
        if export_csv=null
        then p[tblexport]() 
        else p[tblexport]() 
  in
    to