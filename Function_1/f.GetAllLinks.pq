// возвращяет СПИСОК адресов со страницы источника

(	
	SourceURL 			as text, 
	AttributeDelimiter 	as text 
) =>

let
    // получаем HTML источник
    Source = Text.FromBinary( Web.Contents( SourceURL ) ),
    
    // функция поиска каждой ссылки
    GetLink = (Counter as number) =>
				let
					CurrentLink = 
						Text.BetweenDelimiters(
							Source, 
							"href=" & AttributeDelimiter, 	// href= + разделитель, например "
									  AttributeDelimiter,	// разделитель, например " 
							Counter
						)
                in	
					if   CurrentLink = ""
					then {} 
					else List.Combine(
							{ 
								{CurrentLink}, 
								@GetLink(Counter+1)
							} 
					     ),
    
	// вызов функции
    Output = GetLink(0)
in
    Output