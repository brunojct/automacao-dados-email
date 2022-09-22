# Situação Geral

Imagine que você trabalha em uma grande rede de lojas de roupa com 25 lojas espalhadas por todo o Brasil.

Todo dia, pela manhã, a equipe de análise de dados deve calcular os chamados **One Pages** e enviar para o gerente de cada loja o **One Page** específico da sua loja, bem como todas as informações usadas no cálculo dos indicadores.

Um **One Page** é um resumo muito simples e direto ao ponto, usado pela equipe de gerência de loja para saber os principais indicadores e permitir em 1 página tanto a comparação entre diferentes lojas, quanto quais indicadores aquela loja conseguiu cumprir naquele dia ou não.

Segue um exemplo enviado por e-mail:

<div align = 'center'>
 <img src = 'https://user-images.githubusercontent.com/114163919/191793943-d6250355-b0b2-4a31-b938-2928e971bb07.png' width = 500/>
 </div>



## Execução do Projeto Utilizando Python

Primeiramente, as bases de dados foram importadas e visualizadas com o auxílio da biblioteca **Pandas**. Com isso, eu pude fazer os primeiros tratamentos nos arquivos (como juntar diferentes tabelas em somente uma com todas as informações que eu precisava).

Foram então criadas, de forma automática, pastas no meu computador com o nome de cada loja onde posteriormente seriam salvos os arquivos com os dados de cada uma dessas lojas. Para isso, utilizei como ferramenta a bliblioteca **Pathlib**.

Por fim, os indicadores de cada loja foram calculados, os arquivos foram salvos nas pastas correspondentes e cada gerente recebeu um e-mail com os dados de sua loja.
O e-mail foi enviado utilizando a biblioteca **win32com.client** e seu conteúdo era formado por um **One Page** com as principais informações sobre os indicadores além do anexo com os dados de forma mais detalhada.
Também foi enviado um e-mail para a diretoria com o ranking diário e anual das melhores e piores lojas.


## Algumas fotos dos resultados

- __Pastas Criadas Automaticamente__

<div align = 'center'>
 <img src = 'https://user-images.githubusercontent.com/114163919/191796797-dab982d9-8e1a-43ce-a7ee-2b9e54a8dff1.png' width = 500/>
 </div>



- __E-mail enviado para cada gerente de forma automatizada__

<div align = 'center'>
 <img src = 'https://user-images.githubusercontent.com/114163919/191798710-a9aafce9-ee78-47b4-b8df-9440e37a5794.png' width = 500/>
 </div>



- __Conteúdo dos e-mails que cada gerente recebe__

<div align = 'center'>
 <img src = 'https://user-images.githubusercontent.com/114163919/191797633-cb32fa2a-9fad-415e-b1ea-dffabb4e6077.png' width = 500/>
 </div>



- __E-mail que a diretoria recebe com o ranking diário e anual__

<div align = 'center'>
 <img src = 'https://user-images.githubusercontent.com/114163919/191797905-30792eb6-9349-49ad-8450-6bb2dd70e18a.png' width = 500/>
 </div>




Com esse projeto, uma rotina diária que geraria certa complexidade e tempo de trabalho para calcular indicadores e apresentar para os principais responsáveis pode ser feita de forma bem rápida e fácil com apenas um clique. 
