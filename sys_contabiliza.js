// on error resume next em java script
//onerror = continua
function continua()
{ return true; }
// inclusão em pop
function sqlInsPop(sUrl)
{
  with(document.editSql)
  {
	var campos	= sColumn.value;
	var exc = window.open(sUrl+'&sColumn='+ campos,'_ins','width=400,height=200,top=0,left=0')
	exc.focus();
  }	  
}
// exclusão em popus
function sqlDropPop(idLinha,sUrl,sTabela)
{

  with(document.editSql)
  {
	var campos	= sColumn.value;
	
	if(confirm('Deseja excluir o registro '+ idLinha +' da tabela '+ sTabela +' ?')==true)
	{
		var exc = window.open(sUrl+'&sId='+ idLinha +'&sColumn='+ campos,'_exc','width=400,height=200,top=0,left=0')
		exc.focus();
	}
  }	  
}
// edição em popus
function sqlEditPop(idLinha,sUrl)
{
  with(document.editSql)
  {
	var campos	= sColumn.value;
	var edit = window.open(sUrl+'&sId='+ idLinha +'&sColumn='+ campos,'_edit','width=400,height=200,top=0,left=0')
	edit.focus()
  }	
}
// oculta / exibi os layers das Tabelas / Views / Procedures
function banco(intQ)
{
  for(i=0;i<4;i++)
  {
    if(i==intQ){document.getElementById('B'+i).style.display = 'block';}
    else{document.getElementById('B'+i).style.display = 'none';}
  }
}
// oculta / exibi os layers das colunas
function coluna(codId)
{
  var M = document.getElementById(codId).style.display;
  var S = 'none';
  if(M == 'none'){S='block';}
  document.getElementById(codId).style.display = S;
}