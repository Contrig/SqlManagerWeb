select name,crdate,type 
from sysobjects 
where xtype in ('U','V','P') and 
name not like '%dt%' AND
name not like '%sys%' AND
name not like '%dt%' AND
name not like '%df%' 
order by 1

exec sp_spaceused
exec sp_spaceused ATENDENTES
exec sp_columns ATENDENTES 

