USE [DealershipProject]
GO
/****** Object:  StoredProcedure [dbo].[querycols]    Script Date: 9/3/2023 9:51:20 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROCEDURE [dbo].[querycols]
@tablecolumns nvarchar(max), 
@tablename nvarchar(max)
AS
/*
Objective: Take a stored procedure with the input of any given table’s name and column names. 
Using unpivot, create an index table with a column that consists of the column names of the original table. 
Loop through this index table with a while loop. Using dynamic SQL, perform an analysis for each column of the table. 
*/

--Connect the parameters of the stored procedure to the parameters within the stored procedure.
declare @tbname nvarchar(max) 
declare @tbcols nvarchar(max)
set @tbname = @tablename
set @tbcols = @tablecolumns

--Add a standardized datatype into the string of column names from the original table.
declare @tbdatatypecols nvarchar(max)
declare @dynamicsqlindxtb nvarchar(max)
select @tbdatatypecols =replace(@tbcols,']','] nvarchar(max)')

--Prepare the temp table that will hold the column names as rows before dynamic SQL starts.
drop table if exists #indx
create table #indx
(headers nvarchar(max), reads nvarchar(max))

--Create a table optimized for the unpivot, with the same column names as the original table except with a standardized datatype across columns.
set @dynamicsqlindxtb = 
'
	drop table if exists #unpivtprep
	create table #unpivtprep ('+@tbdatatypecols+') 
		insert into #unpivtprep 
			select top(1) *
				from '+@tbname+'

--Unpivot the table optimized for unpivot-ing, and insert that unpivot-ed table into the temp table made to hold the column names as rows.
	Insert into #indx
			select heads ,vals
				from 
				(select '+@tbcols+' from #unpivtprep
				 ) p unpivot (vals for heads in( '+@tbcols+'
							  )) as unpvt

--Add a column for looping to the temp table with the column names as rows
	alter table #indx
	add id int identity(1,1)
'
execute sp_executesql @dynamicsqlindxtb

--Set parameters for the loop.
declare @counter int, @maxid int, @headers nvarchar(max)

--Set parameters for the dynamic SQL.
declare @dynamicsqlcolquery nvarchar(max)
declare @params nvarchar(max) = '@typeselect int output, @nullscount int output'
declare @outcount int, @outcount1 int

--Loop.
select @counter = min(id)
       , @maxid =max(id) 
	from #indx 
		while 
			(
			@counter is not null and 
			@counter <= @maxid
			)
	BEGIN
		select @headers = headers
			from #indx 
				where id = @counter

--Analyze each column of the original table using the name of the column from the index temp table.
		select @dynamicsqlcolquery = 
		'
--Q1: How many types of row values are there in each column?
			set @typeselect = 
				(
				select count(*) as ['+@headers+'] 
					from( select ['+@headers+'] 
						from '+@tbname+' 
						group by ['+@headers+'] 
						) p
				)
--Q2: How many nulls are in each column? 
			set @nullscount = 
				(select count(*) 
					from '+@tbname+' 
					where ['+@headers+'] is null
				)
--Q3: …
--Q4: …
		'
--Print the analysis of the original table’s column. 
		execute sp_executesql @dynamicsqlcolquery
			, @params
			, @typeselect = @outcount output, @nullscount = @outcount1 output
		print @headers + ', types:' + convert(nvarchar(500),@outcount) + ', nulls:' + convert(nvarchar(500),@outcount1) + ','

--Increment the loop.
		set @counter = @counter +1
	END