IF OBJECT_ID('dbo.sp_ExportDataToCsv') IS NOT NULL
BEGIN
	DROP PROCEDURE dbo.sp_ExportDataToCsv
END

CREATE PROCEDURE [dbo].[sp_ExportDataToCsv] (@dbName NVARCHAR(100), @includeHeaders BIT, @filePath NVARCHAR(512), @tableName NVARCHAR(100), @reportName NVARCHAR(100), @delimiter NVARCHAR(4))
AS  
-- Created by Tyler T Updated on 12/12/2022
BEGIN
	DECLARE		@sqlCommand NVARCHAR(2000)
	DECLARE		@headerfileName NVARCHAR(100)
	DECLARE		@serverName NVARCHAR(100)
	--DECLARE		@delimiter NVARCHAR(4)
	DECLARE		@datafileName NVARCHAR(100)
	DECLARE		@removePrevious NVARCHAR(1000)
	DECLARE		@renameCommand NVARCHAR(1000)

	SET		@serverName = @@SERVERNAME

	IF @delimiter IS NULL SET @delimiter = '|'

	IF @includeHeaders = 1
	BEGIN

		SET		@headerfileName = 'headers.csv'

		----------------------------------------------------------
		-- creates the header file

		SET		@sqlCommand = 'bcp "SELECT Column_Name FROM '+@dbName+'.information_schema.columns where TABLE_NAME IN ('''+@tableName+''')" queryout "' + @filePath + '\' + @headerfileName +'" -S ' + @serverName + ' -c -t "' + @delimiter + '" -T -r "' + @delimiter + '"'

		Print @sqlCommand

		EXEC master..xp_cmdshell @sqlCommand

		DECLARE		@addNewLine NVARCHAR(1000)

		-- Adds a new line char to end of header file
		SET			@addNewLine = 'echo. >> "' + @filePath + '\' + @headerfileName + '"'

		Print @addNewLine

		EXEC master..xp_cmdshell @addNewLine

		--------------------------------------------------
		-- creates the data file

		SET		@datafileName = 'exported_data.csv'

		SET		@sqlCommand = 'bcp ' + @dbName +'..' + @tableName + ' out "' + @filePath + '\' + @datafileName + '" -S ' + @serverName + ' -c -t "' + @delimiter + '" -T'

		Print @sqlCommand

		EXEC master..xp_cmdshell @sqlCommand

		----------------------------------------------------
		-- merges the data and header file
		DECLARE		@appfileName	NVARCHAR(100)

		SET    @appfileName = @tableName + '_' + CONVERT(VARCHAR(30), GETDATE(), 23) + '.csv'	--112 works, but 23 has dashes in date > 01-01-2022 vs 20220101

		SET @sqlCommand = 'copy /b "' + @filePath + '\' + @headerfileName + '" + "' + @filePath + '\' + @datafileName + '" "' + @filePath + '\' + @appfileName + '"'

		Print @sqlCommand

		EXEC master..xp_cmdshell @sqlCommand

		----------------------------------------------------
		-- deletes the data & header files
		DECLARE		@deleteCommand NVARCHAR(1000)

		SET		@deletecommand = 'del "' + @filePath + '\' + @datafileName + '"'

		Print @deleteCommand

		EXEC master..xp_cmdshell @deleteCommand

		SET		@deletecommand = 'del "' + @filePath + '\' + @headerfileName + '"'

		EXEC master..xp_cmdshell @deleteCommand

		------------------------------------------------------
		-- deletes the previous version of file

		SET @removePrevious = 'if exist "' + @filePath + '\' + @reportName + '" del "' + @filePath + '\' + @reportName + '"'

		Print @removePrevious

		EXEC master..xp_cmdshell @removePrevious

		------------------------------------------------------
		-- renames the files

		SET @renameCommand = 'rename "' + @filePath + '\' + @appfilename + '" ' + @reportName 

		Print @renameCommand

		EXEC master..xp_cmdshell @renameCommand


	END

	ELSE IF @includeHeaders = 0
	BEGIN
	

		SET @datafileName = 'exported_data.csv'

		SET @sqlCommand = 'bcp ' + @dbName +'..' + @tableName + ' out "' + @filePath + '\' + @datafileName + '" -S ' + @serverName + ' -c -t "' + @delimiter + '" -T'

		Print @sqlCommand

		EXEC master..xp_cmdshell @sqlCommand

		----------------------------------------------------
		-- deletes the previous version of file

		SET @removePrevious = 'if exist "' + @filePath + '\' + @reportName + '" del "' + @filePath + '\' + @reportName + '"'

		Print @removePrevious

		EXEC master..xp_cmdshell @removePrevious
		------------------------------------------------------
		-- renames the files

		SET @renameCommand = 'rename "' + @filePath + '\' + @datafileName + '" ' + @reportName 

		Print @renameCommand

		EXEC master..xp_cmdshell @renameCommand


	END

END
