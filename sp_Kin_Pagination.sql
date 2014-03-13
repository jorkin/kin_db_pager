if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_Kin_Pagination]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_Kin_Pagination]
GO


SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE sp_Kin_Pagination 

@getSql NVARCHAR(4000) OUTPUT,
@iRecordCount INT OUTPUT,
@sTableName NVARCHAR(256),
@sPKey VARCHAR(128) = 'ID',
@sFields VARCHAR(1024) = '*',
@sCondition VARCHAR(2048) = '',
@sOrderByString VARCHAR(1024),
@sOrderByStringRev VARCHAR(1024),
@iPage INT = 1,
@iPageSize FLOAT = 20,
@iPage1Size FLOAT = 20,
@iPageOffSet INT =0,
@iMaxRecords INT = -1,
@iSpeed INT =1,
@bDistinct INT=0,
@bReturn INT=1

AS

SET NOCOUNT ON
PRINT '====sp_Kin_Pagination===='

DECLARE @iPageCount INT
DECLARE @iMaxRecordCount INT
DECLARE @iMaxPageCount INT
DECLARE @iLastPageCount INT
DECLARE @Return INT
DECLARE @iPageRecordCount INT

SET @sPKey = IsNull(@sPKey, 'ID')
SET @sFields = IsNull(@sFields, '*')
SET @sCondition = IsNull(@sCondition, '')
SET @sOrderByString = IsNull(@sOrderByString, '')
SET @sOrderByStringRev = IsNull(@sOrderByStringRev, '')

IF (@iRecordCount IS NULL)
BEGIN
	IF @bDistinct = 0
	BEGIN
		SET @getSql=N'SELECT @iRecordCount=COUNT(*) FROM ' + @sTableName + @sCondition
	END
	ELSE
	BEGIN
		SET @getSql=N'SELECT @iRecordCount=COUNT(*) FROM (SELECT DISTINCT ' + @sFields + ' FROM ' + @sTableName + @sCondition + ') KIN_PAGINATION_TABLE'
	END
	EXECUTE SP_EXECUTESQL @getSql,N'@iRecordCount int OUTPUT',@iRecordCount OUTPUT
	PRINT '@iRecordCount = ' + LTRIM(@iRecordCount)
END
PRINT '@iMaxRecords = ' + LTRIM(STR(@iMaxRecords))
SET @iMaxRecordCount = @iRecordCount
--PRINT '@iMaxPageCount = ' + LTRIM(@iMaxRecordCount)
IF (@iMaxRecords > 0 And @iMaxRecordCount > @iMaxRecords)
BEGIN
	SET @iMaxRecordCount = @iMaxRecords
END
PRINT '@iMaxRecordCount = ' + LTRIM(@iMaxRecordCount)
PRINT '@iPageSize = ' + LTRIM(@iPageSize)
PRINT '@iPage1Size = ' + LTRIM(@iPage1Size) 
IF (@iRecordCount = 0 Or @iPageSize = -1)
BEGIN
	SET @iPageCount = 1
	SET @iMaxPageCount = 1
END
ELSE
BEGIN
--	SET @iPageSize = 0
	SET @iPageCount = CEILING((@iRecordCount - @iPage1Size) / @iPageSize) + 1
	SET @iMaxPageCount = CEILING((@iMaxRecordCount - @iPage1Size) / @iPageSize) + 1
END
PRINT '@iPageCount = ' + LTRIM(@iPageCount)
PRINT '@iMaxPageCount = ' + LTRIM(@iMaxPageCount)
IF @iPage > @iMaxPageCount
BEGIN
	SET @iPage = @iMaxPageCount
END
SET @iLastPageCount = @iMaxRecordCount - (@iPageSize * (@iMaxPageCount -2)) - @iPage1Size

PRINT '@iPageCount = ' + LTRIM(@iPageCount)
PRINT '@iPage = ' + LTRIM(@iPage)
PRINT '@iLastPageCount = ' + LTRIM(@iLastPageCount)

IF @iPageSize > 0
BEGIN
	DECLARE @iStartPosition INT, @iEndPosition INT
	SET @iStartPosition = (@iPage -1) * @iPageSize - @iPageSize + @iPage1Size
	IF @iPageOffset < 0
	BEGIN
		SET @iStartPosition = @iStartPosition + @iPageOffSet
	END 
	IF @iStartPosition < 0
	BEGIN
		SET @iStartPosition = 0
	END
	SET @iEndPosition = @iPage * @iPageSize - @iPageSize + @iPage1Size
	IF @iPageOffset > 0
	BEGIN
		SET @iEndPosition = @iEndPosition + @iPageOffSet
	END
	IF @iEndPosition > @iMaxRecordCount
	BEGIN
		SET @iEndPosition = @iMaxRecordCount
	END
	SET @iPageRecordCount = @iPageSize
	IF @iPage = @iMaxPageCount
	BEGIN
		SET @iPageRecordCount = @iLastPageCount
	END
	IF @iPage = 1
	BEGIN
		SET @iPageRecordCount = @iPage1Size
	END
PRINT '@iStartPosition = ' + LTRIM(@iStartPosition)
PRINT '@iEndPosition = ' + LTRIM(@iEndPosition)
PRINT '@iPage = ' + LTRIM(@iPage)
PRINT '@iPageRecordCount = ' + LTRIM(@iPageRecordCount)
	IF @iSpeed = 1
	BEGIN
		IF @iPage = 1
		BEGIN
PRINT 'FIRSTPAGE'
			SET @getSql = 'SELECT TOP ' + LTRIM(@iPageRecordCount) + ' ' + @sFields + ' FROM ' + @sTableName + @sCondition + @sOrderByString
		END
		ELSE
		BEGIN
			IF @iPage = @iPageCount
			BEGIN
PRINT 'LASTPAGE'
				SET @getSql = 'SELECT * FROM (SELECT TOP ' + LTRIM(@iPageRecordCount) + ' ' + @sFields + ' FROM ' + @sTableName + @sCondition + @sOrderByStringRev + ') AS Kin_Pagination_TABLE1' + @sOrderByString
			END
			ELSE
			BEGIN
				IF @sOrderByString = ' ORDER BY ' + @sPKey + ' ASC'
				BEGIN
PRINT 'PKEY ASC'
					SET @getSql = 'SELECT TOP ' + LTRIM(@iPageRecordCount) + ' ' + @sFields + ' FROM ' + @sTableName + ' WHERE ' + @sPKey + ' > ( SELECT MAX(' + @sPKey + ') FROM ( SELECT TOP ' + LTRIM(@iStartPosition) + ' ' + @sPKey + ' FROM ' + @sTableName + @sCondition + @sOrderByString + ') AS Kin_Pagination_TABLE1 )' + @sOrderByString
				END
				ELSE
				BEGIN
					IF @sOrderByString = ' ORDER BY ' + @sPKey + ' DESC'
					BEGIN
PRINT 'PKEY DESC'
						SET @getSql = 'SELECT ' + @sFields + ' FROM (SELECT TOP ' + LTRIM(@iPageRecordCount) + ' ' + @sFields + ' FROM ' + @sTableName + ' WHERE ' + @sPKey + ' > ( SELECT MAX(' + @sPKey + ') FROM ( SELECT TOP ' + LTRIM(@iRecordCount - (@iStartPosition + @iPageRecordCount)) + ' ' + @sPKey + ' FROM ' + @sTableName + @sCondition + @sOrderByStringRev + ') AS Kin_Pagination_TABLE1 )' + @sOrderByStringRev + ') AS Kin_Pagination_TABLE2' + @sOrderByString
					END
					ELSE
					BEGIN
						IF @iPage * 2 > @iPageCount 
						BEGIN
PRINT 'PAGES 2/2'
							SET @getSql = 'SELECT ' + @sFields + ' FROM ' + @sTableName + ' WHERE ' + @sPKey + ' IN (SELECT TOP ' + LTRIM(@iPageRecordCount) + ' ' + @sPKey + ' FROM (SELECT TOP ' + LTRIM(@iRecordCount - @iStartPosition) + ' * FROM ' + @sTableName + @sCondition + @sOrderByStringRev + ') Kin_Pagination_TABLE1' + @sOrderByString + ')' + @sOrderByString
						END
						ELSE
						BEGIN
PRINT 'PAGES 1/2'
							SET @getSql = 'SELECT ' + @sFields + ' FROM ' + @sTableName + ' WHERE ' + @sPKey + ' IN (SELECT TOP ' + LTRIM(@iPageRecordCount) + ' ' + @sPKey + ' FROM (SELECT TOP ' + LTRIM(@iEndPosition) + ' * FROM ' + @sTableName + @sCondition + @sOrderByString + ') Kin_Pagination_TABLE1' + @sOrderByStringRev + ')' + @sOrderByString
						END
					END
				END
			END
		END
	END
	ELSE
	BEGIN
PRINT 'TOPIN'
		SET @getSql = 'SELECT ' + @sFields + ' FROM ' + @sTableName + ' WHERE ' + @sPKey + ' IN (SELECT TOP ' + LTRIM(@iEndPosition) + ' ' + @sPKey + ' FROM ' + @sTableName + @sCondition + @sOrderByString + ')'
		IF @iPage>1
		BEGIN
			SET @getSql = @getSql + ' AND ' + @sPKey + ' NOT IN (SELECT TOP ' + LTRIM(@iStartPosition) + ' ' + @sPKey + ' FROM ' + @sTableName + @sCondition + @sOrderByString + ')'
		END
		SET @getSql = @getSql + @sOrderByString
	END
END
ELSE
BEGIN
	SET @getSql = 'SELECT ' + @sFields + ' FROM ' + @sTableName + @sCondition + @sOrderByString
END
PRINT '@getSql = ' + @getSql

--PRINT @iMaxRecordCount /0 -------//* µ˜ ‘”√°£ *//-------

IF @bReturn = 1
BEGIN
	EXECUTE SP_EXECUTESQL @getSql
END

SET NOCOUNT OFF
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

