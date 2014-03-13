DECLARE @getSql NVARCHAR(4000)
DECLARE @iLastPageCount INT
DECLARE @iMaxPageCount FLOAT
DECLARE @iMaxRecordCount FLOAT
DECLARE @iMaxRecords FLOAT
DECLARE @iPage INT
DECLARE @iSpeed INT
DECLARE @iPageCount INT
DECLARE @iPageSize INT
DECLARE @iRecordCount INT
DECLARE @sCondition VARCHAR(2048)
DECLARE @sFields VARCHAR(1024)
DECLARE @sOrderByString VARCHAR(1024)
DECLARE @sOrderByStringrEV VARCHAR(1024)
DECLARE @sPKey VARCHAR(128)
DECLARE @sTableName NVARCHAR(256)

DECLARE @iPage_OUTPUT INT
DECLARE @iPageCount_OUTPUT INT
DECLARE @iRecordCount_OUTPUT INT
DECLARE @iLastPageCount_OUTPUT INT 
DECLARE @iMaxPageCount_OUTPUT FLOAT 
DECLARE @iMaxRecordCount_OUTPUT FLOAT
DECLARE @iMaxRecords_OUTPUT FLOAT 

exec sp_Kin_Pagination
@GETSQL OUTPUT
,@iPage_OUTPUT OUTPUT
,@iPageCount_OUTPUT OUTPUT
,@iRecordCount_OUTPUT OUTPUT
,@iLastPageCount_OUTPUT OUTPUT
,@iMaxPageCount_OUTPUT OUTPUT
,@iMaxRecordCount_OUTPUT OUTPUT
,@iMaxRecords_OUTPUT OUTPUT

,@sPKey = '[Article_ID]'
,@sFields = '*'
,@sCondition = ''
,@sOrderByString = ' ORDER BY [ARTICLE_ID] DESC'
,@sOrderByStringRev = ' ORDER BY [ARTICLE_ID] ASC'
,@iPage = 1
,@iSpeed = 0
,@iMaxRecords = -1
,@bReturn = 1
,@sTableName = '[Kin_Article]'


PRINT '@iPage_OUTPUT = ' + LTRIM(@iPage_OUTPUT)
PRINT '@iMaxRecords_OUTPUT = ' + LTRIM(@iMaxRecords_OUTPUT)
PRINT '@iRecordCount_OUTPUT = ' + LTRIM(@iRecordCount_OUTPUT)
PRINT '@iMaxRecordCount_OUTPUT = ' + LTRIM(@iMaxRecordCount_OUTPUT)
PRINT '@iPageCount_OUTPUT = ' + LTRIM(@iPageCount_OUTPUT)
PRINT '@iMaxPageCount_OUTPUT = ' + LTRIM(@iMaxPageCount_OUTPUT)
PRINT '@iLastPageCount_OUTPUT = ' + LTRIM(@iLastPageCount_OUTPUT)
PRINT '@GETSQL = ' + LTRIM(@GETSQL)

--,@sOrderByString = ' ORDER BY ARTICLE_DATETIME ASC, ARTICLE_ID ASC'
--,@sOrderByStringRev = ' ORDER BY ARTICLE_DATETIME DESC, ARTICLE_ID DESC'