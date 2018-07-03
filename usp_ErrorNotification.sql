USE [master]
GO

IF EXISTS(SELECT 1 FROM sys.procedures WHERE Name = 'usp_ErrorNotification')
BEGIN
	DROP PROCEDURE [dbo].[usp_ErrorNotification]
END
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[usp_ErrorNotification]
    (
     @job_id        UNIQUEIDENTIFIER
    ,@run_date        int
    ,@run_time        int
    ,@Recipients    nVARCHAR(500)
    ,@ProfileName    VARCHAR(50)
    )
    AS
    BEGIN

	-- =======================================================================================================================================
	-- Author	:	Sai Kiran Mirdoddi
	-- Contact	:	
	-- Create date	:	07/03/2018
	-- Description	:	HTML formatted mail to send SQL Agent job failure alerts
	--			Currently supports errors from SQL Steps
	--			Exception/Error messages from SSIS Packages deployed to SSISDB Catalog
	--			Sproc uses Job Token's as parameters. These will work when executed from the context of SQL Agent job step.
	--			Read more at <blog link>
	--			Database Mail profile should be created and Profile Name should be replaced in the parameters
	-- Usage
	/* 
		-- Query should be placed in the last step for the job. 
		-- Step should be reached upon failure of preceding steps
		-- This step should fail the job upon its success.

		    EXEC [dbo].[usp_ErrorNotification]
			@job_id = $(ESCAPE_SQUOTE(JOBID))
			,@run_date = $(ESCAPE_NONE(STRTDT))
			,@run_time = $(ESCAPE_NONE(STRTTM))
			,@recipients = 'alerts@gmail.com'
			,@ProfileName = 'PROFNAME'
    */
	-- =======================================================================================================================================
	
    DECLARE @EmailSubject    VARCHAR(200) = ''
    DECLARE @HtmlContent    NVARCHAR(MAX) = ''
    DECLARE @Body            VARCHAR(max) = ''
    DECLARE @footer            VARCHAR(max) = ''
    DECLARE @Content        VARCHAR(max) = ''
    DECLARE @JobName        VARCHAR(max) = ''
    DECLARE @TZ varchar(6)

	SELECT TOP 1 @TZ = DATENAME(tz,created_time) FROM SSISDB.internal.folders
	
	SET @footer            =    '</table></body></html>'

    SELECT @JobName = NAME FROM msdb.dbo.sysjobs WHERE job_id = @job_id
	
    -- To get HostName from ServerName
	SET @EmailSubject = SUBSTRING(@@SERVERNAME, 1, CASE WHEN CHARINDEX('\', @@SERVERNAME) = 0 THEN LEN(@@SERVERNAME) + 1 ELSE CHARINDEX('\', @@SERVERNAME) END - 1) + ' : "' + @JobNAme + '" Failed. '

    SET @body = '<html>
                    <body>
                        <table width="100%" border="0" align="CENTER" cellpadding="2" cellspacing="1">'

    DECLARE @joblog TABLE (
		  joblogID Int IDENTITY(1,1) PRIMARY KEY CLUSTERED
        , JobName NVARCHAR(255)
        , StepName NVARCHAR(255)
        , StepID INT
        , command NVARCHAR(max)
        , StepRunTime DATETIME
        , JobStartTime DATETIME
        , StepFailureMessage NVARCHAR(max)
        )

    INSERT INTO @jobLog
        (
          JobName
        , StepName
        , StepID
        , command
        , StepFailureMessage
        , StepRunTime
        , JobStartTime
        )

    SELECT
          JobName
        , StepName
        , StepID
        , Command
        , StepFailureMessage
        , RunDateTime AS StepRunTime
        , MIN(A.RunDateTime) OVER(Partition by JobName)  AS JobStartTime
    FROM
    (
    SELECT    J.Name AS JobName
            , S.step_name AS StepName
            , S.Step_id AS StepID
            , LEFT(command,CHARINDEX('\"" /SERVER',command)) command
            , max(msdb.dbo.agent_datetime(run_date,run_time)) AS [RunDateTime]
            , (SELECT REPLACE(message, '. ', '. <br/>') FROM msdb.dbo.sysjobhistory T
				WHERE T.job_id = j.job_id AND T.step_id = S.step_id
				AND max(msdb.dbo.agent_datetime(h.run_date,h.run_time)) = msdb.dbo.agent_datetime(T.run_date,T.run_time)
			  ) AS StepFailureMessage
    FROM msdb.dbo.sysjobhistory h
    INNER JOIN msdb.dbo.sysjobs j
     ON h.job_id = j.job_id
    INNER JOIN msdb.dbo.sysjobsteps s
     ON  j.job_id = s.job_id
     AND h.step_id = s.step_id
    WHERE h.run_status = 0
     --AND subsystem = 'SSIS'
    AND h.job_id = @job_id
    AND h.run_date >= @run_date
    AND h.run_time >= @run_time
    GROUP BY
          J.Name
        , j.job_id
        , S.step_name
        , S.Step_id
        , command
    ) A
   
    SELECT  @Content = @Content + '
    <tr>
            <tr>
                <td nowrap style="width:70px";  valign="TOP" ><font face="verdana" size="2"><b>Job Name</b></font></td>
                <td align = "LEFT" valign="TOP"><font face="verdana" SIZE="2">' + ISNULL(CAST(JobName AS NVARCHAR(MAX)), '') + '</font></td>
            </tr>
            <tr>
                <td nowrap style="width:70px" valign="TOP" ><font face="verdana" size="2"><b>Package Path</b></font></td>
                <td align = "LEFT" valign="TOP"><font face="verdana" SIZE="2">' + ISNULL(CAST(PackagePath AS NVARCHAR(MAX)), '') + '</font></td>
            </tr>
            <tr>
                <td nowrap style="width:70px" valign="TOP" ><font face="verdana" size="2"><b>Job StartTime</b></font></td>
                <td align = "LEFT" valign="TOP"><font face="verdana" SIZE="2">' + ISNULL(FORMAT(JobStartTime, 'G'), '') + '</font></td>
            </tr>
            <tr>
                <td nowrap style="width:70px" valign="TOP" ><font face="verdana" size="2"><b>StepID</b></font></td>
                <td align = "LEFT" valign="TOP"><font face="verdana" SIZE="2">' + ISNULL(CAST(StepID AS NVARCHAR(MAX)), '') + '</font></td>
            </tr>
            <tr>
                <td nowrap style="width:70px" valign="TOP" ><font face="verdana" size="2"><b>Step Name</b></font></td>
                <td align = "LEFT" valign="TOP"><font face="verdana" SIZE="2">' + ISNULL(CAST(StepName AS NVARCHAR(MAX)), '') + '</font></td>
            </tr>
            <tr>
                <td nowrap style="width:70px" valign="TOP" ><font face="verdana" SIZE="2"><b>Step RunTime</b></font></td>
                <td align = "LEFT" valign="TOP"><font face="verdana" SIZE="2">' + ISNULL(FORMAT(StepRunTime,'G'), '') + '</font></td>
            </tr>
            <tr>
                <td nowrap style="width:70px" valign="TOP" ><font face="verdana" size="2"><b>Error Message</b></font></td>
                <td align = "LEFT" valign="TOP"><font face="verdana" SIZE="2">' + ISNULL(CAST(PackageErrorMessage AS NVARCHAR(MAX)), '') + '</font></td>
            </tr>
    </tr>
    <tr>
        <td colspan = "2"><hr></td>
    </tr>

    '
    FROM (
    SELECT
          J.JobName
        , J.StepName
        , J.StepID
        , J.Command
        , J.StepRunTime
        , J.JobStartTime
        , CASE WHEN ISNULL(E.folder_name,'') <> '' THEN 'SSISDB\' + ISNULL(E.folder_name,'') + '\' + ISNULL(E.project_name,'') + '\' + ISNULL(E.package_name,'') ELSE '' END AS PackagePath
        , CASE WHEN ISNULL(J.StepFailureMessage,'') <> '' THEN 'Step Failure Message: ' + ISNULL(J.StepFailureMessage,'') ELSE '' END + CASE WHEN ISNULL(M.PackageErrorMessage ,'') <> '' THEN '<br/> Package Failure Message: ' + REPLACE(ISNULL(M.PackageErrorMessage,''), '. ', '. <br/>') ELSE '' END  AS PackageErrorMessage
    FROM @JobLog J
    LEFT JOIN SSISDB.CATALOG.executions E
    ON J.command = '/ISSERVER "\"\SSISDB\' + E.folder_name + '\' + E.project_name + '\' + E.package_name + '\'
    AND e.start_time >= TODATETIMEOFFSET (J.JobStartTime, @TZ)
    OUTER APPLY
            (
            SELECT M.message AS Div
            FROM SSISDB.CATALOG.operations O
            INNER JOIN SSISDB.CATALOG.event_messages M ON M.operation_id = O.operation_id
            WHERE O.operation_id = E.execution_id AND M.message_type IN (120,130)
            ORDER BY event_message_id
            FOR XML PATH('')
            ) M(PackageErrorMessage)
    ) A   

    SELECT @HtmlContent = ISNULL(@body,'') + CHAR(10) + ISNULL(@Content,'') + CHAR(10) + ISNULL(@footer,'') + CHAR(10) + ''
   
    IF EXISTS (SELECT TOP 1 1 FROM @JobLog)
    BEGIN
        EXECUTE msdb..sp_send_dbmail
               @profile_name = @ProfileName
            ,  @recipients = @recipients
            ,  @subject =  @emailSubject
            ,  @body = @HtmlContent
            ,  @body_format = 'HTML'
            ,  @importance = 'HIGH'
        --;THROW 51000, 'Job Failed', 1;
    END
   
    END
GO


