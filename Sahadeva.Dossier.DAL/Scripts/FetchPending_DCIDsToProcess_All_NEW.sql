USE [ADF_SAHADEV]
GO
/****** Object:  StoredProcedure [dbo].[FetchPending_DCIDsToProcess_All_NEW]    Script Date: 27-09-2024 13:20:00 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create  PROC [dbo].[FetchPending_DCIDsToProcess_All_NEW]
AS
BEGIN	
select   cd.CDID, cdc.TemplateID,
    'D:\OneDrive - Adfactors PR Pvt Ltd\Downloads\WordAutomation' AS SampleTemplateFile,
      e.EntityName + '/' + 
    CONVERT(VARCHAR(6),cdc.Createddate, 112) + '/' + e.EntityName+'_'+replace(MDT.DossierType,' ','')+'_'+REPLACE(CONVERT(VARCHAR(10), GETDATE(), 104), '.', '')+'.docx' AS GeneratedFileLocation	  
	  FROM CoverageDossier AS cd WITH (NOLOCK)
	LEFT JOIN ClientDossierConfig AS cdc  WITH (NOLOCK) ON cdc.CDCID = cd.CDCID
	LEFT JOIN Entity AS E  WITH (NOLOCK) ON e.EntityID=cd.EntityID
	LEFT JOIN mstDossierType AS MDT  WITH (NOLOCK) ON MDT.DTID=cdc.DossierTypeID
	--left JOIN [110.173.181.218].Adfactorshrcob.dbo.[User] AS u WITH(NOLOCK) ON u.UserID=cdc.UserID
WHERE cdc.TemplateID NOT IN (2,3) 	and StatusID=5   --AND cdc.IsActive=1 
END