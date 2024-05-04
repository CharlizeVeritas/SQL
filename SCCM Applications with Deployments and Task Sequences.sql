USE XXX -- SCCM SQL Database

DECLARE @ExcludeFolders table (folder nvarchar(255));
DECLARE @ExcludeApps table (name nvarchar(255));
DECLARE @AppName nvarchar(255);
DECLARE @Years int;
SET @AppName = '' -- Can be partial
SET @Years = 5

INSERT INTO @ExcludeApps (name) values 
('')

INSERT INTO @ExcludeFolders (folder) values 
('')

;WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules', 'http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest' as p1)

SELECT * FROM
(SELECT DISTINCT
	LPC.DisplayName as AppName
	,LDT.DisplayName as AppDeploymentType
	,LPC.SDMPackageDigest.value('(/p1:AppMgmtDigest/p1:Application/p1:DisplayInfo/p1:Info/p1:Publisher)[1]', 'nvarchar(max)') AS AppPublisher
	,LPC.SDMPackageDigest.value('(/p1:AppMgmtDigest/p1:Application/p1:DisplayInfo/p1:Info/p1:Version)[1]', 'nvarchar(max)') AS AppVersion
	,LDT.SDMPackageDigest.value('(/p1:AppMgmtDigest/p1:DeploymentType/p1:Installer/@Technology)[1]', 'nvarchar(max)') AS AppType
	,CASE (CAST(LDT.IsExpired  as varchar))
		WHEN 0 THEN 'TRUE'
		WHEN 1 THEN 'FALSE'
		END AS 'AppActive'
	,LPC.SDMPackageVersion as SccmAppVersion
	,LDT.AppModelName as AppID
	,CASE
		WHEN (LDT.SDMPackageDigest.value('(/p1:AppMgmtDigest/p1:DeploymentType/p1:Installer/p1:Contents/p1:Content/p1:Location)[1]', 'nvarchar(max)')) is null THEN 'FALSE'
		ELSE 'TRUE'
		END as HasContent
	,LDT.SDMPackageDigest.value('(/p1:AppMgmtDigest/p1:DeploymentType/p1:Installer/p1:Contents/p1:Content/p1:Location)[1]', 'nvarchar(max)') AS ContentLocation
	,LDT.SDMPackageDigest.value('(/p1:AppMgmtDigest/p1:DeploymentType/p1:Installer/p1:InstallAction/p1:Args/p1:Arg)[1]', 'nvarchar(max)') AS InstallCommandLine
	,LDT.SDMPackageDigest.value('(/p1:AppMgmtDigest/p1:DeploymentType/p1:Installer/p1:UninstallAction/p1:Args/p1:Arg)[1]', 'nvarchar(max)') AS UninstallCommandLine
	,LDT.SDMPackageDigest.value('(/p1:AppMgmtDigest/p1:DeploymentType/p1:Installer/p1:InstallAction/p1:Args/p1:Arg)[8]', 'nvarchar(max)') AS UserInteractionMode
	,LDT.SDMPackageDigest.value('(/p1:AppMgmtDigest/p1:DeploymentType/p1:Installer/p1:DetectAction/p1:Provider)[1]', 'nvarchar(max)') AS DetectAction
	,RIGHT(LDT.CreatedBy,CHARINDEX('\',REVERSE(LDT.CreatedBy))-1) as CreatedBy
	,LDT.DateCreated as DateCreated
	,CASE
		WHEN DATEDIFF(YEAR, LDT.DateCreated, GetDate()) >= @Years THEN 'FALSE'
		ELSE 'TRUE'
		END as CreatedWithin5Years
	,RIGHT(LDT.LastModifiedBy,CHARINDEX('\',REVERSE(LDT.LastModifiedBy))-1) as LastModifiedBy
	,LDT.DateLastModified as DateLastModified
	--,'\Software Library\Overview\Application Management\Applications' + REPLACE(SF.FolderPath,'/','\') as ConsolePath
	,CASE
		WHEN SF.FolderPath is null THEN '\Software Library\Overview\Application Management\Applications'
		ELSE '\Software Library\Overview\Application Management\Applications' + REPLACE(SF.FolderPath,'/','\')
		END as ConsolePath
	,CASE (CAST(LPC.IsSuperseded  as varchar))
		WHEN 0 THEN 'FALSE'
		WHEN 1 THEN 'TRUE'
		END AS 'AppSuperseded'
	,CASE
		WHEN (Select top 1 ToApplicationCIID from vSMS_AppRelation_Flat where FromApplicationCIID = LPC.CI_ID and FromDeploymentTypeCIID = LDT.CI_ID) is not null THEN 'TRUE'
		ELSE 'FALSE'
		END as 'HasDependancy'
	,(Select DisplayName from fn_ListLatestApplicationCIs(1033) where CI_ID = (Select top 1 ToApplicationCIID from vSMS_AppRelation_Flat where FromApplicationCIID = LPC.CI_ID and FromDeploymentTypeCIID = LDT.CI_ID))as 'DependancyApplication'
	,(Select DisplayName from fn_ListDeploymentTypeCIs(1033) where CI_ID = (Select top 1 ToDeploymentTypeCIID from vSMS_AppRelation_Flat where FromApplicationCIID = LPC.CI_ID and FromDeploymentTypeCIID = LDT.CI_ID)) as 'DependancyDeploymentType'
	,CASE
		WHEN (Select NumberOfDependedDTs from fn_ListLatestApplicationCIs(1033) where CI_ID  = (Select top 1 FromApplicationCIID from vSMS_AppRelation_Flat where ToApplicationCIID = LPC.CI_ID and ToDeploymentTypeCIID = LDT.CI_ID)) != 0 THEN 'TRUE'
		ELSE 'FALSE'
		END as 'IsDependancy'
	,CASE
		WHEN LPC.IsSuperseded = 0 and (Select top 1 FromApplicationCIID from vSMS_AppRelation_Flat where ToApplicationCIID = LPC.CI_ID) is not null THEN (Select DisplayName from fn_ListLatestApplicationCIs(1033) where CI_ID = (Select FromApplicationCIID from vSMS_AppRelation_Flat where ToApplicationCIID = LPC.CI_ID and ToDeploymentTypeCIID = LDT.CI_ID))
		ELSE null
		END as 'DependantApplication'
	,CASE
		WHEN LPC.IsSuperseded = 0 and (Select top 1 FromApplicationCIID from vSMS_AppRelation_Flat where ToApplicationCIID = LPC.CI_ID) is not null THEN (Select DisplayName from fn_ListDeploymentTypeCIs(1033) where CI_ID = (Select FromDeploymentTypeCIID from vSMS_AppRelation_Flat where ToApplicationCIID = LPC.CI_ID and ToDeploymentTypeCIID = LDT.CI_ID))
		ELSE null
		END as 'DependantDeploymentType'
	,CASE
		WHEN TS.Name is not null THEN 'TRUE'
		ELSE 'FALSE'
		END as InTaskSequence
	,TS.TS_ID as TaskSequenceID
	,TS.Name as TaskSequenceName
	,CASE
		WHEN Vaa.AssignedCI_UniqueID is not null THEN 'TRUE'
		WHEN DS.PackageID is not null THEN 'TRUE'
		WHEN (Select top 1 CI_ID from v_ConfigurationItems where ModelID = DS.ModelID) is not null THEN 'TRUE'
		ELSE 'FALSE'
		END AS DeploymentActive	
	,ISNULL(CIA.Assignment_UniqueID, DS.OfferID) as DeploymentID
	,CASE (CAST(DS.DeploymentIntent as varchar))
		WHEN 1 THEN 'Required'
		WHEN 2 THEN 'Available'
		ELSE CAST(DS.DeploymentIntent as varchar)
		END AS DeploymentType
	,DS.DeploymentTime as DeploymentTime
	,DS.ModificationTime as ModificationTime
	,DS.CollectionID as CollectionID
	,DS.CollectionName as CollectionName
	,CASE DS.CollectionType
		WHEN 0 THEN 'Other'
		WHEN 1 THEN 'User'
		WHEN 2 THEN 'Device'
		ELSE CAST(DS.CollectionType as varchar)
		END AS CollectionType
	,(Select top 1 MemberCount from v_Collections where SiteID = DS.CollectionID) as MemberCount
	,CASE
		WHEN Vaa.AssignedCI_UniqueID is not null or TS.Name is not null or DS.PackageID is not null or (Select top 1 CI_ID from v_ConfigurationItems where ModelID = DS.ModelID) is not null THEN 'TRUE'
		ELSE 'FALSE'
		END AS AppInUse	
FROM fn_ListLatestApplicationCIs(1033) as LPC 
	left join fn_ListDeploymentTypeCIs(1033) as LDT on LPC.ModelName = LDT.AppModelName
	left join vFolderMembers as FM on FM.InstanceKey = LPC.ModelName and FM.ObjectTypeName = 'SMS_ApplicationLatest'
	left join vSMS_Folders as SF on SF.ContainerNodeID = FM.ContainerNodeID
	left join v_ApplicationAssignment Vaa on Vaa.AssignedCI_UniqueID = LPC.CI_UniqueID
	left join v_DeploymentSummary as Ds on Ds.AssignmentID = Vaa.AssignmentID
	left join v_CIAssignment as CIA on Ds.AssignmentID = CIA.AssignmentID
	left join v_TaskSequenceAppReferencesInfo TSApp on LPC.ModelName=TSApp.RefAppModelName
	left join v_TaskSequencePackage TS on TS.PackageID=TSApp.PackageID
where LDT.CIType_ID = 21 AND LDT.IsLatest = 1
	and LPC.DisplayName not in (Select Name from @ExcludeApps)
	and LPC.DisplayName like '%' + @AppName + '%') As MainTable
Where MainTable.ConsolePath not in (Select folder from @ExcludeFolders)
ORDER By MainTable.AppName
