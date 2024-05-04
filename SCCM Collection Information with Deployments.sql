USE XXX -- SCCM SQL Database

DECLARE @ExcludeFolders table (folder nvarchar(255));
DECLARE @ExcludeCollections table (name nvarchar(255));
DECLARE @CollectionName nvarchar(255);
DECLARE @FeatureType int;

SET @CollectionName = '' -- Can be partial

INSERT INTO @ExcludeCollections (name) values 
('')

INSERT INTO @ExcludeFolders (folder) values 
('\Assets and Compliance\Overview\Device Collections\xxxxx')

SELECT * FROM
(SELECT 
	COL.SiteID as CollectionID
	,COL.CollectionName as CollectionName
	,CASE COL.CollectionType
		WHEN 0 THEN 'Other'
		WHEN 1 THEN 'User'
		WHEN 2 THEN 'Device'
		ELSE 'Unknown' END AS CollectionType
	,CASE
		WHEN COL.MemberCount = 0 THEN 'FALSE'
		ELSE 'TRUE'
		END AS HasMembers
	,COL.MemberCount as MemberCount
	,COL.LimitToCollectionName
	,CASE COL.IsReferenceCollection
		WHEN 1 THEN 'TRUE'
		ELSE 'FALSE' 
		END AS IsReferenceCollection
	,COL.BeginDate as DateCreated
	,CASE 
		WHEN COL.Schedule like '%02000' THEN 'FALSE'
		WHEN COL.Schedule like '%0A000' THEN 'FALSE'
		WHEN COL.Schedule like '%14000' THEN 'FALSE'
		WHEN COL.Schedule like '%1E000' THEN 'FALSE'
		WHEN COL.Schedule like '%28000' THEN 'FALSE'
		WHEN COL.Schedule like '%32000' THEN 'FALSE'
		WHEN COL.Schedule like '%3C000' THEN 'FALSE'
		WHEN COL.Schedule like '%50000' THEN 'FALSE'
		WHEN COL.Schedule like '%5A000' THEN 'FALSE'
		WHEN COL.Schedule like '%00100' THEN 'FALSE'
		WHEN COL.Schedule like '%00200' THEN 'FALSE'
		WHEN COL.Schedule like '%00300' THEN 'FALSE'
		WHEN COL.Schedule like '%00400' THEN 'FALSE'
		WHEN COL.Schedule like '%00500' THEN 'FALSE'
		WHEN COL.Schedule like '%00600' THEN 'FALSE'
		WHEN COL.Schedule like '%00700' THEN 'FALSE'
		WHEN COL.Schedule like '%00B00' THEN 'FALSE'
		WHEN COL.Schedule like '%00C00' THEN 'FALSE'
		WHEN COL.Schedule like '%01000' THEN 'FALSE'
		WHEN COL.Schedule like '%00008' THEN 'TRUE'
		WHEN COL.Schedule like '%00010' THEN 'TRUE'
		WHEN COL.Schedule like '%00020' THEN 'TRUE'
		WHEN COL.Schedule like '%00028' THEN 'TRUE'
		WHEN COL.Schedule like '%00038' THEN 'TRUE'
		WHEN COL.Schedule like '%D2000' THEN 'TRUE'
		WHEN COL.Schedule like '%92000' THEN 'TRUE'
		WHEN COL.Schedule like '%000080000' THEN 'TRUE'
		WHEN COL.Schedule like '%84400' THEN 'TRUE'
		Else 'Unknown'
		END AS SchedulePassFail	
	,CASE (CAST(COL.RefreshType as nvarchar(max)))
		WHEN 1 THEN 'Manual'
		WHEN 2 THEN 'Standard'
		ELSE 'Incremental'
		END as RefreshType
	,CASE 
		WHEN COL.Schedule like '%02000' THEN 'Every 1 min'
		WHEN COL.Schedule like '%0A000' THEN 'Every 5 mins'
		WHEN COL.Schedule like '%14000' THEN 'Every 10 mins'
		WHEN COL.Schedule like '%1E000' THEN 'Every 15 mins'
		WHEN COL.Schedule like '%28000' THEN 'Every 20 mins'
		WHEN COL.Schedule like '%32000' THEN 'Every 25 mins'
		WHEN COL.Schedule like '%3C000' THEN 'Every 30 mins'
		WHEN COL.Schedule like '%50000' THEN 'Every 40 mins'
		WHEN COL.Schedule like '%5A000' THEN 'Every 45 mins'
		WHEN COL.Schedule like '%00100' THEN 'Every 1 hour'
		WHEN COL.Schedule like '%00200' THEN 'Every 2 hours'
		WHEN COL.Schedule like '%00300' THEN 'Every 3 hours'
		WHEN COL.Schedule like '%00400' THEN 'Every 4 hours'
		WHEN COL.Schedule like '%00500' THEN 'Every 5 hours'
		WHEN COL.Schedule like '%00600' THEN 'Every 6 hours'
		WHEN COL.Schedule like '%00700' THEN 'Every 7 hours'
		WHEN COL.Schedule like '%00B00' THEN 'Every 11 hours'
		WHEN COL.Schedule like '%00C00' THEN 'Every 12 hours'
		WHEN COL.Schedule like '%01000' THEN 'Every 16 hours'
		WHEN COL.Schedule like '%00008' THEN 'Every 1 day'
		WHEN COL.Schedule like '%00010' THEN 'Every 2 days'
		WHEN COL.Schedule like '%00020' THEN 'Every 4 days'
		WHEN COL.Schedule like '%00028' THEN 'Every 5 days'
		WHEN COL.Schedule like '%00038' THEN 'Every 7 days'
		WHEN COL.Schedule like '%D2000' THEN 'Every Thursday'
		WHEN COL.Schedule like '%92000' THEN 'Every 1 week'
		WHEN COL.Schedule like '%000080000' THEN 'None'
		WHEN COL.Schedule like '%84400' THEN 'First Day of the Month'
		ELSE 'Unknown'
		END AS RefreshSchedule
	,COL.LastRefreshTime as LastRefreshTime
	,COL.LastIncrementalRefreshTime as LastIncrementalRefreshTime
	,CASE
		WHEN COL.CollectionType = 1 THEN '\Assets and Compliance\Overview\User Collections' + REPLACE(COL.ObjectPath,'/','\')
		WHEN COL.CollectionType = 2 THEN '\Assets and Compliance\Overview\Device Collections' + REPLACE(COL.ObjectPath,'/','\')
		ELSE REPLACE(COL.ObjectPath,'/','\')
		END AS ConsolePath
	,CASE
		WHEN Vaa.AssignedCI_UniqueID is not null THEN 'TRUE'
		WHEN DS.PackageID is not null THEN 'TRUE'
		WHEN (Select top 1 CI_ID from v_ConfigurationItems where ModelID = DS.ModelID) is not null THEN 'TRUE'
		ELSE 'FALSE'
		END AS DeploymentActive
	,ISNULL(CIA.Assignment_UniqueID, DS.OfferID) as DeploymentID
	,DS.DeploymentTime as DeploymentTime
	,DS.ModificationTime as ModificationTime
	,CASE
		WHEN Ds.FeatureType = 1 THEN VAA.ApplicationName
		WHEN Ds.FeatureType = 5 THEN UI.Title
		WHEN Ds.FeatureType = 6 THEN (Select top 1 DisplayName from fn_ListCI_ComplianceState(1033) where ModelId = DS.ModelID)
		ELSE DS.SoftwareName
		END AS ItemName
	,CASE (Ds.FeatureType)
		WHEN 1	THEN 'Application'
		WHEN 2	THEN 'Package'
		WHEN 3	THEN 'MobileProgram'
		WHEN 4	THEN 'Script'
		WHEN 5	THEN 'UpdateGroup'
		WHEN 6	THEN 'Baseline'
		WHEN 7	THEN 'TaskSequence'
		WHEN 8	THEN 'ContentDistribution'
		WHEN 9	THEN 'DistributionPointGroup'
		WHEN 10	THEN 'DistributionPointHealth'
		WHEN 11	THEN 'ConfigurationPolicy'
		WHEN 28	THEN 'AbstractConfigurationItem'
		END	as ItemType
	,CASE
		WHEN Ds.FeatureType = 1 THEN cast(Vaa.AssignedCI_UniqueID as varchar(255))
		WHEN Ds.FeatureType = 5 THEN cast((Select top 1 CI_UniqueID from v_UpdateInfo where ModelId = DS.ModelID) as varchar(255))
		WHEN Ds.FeatureType = 6 THEN cast((select top 1 CI_ID from fn_ListCI_ComplianceState(1033) where ModelId = DS.ModelID) as varchar(255))
		ELSE DS.PackageID
		END AS ItemID
FROM v_Collections as COL
left Join v_DeploymentSummary as Ds on DS.CollectionID = COL.SiteID
left join v_ApplicationAssignment Vaa on Ds.AssignmentID = Vaa.AssignmentID
left join v_CIAssignment as CIA on CIA.AssignmentID = Ds.AssignmentID
left join v_UpdateInfo as UI on UI.ModelId = DS.ModelID
WHERE COL.SiteID not like 'SMS%'
	and COL.CollectionName not in (Select Name from @ExcludeCollections)
	and COL.CollectionName like '%' + @CollectionName + '%'
	--and Vaa.AssignedCI_UniqueID = 'ScopeId_DB12B12E-92D0-4F27-8B63-D0F6AB1A82BC/Application_98770eeb-4027-4ab1-9a81-35b7216d5abb/3'
	--and DS.PackageID ='EST007AD'
	--and Ds.FeatureType = 6
) as MainTable
Where MainTable.ConsolePath not in (Select folder from @ExcludeFolders)
ORDER By MainTable.CollectionName
