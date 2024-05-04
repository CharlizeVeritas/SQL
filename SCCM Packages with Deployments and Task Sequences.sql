USE XXX -- SCCM SQL Database

DECLARE @ExcludeFolders table (folder nvarchar(255));
DECLARE @ExcludeApps table (name nvarchar(255));
DECLARE @PackName nvarchar(255);
DECLARE @Years int;
SET @PackName = '' -- Can be partial
SET @Years = 5

INSERT INTO @ExcludeApps (name) values 
('')

INSERT INTO @ExcludeFolders (folder) values 
('')

SELECT DISTINCT
	Packs.PkgID as PackageID
	,Packs.Name as PackageName
	,Prog.ProgramName as ProgramName
	,Packs.Manufacturer AS PackagePublisher
	,Packs.Version AS PackageVersion
	,Packs.StoredPkgVersion as SccmAppVersion
	,CASE
		WHEN Packs.Source is null THEN 'FALSE'
		ELSE 'TRUE'
		END as HasContent
	,Packs.Source AS ContentLocation
	,Prog.CommandLine AS CommandLine
	,Packs.SourceDate as SourceDate
	,CASE
		WHEN DATEDIFF(YEAR, Packs.SourceDate, GetDate()) >= @Years THEN 'FALSE'
		ELSE 'TRUE'
		END as CreatedWithin5Years
	,Packs.LastRefresh as LastRefresh
	,CASE
		WHEN SF.FolderPath is null THEN '\Software Library\Overview\Application Management\Packages'
		ELSE '\Software Library\Overview\Application Management\Packages' + REPLACE(SF.FolderPath,'/','\')
		END as ConsolePath
	,CASE
		WHEN TS.Name is not null or DS.PackageID is not null or (Select top 1 CI_ID from v_ConfigurationItems where ModelID = DS.ModelID) is not null THEN 'TRUE'
		ELSE 'FALSE'
		END AS PackageInUse	
	,CASE
		WHEN TS.Name is not null THEN 'TRUE'
		ELSE 'FALSE'
		END as InTaskSequence
	,TS.TS_ID as TaskSequenceID
	,TS.Name as TaskSequenceName
	,CASE
		WHEN DS.PackageID is not null THEN 'TRUE'
		WHEN (Select top 1 CI_ID from v_ConfigurationItems where ModelID = DS.ModelID) is not null THEN 'TRUE'
		ELSE 'FALSE'
		END AS DeploymentActive	
	,DS.OfferID as DeploymentID
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
INTO #MainTable
FROM vPackage as Packs
	left outer join v_Program as Prog on Prog.PackageID = Packs.PkgID --and Prog.ProgramName != '*'
	left outer join vFolderMembers as FM on FM.InstanceKey = Packs.PkgID and FM.ObjectTypeName = 'SMS_Package'
	left outer join vSMS_Folders as SF on SF.ContainerNodeID = FM.ContainerNodeID
	left outer join v_DeploymentSummary as Ds on Ds.PackageID = Packs.PkgID
	left outer join v_TaskSequencePackageReferences TSApp on TSApp.ObjectID = Prog.PackageID
	left outer join v_TaskSequencePackage TS on TS.PackageID=TSApp.PackageID
where Packs.Name not in (Select Name from @ExcludeApps)
	and Packs.Name like '%' + @PackName + '%'

SELECT * 
INTO #Asterisk
FROM #MainTable
WHERE ProgramName = '*'

SELECT *
INTO #NoAsterisk
FROM #MainTable
WHERE ProgramName != '*'

SELECT *
INTO #NoProgram
FROM #Asterisk
WHERE PackageID not in (Select PackageID from #NoAsterisk)

SELECT * FROM #NoAsterisk
UNION ALL
SELECT * FROM #NoProgram
Where ConsolePath not in (Select folder from @ExcludeFolders)
Order by PackageName

DROP TABLE #MainTable;
DROP TABLE #Asterisk;
DROP TABLE #NoAsterisk;
DROP TABLE #NoProgram;
