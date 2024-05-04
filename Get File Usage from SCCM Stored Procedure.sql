USE XXX
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		Charlize Veritas
-- Create date: 07/27/2022
-- Modified date: 06/27/2023
-- Description:	Returns EXEs Last Used During Time Frame Provided
-- =============================================
CREATE PROCEDURE [dbo].[SCCM_File_Last_Used] 
	-- Add the parameters for the stored procedure here
	@ComputerName AS VARCHAR(100) = '%',
	@Domain AS VARCHAR(100) = '',
	@ExeLike AS VARCHAR(100) = '%',
	@ExeNotLike AS VARCHAR(100) = 'ABCDEFGHIJKLMNOP',
	@CompanyName AS VARCHAR(100) = '%',
	@ProductName AS VARCHAR(100) = '%',
	@msiDisplayName AS VARCHAR(100) = '%',
	@msiPublisher AS VARCHAR(100) = '%',
	@msiVersion AS VARCHAR(100) = '%',
	@Model AS VARCHAR(100) = '%',
	@Serial AS VARCHAR(100) = '%',
	@Version AS VARCHAR(100) = '%',
	@Path AS VARCHAR(100) = '%',
	@Username AS VARCHAR(100) = '%',
	@Days AS INT = '90',
	@Servers AS VARCHAR(100) = '',
	@Virtual AS VARCHAR(100) = ''
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
DECLARE @ServerLike AS VARCHAR(100)
DECLARE @VMWareLike AS VARCHAR(100)
DECLARE @VirtualLike AS VARCHAR(100)
DECLARE @DomainLike AS VARCHAR(100)
DECLARE @msiDisplayNameLike AS VARCHAR(100)
DECLARE @msiPublisherLike AS VARCHAR(100)
DECLARE @ExeLikeQuery AS VARCHAR(100)
DECLARE @ExeNotLikeQuery AS VARCHAR(100)
DECLARE @DaysQuery AS INT

If @Servers = 'false' or @Servers = 'FALSE' or @Servers = '0'
BEGIN 
	SET @ServerLike = 'Server'
END
ELSE
BEGIN
	SET @ServerLike = '#####'
END

If @Virtual = 'false' or @Virtual = 'FALSE' or @Virtual = '0'
BEGIN 
	SET @VMWareLike = 'vmware'
	SET @VirtualLike = 'virtual'
END
ELSE
BEGIN  
	SET @VMWareLike = '#####'
	SET @VirtualLike = '#####'
END

If @Domain = '' or @Domain is null
BEGIN 
	SET @DomainLike = '%'
END
ELSE
BEGIN  
	SET @DomainLike = @Domain
END

If @msiDisplayName = '' or @msiDisplayName is null
BEGIN 
	SET @msiDisplayNameLike = '%'
END
ELSE
BEGIN  
	SET @msiDisplayNameLike = @msiDisplayName
END

If @msiPublisher = '' or @msiPublisher is null
BEGIN 
	SET @msiPublisherLike = '%'
END
ELSE
BEGIN  
	SET @msiPublisherLike = @msiPublisher
END

If @ExeLike = '' or @ExeLike is null
BEGIN 
	SET @ExeLikeQuery = '%'
END
ELSE
BEGIN  
	SET @ExeLikeQuery = @ExeLike
END

If @ExeNotLike = '' or @ExeNotLike is null
BEGIN 
	SET @ExeNotLikeQuery = 'ABCDEFGHIJK'
END
ELSE
BEGIN  
	SET @ExeNotLikeQuery = @ExeNotLike
END

If @Days = '' or @Days is null
BEGIN 
	SET @DaysQuery = 90
END
ELSE
BEGIN  
	SET @DaysQuery = @Days
END

SELECT DISTINCT  
	Sys.Name0 as 'ComputerName'
	,UPPER(SUBSTRING(Sys.Full_Domain_Name0,0,CHARINDEX('.',Sys.Full_Domain_Name0))) as 'DomainName'
	,CS.Model0 as 'Model'
	,CASE WHEN SE.SerialNumber0 = 'Default string' THEN 'N/A' WHEN SE.SerialNumber0 = 'To Be Filled By O.E.M.' THEN 'N/A' ELSE  SE.SerialNumber0 END as 'Serial'
	,CASE SE.ChassisTypes0
		WHEN 1 THEN 'Other'
		WHEN 2 THEN 'Unknown'
		WHEN 3 THEN 'Desktop'
		WHEN 4 THEN 'Low Profile Desktop'
		WHEN 5 THEN 'Pizza Box'
		WHEN 6 THEN 'Mini Tower'
		WHEN 7 THEN 'Tower'
		WHEN 8 THEN 'Portable'
		WHEN 9 THEN 'Laptop'
		WHEN 10 THEN 'Notebook'
		WHEN 11 THEN 'Hand Held'
		WHEN 12 THEN 'Docking Station'
		WHEN 13 THEN 'All in One'
		WHEN 14 THEN 'Sub Notebook'
		WHEN 15 THEN 'Space-Saving'
		WHEN 16 THEN 'Lunch Box'
		WHEN 17 THEN 'Main System Chassis'
		WHEN 18 THEN 'Expansion Chassis'
		WHEN 19 THEN 'SubChassis'
		WHEN 20 THEN 'Bus Expansion Chassis'
		WHEN 21 THEN 'Peripheral Chassis'
		WHEN 22 THEN 'Storage Chassis'
		WHEN 23 THEN 'Rack Mount Chassis'
		WHEN 24 THEN 'Sealed-Case PC'
		WHEN 30 THEN 'Tablet'
		WHEN 31 THEN 'Convertible'
		WHEN 32 THEN 'Detachable'
		else 'Unknown'
		End As 'Chassis'
	,Sys.operatingSystem0 as 'OperatingSystem'
	,Sys.BuildExt as 'OSBuild'
	,CASE WHEN Sys.BuildExt LIKE '%7601%' THEN 'WIN7'
		WHEN Sys.BuildExt LIKE '%9200%' THEN 'WIN8'
		WHEN Sys.BuildExt LIKE '%9600%' THEN 'WIN81'
		WHEN Sys.BuildExt LIKE '%10240%' THEN '1507'
		WHEN Sys.BuildExt LIKE '%10586%' THEN '1511'
		WHEN Sys.BuildExt LIKE '%14393%' THEN '1607'
		WHEN Sys.BuildExt LIKE '%15063%' THEN '1703'
		WHEN Sys.BuildExt LIKE '%16299%' THEN '1709'
		WHEN Sys.BuildExt LIKE '%17134%' THEN '1803'
		WHEN Sys.BuildExt LIKE '%17763%' THEN '1809'
		WHEN Sys.BuildExt LIKE '%18362%' THEN '1903'
		WHEN Sys.BuildExt LIKE '%18363%' THEN '1909'
		WHEN Sys.BuildExt LIKE '%19041%' THEN '2004'
		WHEN Sys.BuildExt LIKE '%19042%' THEN '20H2'
		WHEN Sys.BuildExt LIKE '%19043%' THEN '21H1'
		WHEN Sys.BuildExt LIKE '%19044%' THEN '21H2'
		WHEN Sys.BuildExt LIKE '%19045%' THEN '22H2'
		ELSE 'N/A'
		END AS 'OSVersion'
	,Usage.ExplorerFileName0 as 'FileName'
	,Usage.FileDescription0 as 'FileDescription'
	,Usage.CompanyName0 as 'FileCompanyName'
	,Usage.FileVersion0 as 'Version'
	,Usage.msiDisplayName0 as 'msiDisplayName'
	,Usage.msiPublisher0 as 'msiPublisher'
	,Usage.msiVersion0 as 'msiVersion'
	,Usage.ProductName0 as 'ProductName'
	,Usage.ProductVersion0 as 'ProductVersion'
	,Usage.FolderPath0 as 'Path'
	,RIGHT(Usage.LastUserName0,CHARINDEX('\',REVERSE(Usage.LastUserName0))-1) as 'LastUserName'
	,USR.Mail0 as 'EmailAddress'
	,Usage.LastUsedTime0 as 'LastUsedTime'
FROM [CM_ESM].[dbo].v_R_System As SYS
	left join [CM_ESM].[dbo].v_GS_COMPUTER_SYSTEM CS on CS.ResourceID = SYS.ResourceID
	left join [CM_ESM].[dbo].v_GS_SYSTEM_ENCLOSURE AS SE on SE.ResourceID = SYS.ResourceID
	left join [CM_ESM].[dbo].v_GS_CCM_RECENTLY_USED_APPS Usage on sys.ResourceID = Usage.ResourceID
	left join [CM_ESM].[dbo].V_R_User USR ON Usage.LastUserName0 = USR.Unique_User_Name0
WHERE Usage.ExplorerFileName0 like @ExeLikeQuery
	and Usage.ExplorerFileName0 not like @ExeNotLikeQuery
	and Usage.LastUsedTime0 >= (GetDate() - @DaysQuery)
	and SE.ChassisTypes0 != 12
	and SE.SerialNumber0 like '%' + @Serial + '%'
	and CS.Model0 like '%' + @Model + '%'
	and CS.Model0 not like '%' + @VMWareLike + '%'
	and CS.Model0 not like '%' + @VirtualLike + '%'
	and Sys.operatingSystem0 not like '%' + @ServerLike + '%'
	and SYS.Netbios_Name0 like '%' + @ComputerName + '%'
	and Usage.CompanyName0 like '%' + @CompanyName + '%'
	and SYS.Full_Domain_name0 like '%' + @Domain + '%'
	and Usage.ProductName0 like '%' + @ProductName + '%'
	and Usage.msiVersion0 like '%' + @msiVersion + '%' 
	and Usage.FileVersion0 like '%' + @Version + '%'
	and Usage.FolderPath0 like '%' + @Path + '%'
	and Usage.LastUserName0 like '%' + @Username + '%'
	and Usage.ProductName0 like '%' + @ProductName + '%'
	and Usage.msiPublisher0 like @msiPublisherLike
	and Usage.msiDisplayName0 like @msiDisplayNameLike
END
