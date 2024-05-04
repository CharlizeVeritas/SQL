Use XXX -- SCCM SQL Database

Declare @SoftwareLike varchar (255) 
Declare @SoftwareNotLike varchar (255) 

Set @SoftwareLike = 'Adobe Acrobat%'
Set @SoftwareNotLike = '%Reader%'

Select distinct vrs.Name0 as ComputerName
	,ISNULL((REPLACE(right(MXG.TopConsoleUser0, charindex('\', MXG.TopConsoleUser0)),'\','')), (IsNULL((REPLACE(right(CDR.CurrentLogonUser, charindex('\', CDR.CurrentLogonUser)),'\','')), Vrs.User_Name0 ))) as 'UserName'
	,vga.DisplayName0 as 'Application'
	,vga.Version0 as 'Version'
	,'32' as 'Architecture'
	,vga.InstallDate0 as InstallDate
	,vga.Publisher0 as Publisher
	,vga.ProdID0 as ProductID
from v_R_System as Vrs 
	left join v_GS_ADD_REMOVE_PROGRAMS as Vga on Vga.ResourceID = Vrs.ResourceID
	LEFT JOIN v_GS_SYSTEM_CONSOLE_USAGE_MAXGROUP as MXG on MXG.ResourceID = Vrs.ResourceID
	LEFT JOIN v_CombinedDeviceResources as CDR on CDR.MachineID = Vrs.ResourceID
where Vga.DisplayName0 like @SoftwareLike
	and Vga.DisplayName0 not like @SoftwareNotLike
	and vga.Publisher0 is not null
UNION ALL
Select distinct vrs.Name0 as ComputerName
	,ISNULL((REPLACE(right(MXG.TopConsoleUser0, charindex('\', MXG.TopConsoleUser0)),'\','')), (IsNULL((REPLACE(right(CDR.CurrentLogonUser, charindex('\', CDR.CurrentLogonUser)),'\','')), Vrs.User_Name0 ))) as 'UserName'
	,vga.DisplayName0 as 'Application'
	,vga.Version0 as 'Version'
	,'64' as 'Architecture'
	,vga.InstallDate0 as InstallDate
	,vga.Publisher0 as Publisher
	,vga.ProdID0 as ProductID
from v_R_System as Vrs 
	left join v_GS_ADD_REMOVE_PROGRAMS_64 as Vga on Vga.ResourceID = Vrs.ResourceID
	LEFT JOIN v_GS_SYSTEM_CONSOLE_USAGE_MAXGROUP as MXG on MXG.ResourceID = Vrs.ResourceID
	LEFT JOIN v_CombinedDeviceResources as CDR on CDR.MachineID = Vrs.ResourceID
where Vga.DisplayName0 like @SoftwareLike
	and Vga.DisplayName0 not like @SoftwareNotLike
	--and vga.Publisher0 is not null
