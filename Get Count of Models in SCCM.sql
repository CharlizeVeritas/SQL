SELECT 
	Manufacturer0 as 'Manufacturer'
	,Model0 as 'Model'
	,COUNT(*) as 'Count'
FROM v_GS_COMPUTER_SYSTEM
WHERE Manufacturer0 not like 'VMware%'
	and Manufacturer0 like 'Dell%'
	or Manufacturer0 like 'LENOVO%'
	or Manufacturer0 like 'Microsoft%'
GROUP BY Manufacturer0, Model0
ORDER BY Count DESC
