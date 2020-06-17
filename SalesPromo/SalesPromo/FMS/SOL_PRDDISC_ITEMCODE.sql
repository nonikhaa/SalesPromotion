SELECT 	T0."ItemCode"
		, T0."ItemName"
		, T0."OnHand"
		, T0."IsCommited"
		, T0."OnOrder"
		, T0."InvntryUom"
FROM 	OITM T0
INNER 	JOIN OITB T1 ON T0."ItmsGrpCod" = T1."ItmsGrpCod"
WHERE 	T1."ItmsGrpNam" = 'BISCUIT'