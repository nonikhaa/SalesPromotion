SELECT	"CardCode", "CardName"
FROM	OCRD
WHERE	"validFor" = 'Y'
		AND "GroupCode" = $[@SOL_PRDDISC_H.U_SOL_CSTGRPCODE]