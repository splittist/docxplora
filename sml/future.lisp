;;;; future.lisp

(in-package #:sml)

(defparameter *dynamic-renamings*
  '(("\\bLET\\(" "_xlfn.LET(")
    ("\\bLAMBDA\\(" "_xlfn.LAMBDA(")
    ("\\bSINGLE\\(" "_xlfn.SINGLE(")
    ("\\bSORTBY\\(" "_xlfn.SORTBY(")
    ("\\bUNIQUE\\(" "_xlfn.UNIQUE(")
    ("\\bXMATCH\\(" "_xlfn.XMATCH(")
    ("\\bSORT\\(" "_xlfn._xlws.SORT(")
    ("\\bXLOOKUP\\(" "_xlfn.XLOOKUP(")
    ("\\bSEQUENCE\\(" "_xlfn.SEQUENCE(")
    ("\\bFILTER\\(" "_xlfn._xlws.FILTER(")
    ("\\bRANDARRAY\\(" "_xlfn.RANDARRAY(")
    ("\\bANCHORARRAY\\(" "_xlfn.ANCHORARRAY(")))

(defparameter *future-renamings*
  '(("\\bCOT\\(" "_xlfn.COT(")
    ("\\bCSC\\(" "_xlfn.CSC(")
    ("\\bIFS\\(" "_xlfn.IFS(")
    ("\\bPHI\\(" "_xlfn.PHI(")
    ("\\bRRI\\(" "_xlfn.RRI(")
    ("\\bSEC\\(" "_xlfn.SEC(")
    ("\\bXOR\\(" "_xlfn.XOR(")
    ("\\bACOT\\(" "_xlfn.ACOT(")
    ("\\bBASE\\(" "_xlfn.BASE(")
    ("\\bCOTH\\(" "_xlfn.COTH(")
    ("\\bCSCH\\(" "_xlfn.CSCH(")
    ("\\bDAYS\\(" "_xlfn.DAYS(")
    ("\\bIFNA\\(" "_xlfn.IFNA(")
    ("\\bSECH\\(" "_xlfn.SECH(")
    ("\\bACOTH\\(" "_xlfn.ACOTH(")
    ("\\bBITOR\\(" "_xlfn.BITOR(")
    ("\\bF.INV\\(" "_xlfn.F.INV(")
    ("\\bGAMMA\\(" "_xlfn.GAMMA(")
    ("\\bGAUSS\\(" "_xlfn.GAUSS(")
    ("\\bIMCOT\\(" "_xlfn.IMCOT(")
    ("\\bIMCSC\\(" "_xlfn.IMCSC(")
    ("\\bIMSEC\\(" "_xlfn.IMSEC(")
    ("\\bIMTAN\\(" "_xlfn.IMTAN(")
    ("\\bMUNIT\\(" "_xlfn.MUNIT(")
    ("\\bSHEET\\(" "_xlfn.SHEET(")
    ("\\bT.INV\\(" "_xlfn.T.INV(")
    ("\\bVAR.P\\(" "_xlfn.VAR.P(")
    ("\\bVAR.S\\(" "_xlfn.VAR.S(")
    ("\\bARABIC\\(" "_xlfn.ARABIC(")
    ("\\bBITAND\\(" "_xlfn.BITAND(")
    ("\\bBITXOR\\(" "_xlfn.BITXOR(")
    ("\\bCONCAT\\(" "_xlfn.CONCAT(")
    ("\\bF.DIST\\(" "_xlfn.F.DIST(")
    ("\\bF.TEST\\(" "_xlfn.F.TEST(")
    ("\\bIMCOSH\\(" "_xlfn.IMCOSH(")
    ("\\bIMCSCH\\(" "_xlfn.IMCSCH(")
    ("\\bIMSECH\\(" "_xlfn.IMSECH(")
    ("\\bIMSINH\\(" "_xlfn.IMSINH(")
    ("\\bMAXIFS\\(" "_xlfn.MAXIFS(")
    ("\\bMINIFS\\(" "_xlfn.MINIFS(")
    ("\\bSHEETS\\(" "_xlfn.SHEETS(")
    ("\\bSKEW.P\\(" "_xlfn.SKEW.P(")
    ("\\bSWITCH\\(" "_xlfn.SWITCH(")
    ("\\bT.DIST\\(" "_xlfn.T.DIST(")
    ("\\bT.TEST\\(" "_xlfn.T.TEST(")
    ("\\bZ.TEST\\(" "_xlfn.Z.TEST(")
    ("\\bCOMBINA\\(" "_xlfn.COMBINA(")
    ("\\bDECIMAL\\(" "_xlfn.DECIMAL(")
    ("\\bRANK.EQ\\(" "_xlfn.RANK.EQ(")
    ("\\bSTDEV.P\\(" "_xlfn.STDEV.P(")
    ("\\bSTDEV.S\\(" "_xlfn.STDEV.S(")
    ("\\bUNICHAR\\(" "_xlfn.UNICHAR(")
    ("\\bUNICODE\\(" "_xlfn.UNICODE(")
    ("\\bBETA.INV\\(" "_xlfn.BETA.INV(")
    ("\\bF.INV.RT\\(" "_xlfn.F.INV.RT(")
    ("\\bNORM.INV\\(" "_xlfn.NORM.INV(")
    ("\\bRANK.AVG\\(" "_xlfn.RANK.AVG(")
    ("\\bT.INV.2T\\(" "_xlfn.T.INV.2T(")
    ("\\bTEXTJOIN\\(" "_xlfn.TEXTJOIN(")
    ("\\bAGGREGATE\\(" "_xlfn.AGGREGATE(")
    ("\\bBETA.DIST\\(" "_xlfn.BETA.DIST(")
    ("\\bBINOM.INV\\(" "_xlfn.BINOM.INV(")
    ("\\bBITLSHIFT\\(" "_xlfn.BITLSHIFT(")
    ("\\bBITRSHIFT\\(" "_xlfn.BITRSHIFT(")
    ("\\bCHISQ.INV\\(" "_xlfn.CHISQ.INV(")
    ("\\bF.DIST.RT\\(" "_xlfn.F.DIST.RT(")
    ("\\bFILTERXML\\(" "_xlfn.FILTERXML(")
    ("\\bGAMMA.INV\\(" "_xlfn.GAMMA.INV(")
    ("\\bISFORMULA\\(" "_xlfn.ISFORMULA(")
    ("\\bMODE.MULT\\(" "_xlfn.MODE.MULT(")
    ("\\bMODE.SNGL\\(" "_xlfn.MODE.SNGL(")
    ("\\bNORM.DIST\\(" "_xlfn.NORM.DIST(")
    ("\\bPDURATION\\(" "_xlfn.PDURATION(")
    ("\\bT.DIST.2T\\(" "_xlfn.T.DIST.2T(")
    ("\\bT.DIST.RT\\(" "_xlfn.T.DIST.RT(")
    ("\\bBINOM.DIST\\(" "_xlfn.BINOM.DIST(")
    ("\\bCHISQ.DIST\\(" "_xlfn.CHISQ.DIST(")
    ("\\bCHISQ.TEST\\(" "_xlfn.CHISQ.TEST(")
    ("\\bEXPON.DIST\\(" "_xlfn.EXPON.DIST(")
    ("\\bFLOOR.MATH\\(" "_xlfn.FLOOR.MATH(")
    ("\\bGAMMA.DIST\\(" "_xlfn.GAMMA.DIST(")
    ("\\bISOWEEKNUM\\(" "_xlfn.ISOWEEKNUM(")
    ("\\bNORM.S.INV\\(" "_xlfn.NORM.S.INV(")
    ("\\bWEBSERVICE\\(" "_xlfn.WEBSERVICE(")
    ("\\bERF.PRECISE\\(" "_xlfn.ERF.PRECISE(")
    ("\\bFORMULATEXT\\(" "_xlfn.FORMULATEXT(")
    ("\\bLOGNORM.INV\\(" "_xlfn.LOGNORM.INV(")
    ("\\bNORM.S.DIST\\(" "_xlfn.NORM.S.DIST(")
    ("\\bNUMBERVALUE\\(" "_xlfn.NUMBERVALUE(")
    ("\\bQUERYSTRING\\(" "_xlfn.QUERYSTRING(")
    ("\\bCEILING.MATH\\(" "_xlfn.CEILING.MATH(")
    ("\\bCHISQ.INV.RT\\(" "_xlfn.CHISQ.INV.RT(")
    ("\\bCONFIDENCE.T\\(" "_xlfn.CONFIDENCE.T(")
    ("\\bCOVARIANCE.P\\(" "_xlfn.COVARIANCE.P(")
    ("\\bCOVARIANCE.S\\(" "_xlfn.COVARIANCE.S(")
    ("\\bERFC.PRECISE\\(" "_xlfn.ERFC.PRECISE(")
    ("\\bFORECAST.ETS\\(" "_xlfn.FORECAST.ETS(")
    ("\\bHYPGEOM.DIST\\(" "_xlfn.HYPGEOM.DIST(")
    ("\\bLOGNORM.DIST\\(" "_xlfn.LOGNORM.DIST(")
    ("\\bPERMUTATIONA\\(" "_xlfn.PERMUTATIONA(")
    ("\\bPOISSON.DIST\\(" "_xlfn.POISSON.DIST(")
    ("\\bQUARTILE.EXC\\(" "_xlfn.QUARTILE.EXC(")
    ("\\bQUARTILE.INC\\(" "_xlfn.QUARTILE.INC(")
    ("\\bWEIBULL.DIST\\(" "_xlfn.WEIBULL.DIST(")
    ("\\bCHISQ.DIST.RT\\(" "_xlfn.CHISQ.DIST.RT(")
    ("\\bFLOOR.PRECISE\\(" "_xlfn.FLOOR.PRECISE(")
    ("\\bNEGBINOM.DIST\\(" "_xlfn.NEGBINOM.DIST(")
    ("\\bPERCENTILE.EXC\\(" "_xlfn.PERCENTILE.EXC(")
    ("\\bPERCENTILE.INC\\(" "_xlfn.PERCENTILE.INC(")
    ("\\bCEILING.PRECISE\\(" "_xlfn.CEILING.PRECISE(")
    ("\\bCONFIDENCE.NORM\\(" "_xlfn.CONFIDENCE.NORM(")
    ("\\bFORECAST.LINEAR\\(" "_xlfn.FORECAST.LINEAR(")
    ("\\bGAMMALN.PRECISE\\(" "_xlfn.GAMMALN.PRECISE(")
    ("\\bPERCENTRANK.EXC\\(" "_xlfn.PERCENTRANK.EXC(")
    ("\\bPERCENTRANK.INC\\(" "_xlfn.PERCENTRANK.INC(")
    ("\\bBINOM.DIST.RANGE\\(" "_xlfn.BINOM.DIST.RANGE(")
    ("\\bFORECAST.ETS.STAT\\(" "_xlfn.FORECAST.ETS.STAT(")
    ("\\bFORECAST.ETS.CONFINT\\(" "_xlfn.FORECAST.ETS.CONFINT(")
    ("\\bFORECAST.ETS.SEASONALITY\\(" "_xlfn.FORECAST.ETS.SEASONALITY(")))
