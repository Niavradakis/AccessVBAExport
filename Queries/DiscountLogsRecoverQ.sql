PARAMETERS Issued_Document_IDPar Long, Discount_OR_Offer_IDPar Long, [%Discount_PercentagePar] IEEEDouble, [#Discount_ValuePar] IEEEDouble, Is_DeletedPar Bit, DiscountLogsBackupIDPar Long, Discount_For_Single_ProductDetailsID_OnlyPar Bit;
UPDATE DiscountLogsT SET DiscountLogsT.Issued_Document_ID = [Issued_Document_IDPar], DiscountLogsT.Discount_OR_Offer_ID = [Discount_OR_Offer_IDPar], DiscountLogsT.[%Discount_Percentage] = [%Discount_PercentagePar], DiscountLogsT.[#Discount_Value] = [#Discount_ValuePar], DiscountLogsT.Discount_For_Single_ProductDetailsID_Only = [Discount_For_Single_ProductDetailsID_OnlyPar], DiscountLogsT.Is_Deleted = [Is_DeletedPar], DiscountLogsT.DiscountLogs_Backup_ID = Null
WHERE (((DiscountLogsT.DiscountLogs_Backup_ID)=[DiscountLogsBackupIDPar]));

