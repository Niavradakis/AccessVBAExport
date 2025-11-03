PARAMETERS Discount_Logs_IDPar Long, Product_Details_IDPar Long, Unit_Price_Before_This_DiscountPar IEEEDouble, Unit_Price_After_This_DiscountPar IEEEDouble, Is_DeletedPar Bit, DiscountLogsDetails_Backup_IDPar Long;
UPDATE DiscountLogsDetailsT SET DiscountLogsDetailsT.Discount_Logs_ID = [Discount_Logs_IDPar], DiscountLogsDetailsT.Product_Details_ID = [Product_Details_IDPar], DiscountLogsDetailsT.Unit_Price_Before_This_Discount = [Unit_Price_Before_This_DiscountPar], DiscountLogsDetailsT.Unit_Price_After_This_Discount = [Unit_Price_After_This_DiscountPar], DiscountLogsDetailsT.Is_Deleted = [Is_DeletedPar], DiscountLogsDetailsT.DiscountLogsDetails_Backup_ID = Null
WHERE (((DiscountLogsDetailsT.DiscountLogsDetails_Backup_ID)=[DiscountLogsDetails_Backup_IDPar]));

