PARAMETERS currenttimestamp DateTime, currentuserID Long, DiscountLogIDPar Long;
INSERT INTO DiscountLogsBackupT ( Discount_Logs_ID, Issued_Document_ID, Discount_OR_Offer_ID, [%Discount_Percentage], [#Discount_Value], Discount_For_Single_ProductDetailsID_Only, Is_Deleted, Discount_Insert_Timestamp, Discount_Insert_User_ID, Discount_Edit_Timestamp, Discount_Edit_User_ID )
SELECT DiscountLogsT.Discount_Logs_ID, DiscountLogsT.Issued_Document_ID, DiscountLogsT.Discount_OR_Offer_ID, DiscountLogsT.[%Discount_Percentage], DiscountLogsT.[#Discount_Value], DiscountLogsT.Discount_For_Single_ProductDetailsID_Only, DiscountLogsT.Is_Deleted, DiscountLogsT.Discount_Insert_Timestamp, DiscountLogsT.Discount_Insert_User_ID, [currenttimestamp] AS Discount_Edit_Timestamp, [currentuserID] AS Discount_Edit_User_ID
FROM DiscountLogsT
WHERE (((DiscountLogsT.Discount_Logs_ID)=[DiscountLogIDPar]));

