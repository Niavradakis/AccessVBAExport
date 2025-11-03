SELECT LinkProductsToPricelistsT.[Price_Net(Before_Discount)], LinkProductsToPricelistsT.Discount_Or_Offer_ID, LinkProductsToPricelistsT.Product_ID, LinkProductsToPricelistsT.Activation_Timestamp, LinkPricelistsToIntentionsT.Intention_ID, LinkProductsToPricelistsT.Product_Included_In_This_Pricelist_For_This_Period, LinkProductsToPricelistsT.Price_List_ID
FROM (PriceListsT INNER JOIN LinkPricelistsToIntentionsT ON PriceListsT.Price_List_ID = LinkPricelistsToIntentionsT.Pricelist_ID) INNER JOIN LinkProductsToPricelistsT ON PriceListsT.Price_List_ID = LinkProductsToPricelistsT.Price_List_ID
WHERE (((LinkProductsToPricelistsT.Product_ID)=1) AND ((LinkProductsToPricelistsT.Activation_Timestamp)<Now()) AND ((LinkPricelistsToIntentionsT.Intention_ID)=9) AND ((LinkProductsToPricelistsT.Product_Included_In_This_Pricelist_For_This_Period)=Yes))
ORDER BY LinkProductsToPricelistsT.Activation_Timestamp DESC;

