Option Compare Database

' 1. when i soft delete a document, i must update transactorsT total_debit and total_credit, and ProductInventoryBalanceF total_debit and total_credit
' 2. In DocumentSearchF when we do a delete, it marked as deleted only the document and not the transaction. It should mark as deleted also the transaction
' 3. In ProductDocumentF, when i change the otherentityfinancialtransactor, there should be also update of the VAT% in the productdocumentdetailsF in case the selected OtherEntity'sTransactionPoint has different VAT status than the previous selection
'    The same must happen and if only a simple change of OtherEntity'sTransactionPoint take place, without any change to otherentityfinancialtransactor

'4 In check form I must also check for commercial documents without transactors (product or financial- if they have financial movements)
'5 In check form I must look for unbalanced documents
'6 When I soft delete whole document (from the search form), there should be an adjustment accordingle to Transactor's financial balance and to the inventory balance also.
'8 In opposite commercial documents, the monetary records(if there are any monetary movement in the transation) are not following the negative sign as the product details do by having by defayult negative quantities. I must do exactly what I have done with the quantity at product details, which is after user tyoes a number I check if it is positive and then I transform it to negative.
'9 in C4CProductTransactionF when I delete one product detail and I abort the edit document by closing the form, it does not recover the non soft deleted record from back up
'10 in TransactorsAddF notesTbox it receives focus, it is enabled and it is yellow but I cannot write. In the edit form, I can write.
'11 in form link attribute vales to entitie, at filtered option for the attribute combobox, it does not take into account if an attribute has been used in the past for this entity type id but flagged as not to take into account as suggested attribute for filtered future suggestions. So, it brings it always as suggested, eventhough it has been flagged as not...