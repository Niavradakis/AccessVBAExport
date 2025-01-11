'1   start______________________________________________________________________________________________________________
'Some notes about Issuable Documents, the fields of IssuableDocumentsT and how they help to construct reports

'Field Main_Product_Entity_Debited(For_Product_Documents) is for documents that are issued in commercial sector and have product details.
'Field notes in the table design view,have the following info ""YES" MEANS THAT COMPANY BUYS OR TRANSFERS IN PRODUCTS. "NO" MEANS THAT COMPANY SELLS OR TRANSFERS OUT PRODUCTS.
'WHEN MAIN ENTITY IS DEBITED, THIS APPLIES FOR PRODUCT BUT ALSO FOR THE FINANCIAL TRANSLATION OF THE PRODUCTDETAILS RECORD. SAME IF CREDITED..."_
'So, to have a wider view of the above compressed notes, please read the following:
'We use this field in reporting, to extract two main informations, in coperation with the field "Affects Inventory" of the same table and in coperation with
'IssuedDocuments's Intention Type - its foreign key (Affects Ownership field).
'In general, company needs to extract the following informations:
'a) which of her entities warehouses, production lines, employees etc have products in their hands and how many
'b) Which products and of what quantity has the company under its possession(ownership) (no matter if they are in hands of any of company's entity or are being owed from or to any 3d party (non company's transactor)

'So, as we can understand, it is not rare to find that there are in hands of company's transactors products that are not in company's possesion(ownership) (this can occur if a supplier sends us products
'just with a transport document, without invoicing them). This means that possesion (ownership) of the goods is still on supplier but the products are in our hands.
'Of course the opposite can occur, for example supplier may have invoiced us for goods that he has not send to us yet and they are not in any of our company's transactors hands.
'So, we use the following pattern in reporting (taking into account that in table IssuableDocumentT, field Used_In_Commercial_Sector is true and field Has_Product_Details is true also) :

'Document affects inventory | Main_Product_Entity_Debited(For_Product_Documents) | intention type (Affects Ownership field)   |Ownership product result | Products In company's transactor hands result
'yes (ID = 1)                                  yes                                           yes                                              +                         +
'yes (ID = 1)                                  no                                            yes                                              -                         -
'no (ID = 2)                                   yes                                           yes                                              +                    not affected
'no (ID = 2)                                   no                                            yes                                              -                    not affected
'yes (ID = 1)                                  yes                                           no                                         not affected                    +
'yes (ID = 1)                                  no                                            no                                         not affected                    -
'no (ID = 2)                                   yes                                           no                                         not affected               not affected
'no (ID = 2)                                   no                                            no                                         not affected               not affected
'NO BUT...INFORMATION REASONS (ID = 3) |       yes/no                                        yes/no                                     not affected               not affected
 
'Field Main_Financial_Entity_Debited(Commercial_Sector) is for documents that are issued in commercial sector and have financial details, no matter if they have product details.
'So, this field regulates if the main financial transactor(or transactors if they are many, like bank, cash, credit card etc) is debited or credited. Probably the name of the field is not
'correct, and should be better "Other_Financial_Entity_Debited(Commercial_Sector)" as the case will be that other financial entity will always be one transactor, but it is ok, this is
'not a big problem, it works as it is, by letting us know what Other Financial entity does by informing us what main entity financial transactor does.
'We must keep in mind that this most of the time regulates documents (or part of them, like invoice with receipt) that do have money transactions.
'It does not regulate in any way how the product details records are being transformed to financial records (they are being transformed using the table "LinkProductsToCompanyFinancialTransactorsT"
'following the debit/credit pattern that comes from field "Main_Product_Entity_Debited(For_Product_Documents)" of IssuableDocumentsT
'So the financial transactor that will be selected from table LinkProductsToCompanyFinancialTransactorsT for each Product Details record, will be debited or credited similarly as
'the Main_Product_Entity did in this specific Product Details record.
'1       end____________________________________________________________________________________________________________________


'2  start ___________________________________________________________
'About Financial Details linked with Vat Transactor ID (as Attribute_ID 351) and Vat& (as Attribute_ID 16).
' The action of linking is executed on the after update event of control TransactorIDCbo, in which the code calls sub LinkFinancialTransactorRecordsWithVatTransactors located in the Module Public Functions.
' The sub LinkFinancialTransactorRecordsWithVatTransactors simply checks if there is already any record in the table LinkAttributeValueToEntitiesT for this specific Financial_Details_ID (linked with Aattribute_ID 351) and it does it by calling
' the function CheckIfEntityMayHaveMultipleAttributeValues (giving the value 351 for the attribute argument). Depending on the value returned, it inserts new record or updates existing record in LinkAttributeValueToEntitiesT with Attribute_ID 351
' The sub LinkFinancialTransactorRecordsWithVatTransactors does the same job one more time for the Vat% as Attribute_ID 16.
'2  end __________________________________________________________________

'3  start_____________________________________________________
' About  financial details updating Financial Transactors TotalDebit and TotalCredit fields in the TransactorsT.
' The action of updating the TotalDebit and TotalCredit fields takes place in the savebtn click sub of the form.
'If the form is the DocumentFinancialDetailsAddF then the code that updates the fields is located by the end of the sub, right before ExitProcedure.
'This happens in order to be the last action and to be executed when all records have been finalized and entered permanently in the IssuedDocumentFinancialDetailsT.
'3  end_______________________________________________________

'4  start_____________________________________________________
' About  Product details which have to do the following updates:
'1. updating Financial Transactors TotalDebit and TotalCredit fields in the TransactorsT for the link financial transactors with the product.
'2. updating Financial Transactors TotalDebit and TotalCredit fields in the TransactorsT for the link Vat transactors with the product.
'3. updating ProductInventoryBalanceT TotalDebit and TotalCredit fields for the specific product at the specific warehouse (main entity).
'4. updating ProductInventoryBalanceT TotalDebit and TotalCredit fields for the specific product at the specific warehouse (other entity).
' The above actions take place in the savebtn click sub of the form.
'If the form is the C4CProductTransactionF then the code that updates the fields is triggered in the savebtn click sub of the form, by the end of it, right before ExitProcedure.
'The savebtn click sub also checks if document has records in financial details (e.g. for money payment or receival) and do exactly what case 3 does (updates Financial Transactors TotalDebit and TotalCredit fields in the TransactorsT for all records from IssuedDocumentFinancialDetailsT))
'It calls UpdateTransactorsBalanceByProductDocumentDetailsRecordset which is located  in the ModulePublicFunctions. This sub needs an argument which is a recordset, constructed in the Savebtn_click sub,
'before it itterates all productDetails records removing Is_New flags or ProductDetailsBackupIDs.
'The recordset is named rstForUpdateTranBalanceAndInvBalanceByProductDoc and has all the necessary info to let the function UpdateTransactorsBalanceByProductDocumentRecordset
'do the above 4 jobs.
'This happens in order to be the last action and to be executed when all records have been finalized and entered permanently in the IssuedDocumentProductDetailsT.
'4  end_______________________________________________________

'5 start _____________________________________________________
'About updating document's financial transactors
'This action takes place like the above No4 action,in the savebtn click sub of the form.
'If the form is the C4CProductTransactionF then the code that updates the fields is triggered in the savebtn click sub of the form, by the end of it, right before ExitProcedure.
'It calls UpdateTransactorsBalanceByProductDocumentRecordset which is located  in the ModulePublicFunctions. This sub needs an argument which is a recordset, constructed in the Savebtn_click sub,
'before code removes from product document Is_New flags or ProductDocumentBackupID.
'The recordset is named rstForUpdateDocumentTranBalance and has all the necessary info to let the function UpdateTransactorsBalanceByProductDocumentRecordset
'do the job.
'This happens in order to be the last action and to be executed when all records have been finalized and entered permanently in the IssuedDocumentProductDetailsT.
'5 end _______________________________________________________

'6  start ___________________________________________________________
'About Product Details linked with FinancialTransactorsID and VatTransactorsID .
' The action of linking is executed on the after update event of control ProductIDCbo of C4CProductDocumentDetailsF, described as step 4 in the comments located in form's module
'6  end __________________________________________________________________