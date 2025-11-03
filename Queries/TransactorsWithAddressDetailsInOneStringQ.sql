SELECT TransactorsT.Transactor_ID, FetchTransactorAllAddressDetailsAsOneString([Transactor_ID]) AS FullAddressString
FROM TransactorsT;

