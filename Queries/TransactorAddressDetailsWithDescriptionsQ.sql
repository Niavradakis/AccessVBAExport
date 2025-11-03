SELECT TransactorAddressDetailsT.*, CitiesT.City_Description, CountiesT.County_Description, CountriesT.Country_Description, ProvincesT.Province_Description
FROM (((TransactorAddressDetailsT LEFT JOIN CountriesT ON TransactorAddressDetailsT.Country_ID = CountriesT.Country_ID) LEFT JOIN CitiesT ON TransactorAddressDetailsT.City_ID = CitiesT.City_ID) LEFT JOIN CountiesT ON CitiesT.County_ID = CountiesT.County_ID) LEFT JOIN ProvincesT ON CountiesT.Province_ID = ProvincesT.Province_ID;

